import json
import re
import requests
from typing import Dict, Any, List, Tuple
from pdf_extractor.logging_config import logger

CLASSIFICATION_PROMPT = """You are a financial classification agent. Your job is to classify the extracted PDF sections into exactly three logical groups for an Excel workbook.

Classification Criteria:
- Group 1 (Rates & Prices): Rates, Room Types, Seasons, Occupancy, Inventory, Prices, Tariffs, Rate Grids.
- Group 2 (Services & Amenities): Services, Amenities, Inclusions, Promotions, Offers, Taxes, Meal Plans, Child Policies, Transfers, Extra Charges, Surcharges.
- Group 3 (Policies & Cancellations): Cancellation Rules, Payment Terms, Billing, Booking Conditions, Deposit Rules, No Show, Amendments, Contact/General Info.

For each group, you must also generate a concise, professional worksheet title (under 30 characters) based on the contents classified into it. For example, if Group 1 contains 2026 rate tables, you might title it "2026 Room Rates" or "Rates & Inventory". Do not use generic names like "Group 1" or "Sheet 1".

Input Sections:
{sections_info}

Output Schema:
Return ONLY a valid JSON object matching the following schema. No markdown code block fences (e.g. do not wrap in ```json), no explanations, no preamble:
{{
  "group_1": {{
    "sheet_title": "dynamic_title_for_rates",
    "section_names": ["section_a", "section_b"]
  }},
  "group_2": {{
    "sheet_title": "dynamic_title_for_services",
    "section_names": ["section_c"]
  }},
  "group_3": {{
    "sheet_title": "dynamic_title_for_policies",
    "section_names": ["section_d"]
  }}
}}
"""

def get_sections_summary(sections: List[Dict[str, Any]]) -> str:
    """Create a textual summary of sections to help the LLM classify them."""
    summary_lines = []
    for sec in sections:
        name = sec.get("name", "")
        sec_type = sec.get("type", "")
        headers = sec.get("headers", [])
        
        # Take a tiny sample of rows
        rows_sample = ""
        if sec_type == "table" and sec.get("rows"):
            rows_sample = str(sec.get("rows", [])[:2])
        elif sec_type == "key_value" and sec.get("rows"):
            rows_sample = str(sec.get("rows", {}))
            
        summary_lines.append(
            f"- Name: {name}\n  Type: {sec_type}\n  Headers: {headers}\n  Sample: {rows_sample[:200]}"
        )
    return "\n".join(summary_lines)

def fallback_classify(sections: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Fallback rule-based classification based on section names and headers."""
    logger.info("Using fallback keyword-based classification...")
    
    group_1_names = ["rate", "price", "season", "grid", "inventory", "occupancy", "tariff", "dates"]
    group_3_names = ["cancel", "policy", "payment", "bill", "deposit", "no show", "amend", "term", "condition", "contact", "agent"]
    
    classification = {
        "group_1": {"sheet_title": "Rates & Inventory", "section_names": []},
        "group_2": {"sheet_title": "Services & Amenities", "section_names": []},
        "group_3": {"sheet_title": "Policies & Cancellation", "section_names": []}
    }
    
    for sec in sections:
        name = sec.get("name", "").lower()
        # Check matching keywords in name
        if any(kw in name for kw in group_1_names):
            classification["group_1"]["section_names"].append(sec["name"])
        elif any(kw in name for kw in group_3_names):
            classification["group_3"]["section_names"].append(sec["name"])
        else:
            classification["group_2"]["section_names"].append(sec["name"])
            
    return classification

def sanitize_sheet_title(name: str, default: str) -> str:
    """Sanitize sheet title for Excel compliance (remove []:*?/\\ and limit to 31 chars)."""
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", " ", name).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    if not cleaned:
        cleaned = default
    return cleaned[:31]

def classify_sections(
    sections: List[Dict[str, Any]],
    api_key: str,
    model: str
) -> Dict[str, Dict[str, Any]]:
    """Classify the sections into three groups and define dynamic worksheet titles."""
    if not sections:
        return {
            "group_1": {"sheet_title": "Rates & Inventory", "sections": []},
            "group_2": {"sheet_title": "Services & Amenities", "sections": []},
            "group_3": {"sheet_title": "Policies & Cancellation", "sections": []}
        }

    sections_by_name = {sec["name"]: sec for sec in sections}
    
    # Try LLM classification first
    try:
        sections_info = get_sections_summary(sections)
        prompt = CLASSIFICATION_PROMPT.format(sections_info=sections_info)
        
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers=headers,
            json={
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.0,
                "max_tokens": 2048
            },
            timeout=90
        )
        response.raise_for_status()
        content = response.json()["choices"][0]["message"]["content"].strip()
        
        # Clean potential markdown fences
        if content.startswith("```"):
            content = "\n".join(line for line in content.splitlines() if not line.strip().startswith("```")).strip()
            
        classification_data = json.loads(content)
    except Exception as e:
        logger.warning(f"LLM classification failed: {e}. Falling back to rule-based.")
        classification_data = fallback_classify(sections)

    # Reconstruct into structured groups
    groups = {
        "group_1": {"sheet_title": "Rates & Inventory", "sections": []},
        "group_2": {"sheet_title": "Services & Amenities", "sections": []},
        "group_3": {"sheet_title": "Policies & Cancellation", "sections": []}
    }
    
    # Match the classified section names back to full section objects
    for group_key in ["group_1", "group_2", "group_3"]:
        llm_group_data = classification_data.get(group_key, {})
        title_suggestion = llm_group_data.get("sheet_title", groups[group_key]["sheet_title"])
        groups[group_key]["sheet_title"] = sanitize_sheet_title(title_suggestion, groups[group_key]["sheet_title"])
        
        assigned_names = llm_group_data.get("section_names", [])
        for name in assigned_names:
            if name in sections_by_name:
                groups[group_key]["sections"].append(sections_by_name[name])
                
    # Check if any section was missed by the LLM classification, and if so, place it via fallback rules
    all_assigned = set()
    for g in groups.values():
        for s in g["sections"]:
            all_assigned.add(s["name"])
            
    for sec in sections:
        if sec["name"] not in all_assigned:
            logger.info(f"Section '{sec['name']}' was not assigned by LLM classification. Assigning via fallback rules.")
            name_lower = sec["name"].lower()
            if any(kw in name_lower for kw in ["rate", "price", "season", "grid", "inventory", "occupancy", "tariff", "dates"]):
                groups["group_1"]["sections"].append(sec)
                logger.info(f"Assigned '{sec['name']}' to Group 1")
            elif any(kw in name_lower for kw in ["cancel", "policy", "payment", "bill", "deposit", "no show", "amend", "term", "condition"]):
                groups["group_3"]["sections"].append(sec)
                logger.info(f"Assigned '{sec['name']}' to Group 3")
            else:
                groups["group_2"]["sections"].append(sec)
                logger.info(f"Assigned '{sec['name']}' to Group 2")
                
    return groups
