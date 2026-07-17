import os
from pathlib import Path
from pdf_extractor import process_pdf
from pdf_to_excel import load_env_value

def main():
    api_key = load_env_value("OPENROUTER_API_KEY")
    if not api_key:
        print("Error: OPENROUTER_API_KEY not found.")
        return

    test_dir = Path(r"C:\Users\hari9\Downloads\test")
    pdf_files = list(test_dir.glob("*.pdf"))
    print(f"Found {len(pdf_files)} PDF files in {test_dir}")

    for pdf_path in pdf_files:
        print(f"\nProcessing {pdf_path.name}...")
        try:
            pdf_bytes = pdf_path.read_bytes()
            excel_bytes, groups, warnings = process_pdf(
                pdf_bytes=pdf_bytes,
                api_key=api_key,
                model="openai/gpt-4o-mini"
            )
            out_path = pdf_path.with_suffix(".xlsx")
            out_path.write_bytes(excel_bytes)
            print(f"Successfully processed {pdf_path.name} -> {out_path.name}")
            print("Sheet titles:")
            for g_key, g_data in groups.items():
                print(f"  {g_key}: {g_data.get('sheet_title')} ({len(g_data.get('sections', []))} sections)")
            if warnings:
                print("Warnings:")
                for w in warnings:
                    print(f"  - {w}")
        except Exception as e:
            import traceback
            print(f"Error processing {pdf_path.name}: {e}")
            traceback.print_exc()

if __name__ == "__main__":
    main()
