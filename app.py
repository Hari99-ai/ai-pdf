from flask import Flask, jsonify, Response


app = Flask(__name__)


@app.get("/")
def homepage():
    html = """<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>PDF to Excel Extractor</title>
    <style>
      body {
        font-family: Arial, sans-serif;
        margin: 0;
        min-height: 100vh;
        display: grid;
        place-items: center;
        background: #f6f8fb;
        color: #1f2937;
      }
      .card {
        max-width: 720px;
        background: white;
        border: 1px solid #dbe2ea;
        border-radius: 16px;
        padding: 32px;
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.08);
      }
      h1 { margin-top: 0; }
      code {
        background: #eef2ff;
        padding: 2px 6px;
        border-radius: 6px;
      }
    </style>
  </head>
  <body>
    <div class="card">
      <h1>PDF to Excel Extractor</h1>
      <p>
        This deployment entrypoint exists for Vercel. The Streamlit app lives in
        <code>pdf_to_excel.py</code> and should be run locally.
      </p>
      <p>
        Local command:
        <code>python -m streamlit run pdf_to_excel.py</code>
      </p>
      <p>
        If you want the full PDF extractor on Vercel, the Streamlit UI must be
        rewritten into a normal web app.
      </p>
    </div>
  </body>
</html>"""
    return Response(html, mimetype="text/html")


@app.get("/health")
def health():
    return jsonify(status="ok")
