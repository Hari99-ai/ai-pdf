from starlette.applications import Starlette
from starlette.responses import HTMLResponse, JSONResponse
from starlette.routing import Route


async def homepage(request):
    html = """
    <!doctype html>
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
            This repository contains a Streamlit app in <code>pdf_to_excel.py</code>.
            Vercel requires a Python entrypoint, so this page exists to satisfy the
            deployment build.
          </p>
          <p>
            To run the real UI, use Streamlit locally:
            <code>python -m streamlit run pdf_to_excel.py</code>
          </p>
          <p>
            If you want a Vercel deployment, the app needs to be converted to a
            Vercel-compatible web app instead of Streamlit.
          </p>
        </div>
      </body>
    </html>
    """
    return HTMLResponse(html)


async def health(request):
    return JSONResponse({"status": "ok"})


app = Starlette(
    debug=False,
    routes=[
        Route("/", homepage),
        Route("/health", health),
    ],
)
