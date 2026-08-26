# eCRF Document Automation - Streamlit

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Render

Use a Render Web Service with Docker.

Set the service environment to Docker and use the repository `Dockerfile`.
Docker provides LibreOffice Writer for the most faithful DOCX-to-PDF layout.
Python-only deployments also have a basic ReportLab fallback for PDF output.

After changing the Dockerfile, trigger a fresh deploy so Render rebuilds the
image instead of reusing an older image without LibreOffice.

If Render asks for a Start Command, use:

```bash
streamlit run app.py --server.address=0.0.0.0 --server.port=$PORT
```
