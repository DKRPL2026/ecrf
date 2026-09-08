# eCRF Document Automation - Streamlit

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Render

For Streamlit Community Cloud, keep `packages.txt` in the repository root.
It installs LibreOffice Writer, which is required for PDF output to match the
Word template's layout, images, fonts, and checkbox appearance.

For Render, use a Web Service with Docker.

Set the service environment to Docker and use the repository `Dockerfile`.
Docker provides LibreOffice Writer for the most faithful DOCX-to-PDF layout.
Python-only deployments also have a basic ReportLab fallback for PDF output.

After changing `packages.txt` or the Dockerfile, trigger a fresh deploy so the
platform rebuilds the environment with LibreOffice. If Streamlit Cloud again
reports an expired Debian Release file, retry the rebuild or use the Docker
deployment, which installs LibreOffice from the Dockerfile.

If Render asks for a Start Command, use:

```bash
streamlit run app.py --server.address=0.0.0.0 --server.port=$PORT
```
