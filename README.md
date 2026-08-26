# eCRF Document Automation - Streamlit

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Render

Use a Render Web Service with Docker.

No custom apt-get Build Command is required.

The Dockerfile installs LibreOffice and starts Streamlit automatically.

If Render asks for a Start Command, use:

```bash
streamlit run app.py --server.address=0.0.0.0 --server.port=$PORT
```
