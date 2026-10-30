FROM python:3.12-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-calc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py audit_general.py audit_general_commentaires.py audit_pipeline.py \
     audit_commentaires_worker.py causes_format.py pretraitement.py pretraitement_grada.py \
     pretraitement_auto.py pretraitement_a_formater.py ./
COPY first_scripts/first_attempt_v2_prototype.py ./first_scripts/first_attempt_v2_prototype.py

EXPOSE 8501

ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
