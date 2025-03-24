# Kies een basis-image
FROM python:3.12-slim

# Stel de werkmap in
WORKDIR /app

# Kopieer alle bestanden van je project naar de container
COPY . /app

# Installeer de benodigde Python-pakketten
RUN pip install --no-cache-dir -r requirements.txt

# Maak poort 8501 toegankelijk
EXPOSE 8501

# Start de Streamlit-applicatie
CMD ["streamlit", "run", "PricePilot.py", "--server.port=8501", "--server.address=0.0.0.0"]
