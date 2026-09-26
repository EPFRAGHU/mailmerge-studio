FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=5000
EXPOSE 5000
# DATA_DIR lets the container persist data via a mounted volume
ENV DATA_DIR=/app/data
CMD gunicorn --bind 0.0.0.0:${PORT} --workers 2 --timeout 300 app:app
