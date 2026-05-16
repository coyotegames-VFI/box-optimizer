FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=8000

WORKDIR /app

COPY requirements.txt pyproject.toml README.md ./
COPY box_optimizer ./box_optimizer

RUN pip install --no-cache-dir -r requirements.txt

EXPOSE 8000

CMD ["sh", "-c", "uvicorn box_optimizer.api:app --host 0.0.0.0 --port ${PORT:-8000}"]
