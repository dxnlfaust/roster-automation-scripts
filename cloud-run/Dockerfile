FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY roster_optimiser_month.py main.py ./

# Cloud Run sets PORT. One worker: OR-Tools is CPU-bound and already threaded,
# so extra workers would fight over the same cores. Timeout must exceed
# SOLVE_SECONDS plus parsing overhead.
ENV SOLVE_SECONDS=90 \
    SOLVE_WORKERS=8 \
    PORT=8080

CMD exec gunicorn --bind :$PORT --workers 1 --threads 4 --timeout 300 main:app
