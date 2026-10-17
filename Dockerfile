# Shop YoY 分析系統 — Docker 部署（Railway 等雲端平台）
FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

# 單一 worker + 多執行緒：已上載資料存於記憶體，需確保同一 worker 處理
CMD ["sh", "-c", "gunicorn -w 1 --threads 8 -t 120 -b 0.0.0.0:${PORT:-8080} app:app"]
