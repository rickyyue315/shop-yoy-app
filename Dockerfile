# Shop YoY 分析系統 — Docker 部署（Railway 等雲端平台）
FROM python:3.12-slim

WORKDIR /app

# MALLOC_ARENA_MAX=2：減少 glibc 記憶體碎片（降低 Railway 記憶體用量／成本）
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    MALLOC_ARENA_MAX=2 \
    PORT=8080

COPY requirements.txt .
# pip 自動重試 + 外層最多重試 3 次：雲端 builder 偶爾斷線（BrokenPipeError）都唔會成個建置失敗
RUN for i in 1 2 3; do \
      pip install --no-cache-dir --retries 10 --timeout 120 -r requirements.txt && exit 0; \
      echo "pip install 失敗（第 $i/3 次），15 秒後重試…"; \
      sleep 15; \
    done; exit 1

COPY . .

EXPOSE 8080

# 必須 -w 1（單一 worker）：已上載資料存於記憶體，多 worker 會各自為政
# Railway 自動注入 PORT；--access-logfile - 將請求記錄寫入 Railway 部署日誌
CMD ["sh", "-c", "gunicorn -w 1 --threads 8 -t 120 --keep-alive 10 --graceful-timeout 30 --access-logfile - -b 0.0.0.0:${PORT:-8080} app:app"]
