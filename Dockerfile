# Zeabur 用。不指定的話 buildpack 猜錯就很難查，直接寫死最省事。
FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY build_seats.py server.py index.html ./

# Zeabur 會自己給 PORT，這裡只是本機跑的預設值
ENV PORT=8080 \
    REFRESH_MINUTES=30 \
    PYTHONUNBUFFERED=1
EXPOSE 8080

CMD ["python", "server.py"]
