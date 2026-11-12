FROM python:3.11-slim

WORKDIR /app

COPY . /app

ENV PORT=8000
ENV API_TESTER_HOST=0.0.0.0

EXPOSE 8000

CMD ["python", "app.py"]

