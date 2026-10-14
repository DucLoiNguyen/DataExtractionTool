# syntax=docker/dockerfile:1
# Hai giai doan: "builder" cai thu vien thanh wheel, "runtime" chi chep wheel vao
# -> anh cuoi nho hon, khong co trinh bien dich/cache pip.
FROM python:3.13-slim AS builder
WORKDIR /build
COPY webapp/requirements.txt .
RUN pip wheel --no-cache-dir --wheel-dir /wheels -r requirements.txt

FROM python:3.13-slim AS runtime
ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1
# LibreOffice chi can de doc .doc cu; tat di (--build-arg WITH_LIBREOFFICE=0) neu khong can, anh nho di rat nhieu.
ARG WITH_LIBREOFFICE=0
RUN if [ "$WITH_LIBREOFFICE" = "1" ]; then \
      apt-get update && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-calc \
      && rm -rf /var/lib/apt/lists/*; \
    fi
COPY --from=builder /wheels /wheels
RUN pip install --no-cache-dir /wheels/* && rm -rf /wheels

# Chay bang user thuong, khong dung root
RUN useradd --create-home --uid 10001 app
WORKDIR /app
COPY --chown=app:app extract_contacts.py process_template_v2.py cls_template_users.xlsx TemplateV2.xlsx ./
COPY --chown=app:app webapp ./webapp
USER app

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 \
  CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://localhost:8000/api/health').status==200 else 1)"
CMD ["uvicorn", "webapp.app:app", "--host", "0.0.0.0", "--port", "8000"]
