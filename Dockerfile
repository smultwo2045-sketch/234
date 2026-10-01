FROM python:3.12-slim

WORKDIR /srv

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    WRS_DATA_DIR=/data \
    PYTHONPATH=/srv

COPY requirements.txt /srv/requirements.txt
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r /srv/requirements.txt

COPY app /srv/app
COPY start.sh /srv/start.sh
RUN chmod +x /srv/start.sh && mkdir -p /data

EXPOSE 8000
CMD ["/srv/start.sh"]
