FROM node:24-bookworm-slim AS runtime
WORKDIR /opt/office/harness
COPY harness/package.json harness/package-lock.json ./
RUN npm ci --omit=dev --no-audit --no-fund

FROM python:3.12-slim-bookworm
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
WORKDIR /opt/office
COPY --from=runtime /usr/local/bin/node /usr/local/bin/node
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt \
    && groupadd -r office && useradd -r -g office office \
    && mkdir -p /opt/office/data && chown office:office /opt/office/data
COPY server server
COPY --from=runtime /opt/office/harness/node_modules harness/node_modules
COPY harness/office.patch.yml harness/office-context.mjs harness/todo.example.patch.yml harness/
USER office
EXPOSE 8765
CMD ["python","-m","uvicorn","server.app:app","--host","0.0.0.0","--port","8765","--no-access-log"]
