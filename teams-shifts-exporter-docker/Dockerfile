FROM node:20-slim

RUN apt-get update && apt-get install -y \
    wget gnupg2 ca-certificates \
    xvfb x11vnc \
    --no-install-recommends \
  && wget -q -O - https://dl.google.com/linux/linux_signing_key.pub \
     | gpg --dearmor -o /usr/share/keyrings/google-chrome.gpg \
  && echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome.gpg] \
     http://dl.google.com/linux/chrome/deb/ stable main" \
     > /etc/apt/sources.list.d/google-chrome.list \
  && apt-get update && apt-get install -y google-chrome-stable \
  && rm -rf /var/lib/apt/lists/*

COPY extension/ /app/extension/

COPY runner/package.json runner/package-lock.json* /app/runner/
RUN cd /app/runner && npm ci --omit=dev

COPY runner/ /app/runner/
RUN chmod +x /app/runner/entrypoint.sh

CMD ["/app/runner/entrypoint.sh"]
