FROM node:20-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice \
      fonts-crosextra-carlito \
      fonts-crosextra-caladea \
      fonts-liberation \
      fonts-dejavu \
      python3 \
      python3-pip \
    && rm -rf /var/lib/apt/lists/*

RUN pip3 install --no-cache-dir --break-system-packages pdf2docx

WORKDIR /app

COPY index.mjs pdf_to_docx.py ./

ENV ALLOWED_ORIGIN=*

CMD ["node", "index.mjs"]