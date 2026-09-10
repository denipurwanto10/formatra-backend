FROM node:20-slim

# LibreOffice headless (Word->PDF, PPT->PDF, and the PDF->Word fallback path)
# + fonts that are metrically compatible with common proprietary fonts
# (Carlito~=Calibri, Caladea~=Cambria, Liberation~=Arial/Times New Roman).
# This matters for fidelity: without them, documents built with those fonts
# get substituted with something whose character widths differ, and text
# reflows / line-wraps differently than the source.
#
# python3-pip is for pdf2docx, the primary PDF->Word engine (see
# pdf_to_docx.py) — used only by that one route; Word->PDF and PPT->PDF never
# touch Python.
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice \
      fonts-crosextra-carlito \
      fonts-crosextra-caladea \
      fonts-liberation \
      fonts-dejavu \
      python3 \
      python3-pip \
    && rm -rf /var/lib/apt/lists/*

# --break-system-packages: this is a dedicated single-purpose container
# image, so installing into the system Python (rather than a venv) is fine.
RUN pip3 install --no-cache-dir --break-system-packages pdf2docx

WORKDIR /app
COPY index.mjs pdf_to_docx.py ./

ENV PORT=8787
ENV ALLOWED_ORIGIN=*
EXPOSE 8787

# No npm install needed — index.mjs only uses Node's built-in modules.
CMD ["node", "index.mjs"]
