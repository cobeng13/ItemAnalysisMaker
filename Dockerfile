FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1
WORKDIR /app
COPY requirements.txt requirements-web.txt ./
RUN pip install --no-cache-dir --disable-pip-version-check -r requirements-web.txt \
    && useradd --create-home --uid 10001 appuser
COPY item_analysis_maker ./item_analysis_maker
COPY ITEM_ANALYSIS_FORMAT.xlsx ./ITEM_ANALYSIS_FORMAT.xlsx
COPY ITEM_ANALYSIS_SUMMARY_FORMAT.docx ./ITEM_ANALYSIS_SUMMARY_FORMAT.docx
USER appuser
EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 CMD ["python", "-c", "from urllib.request import urlopen; urlopen('http://127.0.0.1:8000/', timeout=3).read(1)"]
CMD ["uvicorn", "item_analysis_maker.web.app:app", "--host", "0.0.0.0", "--port", "8000"]
