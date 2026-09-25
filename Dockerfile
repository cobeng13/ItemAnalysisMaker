FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1
WORKDIR /app
COPY requirements.txt requirements-web.txt ./
RUN pip install --no-cache-dir --disable-pip-version-check -r requirements-web.txt \
    && useradd --create-home --uid 10001 appuser
COPY item_analysis_maker ./item_analysis_maker
COPY ITEM_ANALYSIS_FORMAT.xlsx ./ITEM_ANALYSIS_FORMAT.xlsx
USER appuser
EXPOSE 8000
CMD ["uvicorn", "item_analysis_maker.web.app:app", "--host", "0.0.0.0", "--port", "8000"]
