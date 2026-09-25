# ItemAnalysisMaker

Windows desktop app for LMS and ZipGrade response exports. It creates an item analysis workbook from ITEM_ANALYSIS_FORMAT.xlsx and a Word summary report.

## Run

Install Python 3.10+ and dependencies in a virtual environment:

~~~powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m item_analysis_maker
~~~

Tkinter is included with the standard Windows Python installer. Choose one or more class CSV files for the same course and an output folder, enter the exam metadata, then generate one Excel workbook and one Word report. The app pools eligible responses across the selected CSVs and calculates one high/low pair for the course. Every CSV is validated before generation; all selected exports must contain the same number of items.

The analysis engine is separate from Tkinter so a future web app can reuse parse_csv, analyze, and generate_outputs.

## Analysis rules

- LMS rows count only when Status is Finished; all ZipGrade attempts are included.
- Student identity columns are ignored and never written to output. Groups use generic Student 1, Student 2, ... labels.
- Each pooled group begins at ceil(27% of eligible attempts), includes ties in score order, and is capped at 16 students. High labels run Student 1 to Student 16; low labels continue Student 17 to Student 32.
- Difficulty is the average correct rate for high and low groups. Discrimination is high-group correct rate minus low-group correct rate.
- Discrimination >0.20 is Good; >0.10 through 0.20 is Marginal; <=0.10 is Poor.

LMS requires Status, Grade/100.00, and consecutive Q. n / points columns. ZipGrade requires Num Questions, Num Correct, and consecutive Qn columns. ZipGrade responses must be 0, 1, or blank. Blank item responses count as incorrect; LMS scores count correct when equal to the question maximum in the header.




## Web version

Install the web dependencies and start the app locally:

~~~powershell
python -m pip install -r requirements-web.txt
uvicorn item_analysis_maker.web.app:app --reload
~~~

Open http://localhost:8000. Select all class exports for one course; the app validates matching item counts, pools the responses, then downloads one ZIP containing the Excel analysis and Word summary.

Build and run the Docker image using the same container pattern as GIFT2HardCopy:

~~~powershell
docker build -t item-analysis-maker .
docker run --rm -p 8000:8000 item-analysis-maker
~~~

The container listens on port 8000 and runs as a non-root user.
