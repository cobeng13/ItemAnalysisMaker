from dataclasses import dataclass
from pathlib import Path
import csv, re

class InputFormatError(ValueError): pass

@dataclass(frozen=True)
class StudentRecord:
    score: float
    answers: tuple[bool, ...]

@dataclass(frozen=True)
class InputData:
    format: str
    records: tuple[StudentRecord, ...]
    question_count: int
    excluded_rows: int

_LMS = re.compile(r"^Q\.\s*(\d+)\s*/\s*([0-9.]+)$", re.I)
_ZIP = re.compile(r"^Q\s*(\d+)$", re.I)

def _index(headers, name):
    try: return [h.strip().casefold() for h in headers].index(name.casefold())
    except ValueError as e: raise InputFormatError(f"Required column is missing: {name}") from e

def _num(v, where, blank=False):
    v=v.strip()
    if not v and blank: return None
    try: n=float(v)
    except (TypeError,ValueError) as e: raise InputFormatError(f"Invalid number in {where}: {v or '(blank)'}") from e
    if not (-float("inf") < n < float("inf")): raise InputFormatError(f"Invalid number in {where}.")
    return n

def _zip_records(rows, score_col, specs, n):
    records=[]
    for rn,row in rows:
        score=_num(row[score_col] if score_col<len(row) else "",f"row {rn} Num Correct")
        if score<0 or score>n or score!=int(score): raise InputFormatError(f"Num Correct in row {rn} must be an integer from 0 to {n}.")
        answers=[]
        for q,col in specs:
            v=_num(row[col] if col<len(row) else "",f"row {rn} question {q}",True)
            if v is not None and v not in (0,1): raise InputFormatError(f"ZipGrade question {q} in row {rn} must be 0, 1, or blank.")
            answers.append(v==1 if v is not None else False)
        records.append(StudentRecord(score,tuple(answers)))
    if not records: raise InputFormatError("No ZipGrade attempts were found.")
    return InputData("ZipGrade",tuple(records),n,0)

def parse_csv(path):
    path=Path(path)
    if path.suffix.casefold()!=".csv": raise InputFormatError("Choose an LMS or ZipGrade .csv file.")
    try:
        with path.open(encoding="utf-8-sig",newline="") as f:
            sample=f.read(8192); f.seek(0)
            try: dialect=csv.Sniffer().sniff(sample,delimiters=",;\t")
            except csv.Error: dialect=csv.excel
            reader=csv.reader(f,dialect); headers=next(reader,None)
            if not headers: raise InputFormatError("The CSV is empty.")
            headers=[h.strip() for h in headers]; folded={h.casefold() for h in headers}
            lms=[]; zipq=[]
            for col,h in enumerate(headers):
                m=_LMS.match(h)
                if m: lms.append((int(m.group(1)),col,float(m.group(2)))); continue
                m=_ZIP.match(h)
                if m: zipq.append((int(m.group(1)),col))
            if {"status","grade/100.00"}<=folded and lms:
                lms.sort()
                if [x[0] for x in lms]!=list(range(1,len(lms)+1)): raise InputFormatError("LMS question columns must be numbered consecutively from 1.")
                if len(lms)>100: raise InputFormatError("The workbook template supports up to 100 items.")
                status_col=_index(headers,"Status"); score_col=_index(headers,"Grade/100.00")
                records=[]; excluded=0
                for rn,row in enumerate(reader,2):
                    if not row or not any(x.strip() for x in row): continue
                    status=row[status_col].strip().casefold() if status_col<len(row) else ""
                    if status!="finished": excluded+=1; continue
                    score=_num(row[score_col] if score_col<len(row) else "",f"row {rn} Grade/100.00")
                    if score<0 or score>100: raise InputFormatError(f"Grade/100.00 in row {rn} must be between 0 and 100.")
                    answers=[]
                    for q,col,maximum in lms:
                        v=_num(row[col] if col<len(row) else "",f"row {rn} question {q}",True)
                        if v is not None and (v<0 or v>maximum): raise InputFormatError(f"Question {q} in row {rn} is outside 0 to {maximum:g}.")
                        answers.append(v is not None and v>=maximum)
                    records.append(StudentRecord(score,tuple(answers)))
                if not records: raise InputFormatError("No Finished LMS attempts were found.")
                return InputData("LMS",tuple(records),len(lms),excluded)
            if {"num questions","num correct"}<=folded and zipq:
                zipq.sort()
                if [x[0] for x in zipq]!=list(range(1,len(zipq)+1)): raise InputFormatError("ZipGrade question columns must be numbered consecutively from Q1.")
                declared_col=_index(headers,"Num Questions"); score_col=_index(headers,"Num Correct")
                rows=[]; counts=set()
                for rn,row in enumerate(reader,2):
                    if not row or not any(x.strip() for x in row): continue
                    declared=_num(row[declared_col] if declared_col<len(row) else "",f"row {rn} Num Questions")
                    if declared<1 or declared>len(zipq) or declared!=int(declared): raise InputFormatError(f"Invalid Num Questions value in row {rn}.")
                    counts.add(int(declared)); rows.append((rn,row))
                if len(counts)!=1: raise InputFormatError("ZipGrade attempts must have the same Num Questions value.")
                n=counts.pop()
                if n>100: raise InputFormatError("The workbook template supports up to 100 items.")
                return _zip_records(rows,score_col,[x for x in zipq if x[0]<=n],n)
            raise InputFormatError("CSV format not recognized. Use an LMS or ZipGrade export.")
    except UnicodeDecodeError as e: raise InputFormatError("The CSV must use UTF-8 encoding.") from e
    except OSError as e: raise InputFormatError(f"Could not read the CSV: {e}") from e


def parse_csv_batch(paths):
    """Parse every class export and require a common item count before analysis."""
    paths=tuple(Path(path) for path in paths)
    if not paths:
        raise InputFormatError("Choose at least one LMS or ZipGrade CSV.")
    parsed=[]
    for path in paths:
        try:
            parsed.append((path,parse_csv(path)))
        except InputFormatError as exc:
            raise InputFormatError(f"{path.name}: {exc}") from exc
    counts={data.question_count for _,data in parsed}
    if len(counts)>1:
        details=", ".join(f"{path.name}: {data.question_count} items" for path,data in parsed)
        raise InputFormatError("All CSVs must contain the same number of items. No files were generated. "+details)
    return tuple(parsed)
