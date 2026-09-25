from dataclasses import dataclass
from math import ceil
from .parsing import InputData, InputFormatError, StudentRecord

@dataclass(frozen=True)
class ItemResult:
    number: int
    difficulty: float
    discrimination: float
    high_correct: int
    low_correct: int
    classification: str

@dataclass(frozen=True)
class Analysis:
    data: InputData
    high_group: tuple[StudentRecord,...]
    low_group: tuple[StudentRecord,...]
    nominal_group_size: int
    items: tuple[ItemResult,...]

def _boundary(records,size,reverse):
    ordered=sorted(enumerate(records),key=lambda p:(-p[1].score if reverse else p[1].score,p[0]))
    boundary=ordered[size-1][1].score
    selected=tuple(r for _,r in ordered if (r.score>=boundary if reverse else r.score<=boundary))
    return selected[:16]

def analyze(data):
    if not data.records: raise ValueError("There are no eligible attempts to analyze.")
    nominal=ceil(len(data.records)*.27)
    high=_boundary(data.records,nominal,True); low=_boundary(data.records,nominal,False)
    items=[]
    for i in range(data.question_count):
        hc=sum(r.answers[i] for r in high); lc=sum(r.answers[i] for r in low)
        ph=hc/len(high); pl=lc/len(low); d=ph-pl
        cls="Good" if d>.20 else "Marginal" if d>.10 else "Poor"
        items.append(ItemResult(i+1,(ph+pl)/2,d,hc,lc,cls))
    return Analysis(data,high,low,nominal,tuple(items))


def combine_datasets(datasets):
    """Pool eligible responses from same-length class exports for one course analysis."""
    datasets=tuple(datasets)
    if not datasets:
        raise InputFormatError("Choose at least one LMS or ZipGrade CSV.")
    counts={data.question_count for data in datasets}
    if len(counts)>1:
        detail=", ".join(f"{data.format}: {data.question_count} items" for data in datasets)
        raise InputFormatError("All CSVs must contain the same number of items. No outputs were generated. "+detail)
    formats={data.format for data in datasets}
    source_format=next(iter(formats)) if len(formats)==1 else "Mixed LMS/ZipGrade"
    return InputData(source_format,tuple(record for data in datasets for record in data.records),
                     datasets[0].question_count,sum(data.excluded_rows for data in datasets))
