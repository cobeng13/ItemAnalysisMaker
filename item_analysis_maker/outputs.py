from pathlib import Path
from copy import copy
import re
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches,Pt
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from .analysis import Analysis

class OutputError(RuntimeError): pass

def safe_part(value,fallback):
    return (re.sub(r"[^A-Za-z0-9_-]+","_",value.strip()).strip("_-")[:60] or fallback)

def _shift_merges(ws,at,amount):
    old=list(ws.merged_cells.ranges)
    for r in old: ws.unmerge_cells(str(r))
    for r in old:
        c1,r1,c2,r2=r.bounds
        if r1>=at: r1+=amount; r2+=amount
        elif r2>=at: raise OutputError("A merged template cell crosses the student rows.")
        if c1!=c2 or r1!=r2: ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)

def _copy_style(ws,source,target):
    for col in range(1,103):
        a,b=ws.cell(source,col),ws.cell(target,col)
        if a.has_style: b._style=a._style
        b.alignment=copy(a.alignment)
    ws.row_dimensions[target].height=ws.row_dimensions[source].height

def _expand(ws,et,el):
    if et:
        _shift_merges(ws,34,et); ws.insert_rows(34,et)
        for r in range(34,34+et): _copy_style(ws,33,r)
    base=61+et
    if el:
        _shift_merges(ws,base,el); ws.insert_rows(base,el)
        for r in range(base,base+el): _copy_style(ws,base-1,r)
    ws.print_area=f"A1:CX{78+et+el}"

def _metadata(ws,m,n):
    labels={"Prelims":"Prelim","Midterms":"Midterm","Finals":"Finals"}
    for row,label in ((10,"Prelim"),(11,"Midterm"),(12,"Finals")):
        ws.cell(row,1).value=("X" if labels[m["exam_type"]]==label else " ")+" "+label
    ws["E10"]=m["academic_year"]; ws["E11"]=m["semester"]
    ws["J10"]=m["subject"]; ws["J11"]=m["description"]; ws["X10"]=n
    ws["X11"]="=X10*0.27"; ws["X14"]="=MIN(ROUNDUP(X11,0),16)"
    ws["A72"]=f"Prepared by: {m['prepared_by']}"
    ws["V72"]=f"Approved by: {m['department_chairperson']}"

def _student_header(ws,row):
    left=f"A{row}:B{row+1}"
    label=ws.cell(row,1)
    style=label._style
    alignment=copy(label.alignment)
    ws.unmerge_cells(left)
    label=ws.cell(row,1)
    target=ws.cell(row,2)
    label.value=None
    target._style=style
    target.alignment=alignment
    target.value="Student Names"
    ws.merge_cells(start_row=row,start_column=2,end_row=row+1,end_column=2)

def _group(ws,records,start,slots,nitems,first_number):
    for i in range(slots):
        row=start+i
        number=first_number+i
        ws.cell(row,1).value=number
        ws.cell(row,2).value=f"Student {number}" if i<len(records) else None
        for col in (1,2):
            cell=ws.cell(row,col)
            old=cell.font
            cell.font=Font(name="Arial Narrow",size=10,bold=old.bold,italic=old.italic,
                           vertAlign=old.vertAlign,underline=old.underline,strike=old.strike,color=old.color)
        for q in range(nitems):
            ws.cell(row,q+3).value=int(records[i].answers[q]) if i<len(records) else None

def _formulas(ws,nitems,hs,hn,ls,ln):
    ht=hs+hn; lt=ls+ln; ph=ht+1; pl=lt+1; p=lt+2; di=lt+3; remark=lt+4
    for item in range(1,101):
        col=item+2; letter=get_column_letter(col)
        if item<=nitems:
            ws.cell(ht,col).value=f'=COUNTIF({letter}{hs}:{letter}{ht-1},1)'
            ws.cell(ph,col).value=f'=IFERROR({letter}{ht}/COUNTA($B$'+str(hs)+':$B$'+str(ht-1)+'),0)'
            ws.cell(lt,col).value=f'=COUNTIF({letter}{ls}:{letter}{lt-1},1)'
            ws.cell(pl,col).value=f'=IFERROR({letter}{lt}/COUNTA($B$'+str(ls)+':$B$'+str(lt-1)+'),0)'
            ws.cell(p,col).value=f'=({letter}{ph}+{letter}{pl})/2'
            ws.cell(di,col).value=f'={letter}{ph}-{letter}{pl}'
            ws.cell(remark,col).value=f'=IF({letter}{di}>0.2,"G",IF({letter}{di}>0.1,"M","P"))'
        else:
            for row in (ht,ph,lt,pl,p,di,remark): ws.cell(row,col).value=None

def _workbook(template,out,analysis,m):
    wb=load_workbook(template); ws=wb["Items Analysis"]
    hn=max(16,len(analysis.high_group)); ln=max(16,len(analysis.low_group))
    et,el=hn-16,ln-16; _expand(ws,et,el)
    hs=18; ls=45+et
    _metadata(ws,m,len(analysis.data.records))
    _student_header(ws,16)
    _student_header(ws,ls-2)
    _group(ws,analysis.high_group,hs,hn,analysis.data.question_count,1)
    _group(ws,analysis.low_group,ls,ln,analysis.data.question_count,len(analysis.high_group)+1)
    _formulas(ws,analysis.data.question_count,hs,hn,ls,ln)
    wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode="auto"
    wb.save(out)

def _report(path,analysis,m):
    doc=Document(); section=doc.sections[0]; section.top_margin=Inches(.65); section.bottom_margin=Inches(.65)
    doc.styles["Normal"].font.name="Arial"; doc.styles["Normal"].font.size=Pt(10)
    title=doc.add_heading("ITEM ANALYSIS SUMMARY",0); title.alignment=WD_ALIGN_PARAGRAPH.CENTER
    meta=doc.add_table(rows=0,cols=2); meta.style="Light Shading Accent 1"
    for label,value in (("Exam",m["exam_type"]),("Academic Year",m["academic_year"]),("Semester",m["semester"]),("Subject",m["subject"]),("Description",m["description"]),("Students analyzed",str(len(analysis.data.records))),("Top group / bottom group",f"{len(analysis.high_group)} / {len(analysis.low_group)}"),("Prepared by",m["prepared_by"]),("Department Chairperson",m["department_chairperson"])):
        cells=meta.add_row().cells; cells[0].text=label; cells[1].text=value
    counts={c:sum(x.classification==c for x in analysis.items) for c in ("Good","Marginal","Poor")}
    doc.add_heading("Summary",level=1)
    doc.add_paragraph(f"Good: {counts['Good']} items | Marginal: {counts['Marginal']} items | Poor: {counts['Poor']} items")
    if m.get("class_label"):
        cells=meta.add_row().cells; cells[0].text="Class / CSV"; cells[1].text=m["class_label"]
    doc.add_paragraph(f"Difficulty is the average correct rate for the high and low groups. Discrimination is high-group correct rate minus low-group correct rate. Good: >0.20; Marginal: >0.10 through 0.20; Poor: <=0.10. Nominal group size is ceil(27% of {len(analysis.data.records)}) = {analysis.nominal_group_size}; ties at score boundaries are included, subject to a maximum of 16 students in each group.")
    for category in ("Good","Marginal","Poor"):
        items=[x for x in analysis.items if x.classification==category]
        doc.add_heading(f"{category} Items ({len(items)})",level=1)
        if not items: doc.add_paragraph("None."); continue
        table=doc.add_table(rows=1,cols=4); table.style="Light Shading Accent 1"
        for cell,label in zip(table.rows[0].cells,("Item","Difficulty (p)","Discrimination (D)","High / Low Correct")): cell.text=label
        for item in items:
            cells=table.add_row().cells
            cells[0].text=str(item.number); cells[1].text=f"{item.difficulty:.3f}"; cells[2].text=f"{item.discrimination:.3f}"
            cells[3].text=f"{item.high_correct}/{len(analysis.high_group)}; {item.low_correct}/{len(analysis.low_group)}"
    doc.add_paragraph("Student names, email addresses, and other identifying information are not included.")
    doc.save(path)

def output_paths(output_dir,metadata,source_label=None):
    folder=Path(output_dir)
    stem=f"Item_Analysis_{safe_part(metadata['subject'],'Subject')}_{safe_part(metadata['exam_type'],'Exam')}"
    if source_label:
        stem+=f"_{safe_part(source_label,'Class')}"
    return folder/f"{stem}.xlsx",folder/f"{stem}_Summary.docx"

def generate_outputs(template_path,output_dir,analysis,metadata,source_label=None):
    template=Path(template_path)
    if not template.is_file(): raise OutputError(f"Excel template not found: {template}")
    folder=Path(output_dir); folder.mkdir(parents=True,exist_ok=True)
    xlsx,docx=output_paths(folder,metadata,source_label)
    _workbook(template,xlsx,analysis,metadata)
    report_metadata=dict(metadata)
    if source_label: report_metadata["class_label"]=source_label
    _report(docx,analysis,report_metadata)
    return xlsx,docx


