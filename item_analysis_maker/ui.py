import tkinter as tk
from tkinter import filedialog,messagebox,ttk
from pathlib import Path
from .analysis import analyze,combine_datasets
from .outputs import generate_outputs
from .parsing import InputFormatError,parse_csv_batch

class ItemAnalysisApp:
    def __init__(self,root):
        self.root=root; root.title("Item Analysis Maker"); root.minsize(700,510)
        self.csv_paths=()
        self.csv=tk.StringVar()
        self.output=tk.StringVar(value=str(Path.home()/"Documents"))
        self.values={k:tk.StringVar(value=v) for k,v in {"exam_type":"Prelims","academic_year":"","semester":"1st Sem","subject":"","description":"","prepared_by":"","department_chairperson":""}.items()}
        self._build()
    def _build(self):
        f=ttk.Frame(self.root,padding=18); f.grid(sticky="nsew")
        self.root.columnconfigure(0,weight=1); self.root.rowconfigure(0,weight=1); f.columnconfigure(1,weight=1)
        ttk.Label(f,text="Item Analysis Maker",font=("Segoe UI",18,"bold")).grid(row=0,column=0,columnspan=3,sticky="w",pady=(0,14))
        ttk.Label(f,text="Class CSV files").grid(row=1,column=0,sticky="w",pady=5)
        ttk.Entry(f,textvariable=self.csv,state="readonly").grid(row=1,column=1,sticky="ew",padx=8)
        ttk.Button(f,text="Choose CSVs...",command=self._choose_csv).grid(row=1,column=2)
        ttk.Label(f,text="Output folder").grid(row=2,column=0,sticky="w",pady=5)
        ttk.Entry(f,textvariable=self.output).grid(row=2,column=1,sticky="ew",padx=8)
        ttk.Button(f,text="Browse...",command=self._choose_output).grid(row=2,column=2)
        ttk.Separator(f).grid(row=3,column=0,columnspan=3,sticky="ew",pady=12)
        r=4
        for label,key,choices in (("Exam type","exam_type",("Prelims","Midterms","Finals")),("Academic year","academic_year",None),("Semester","semester",("1st Sem","2nd Sem","Term Break")),("Subject","subject",None),("Description","description",None),("Prepared by","prepared_by",None),("Department Chairperson","department_chairperson",None)):
            ttk.Label(f,text=label).grid(row=r,column=0,sticky="w",pady=4)
            if choices: ttk.Combobox(f,textvariable=self.values[key],values=choices,state="readonly").grid(row=r,column=1,columnspan=2,sticky="ew",padx=8,pady=4)
            else: ttk.Entry(f,textvariable=self.values[key]).grid(row=r,column=1,columnspan=2,sticky="ew",padx=8,pady=4)
            r+=1
        ttk.Separator(f).grid(row=r,column=0,columnspan=3,sticky="ew",pady=12); r+=1
        ttk.Button(f,text="Analyze CSVs",command=self._preview).grid(row=r,column=0,sticky="w")
        ttk.Button(f,text="Generate course analysis",command=self._generate).grid(row=r,column=1,sticky="w",padx=8)
        self.status=ttk.Label(f,text="Choose one or more LMS or ZipGrade CSVs for the same course. Responses are pooled for one analysis.",wraplength=620)
        self.status.grid(row=r+1,column=0,columnspan=3,sticky="w",pady=(12,0))
    def _choose_csv(self):
        paths=filedialog.askopenfilenames(title="Choose class LMS or ZipGrade CSV files",filetypes=[("CSV files","*.csv")])
        if paths:
            self.csv_paths=tuple(Path(p) for p in paths)
            self.csv.set("; ".join(p.name for p in self.csv_paths))
            self._preview()
    def _choose_output(self):
        p=filedialog.askdirectory(title="Choose output folder")
        if p: self.output.set(p)
    def _load(self):
        if not self.csv_paths: raise InputFormatError("Choose at least one LMS or ZipGrade CSV.")
        parsed=parse_csv_batch(self.csv_paths)
        combined=combine_datasets(tuple(data for _,data in parsed))
        return tuple(path for path,_ in parsed),analyze(combined)
    def _preview(self):
        try:
            paths,result=self._load()
            self.status.configure(text=f"{len(paths)} class CSV(s) combined: {len(result.data.records)} eligible students, "
                                       f"{result.data.question_count} items; one pooled high group of {len(result.high_group)} "
                                       f"and low group of {len(result.low_group)}.")
        except Exception as e:
            messagebox.showerror("Could not analyze CSVs",str(e))
            self.status.configure(text="CSV batch validation failed. No outputs were generated.")
    def _generate(self):
        try:
            paths,result=self._load()
            if not self.values["subject"].get().strip(): raise ValueError("Enter a subject.")
            if not self.output.get().strip(): raise ValueError("Choose an output folder.")
            metadata={k:v.get().strip() for k,v in self.values.items()}
            template=Path(__file__).resolve().parents[1]/"ITEM_ANALYSIS_FORMAT.xlsx"
            from .outputs import output_paths
            xlsx,docx=output_paths(self.output.get(),metadata)
            existing=[str(p) for p in (xlsx,docx) if p.exists()]
            if existing and not messagebox.askyesno("Overwrite files?","Replace these existing files?\n\n"+"\n".join(existing)): return
            xlsx,docx=generate_outputs(template,self.output.get(),result,metadata)
            messagebox.showinfo("Analysis complete",f"Combined {len(paths)} class CSVs into one course analysis.\n\nCreated:\n{xlsx}\n{docx}")
            self.status.configure(text=f"Created one workbook and report from {len(paths)} CSVs.")
        except Exception as e:
            messagebox.showerror("Generation failed",str(e))

def main():
    root=tk.Tk(); ItemAnalysisApp(root); root.mainloop()


