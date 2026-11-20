# FOUR4 CUT PRO - Windows panel cutting optimizer
# Requires Python 3.10+. Optional exports: pip install openpyxl reportlab
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json, math, os
from dataclasses import dataclass, asdict

APP_NAME = "FOUR4 CUT PRO"

@dataclass
class Part:
    name: str
    w: float
    h: float
    qty: int
    material: str
    thickness: str
    edge_top: bool = False
    edge_bottom: bool = False
    edge_left: bool = False
    edge_right: bool = False
    grain: bool = False

@dataclass
class Placed:
    part: Part
    x: float
    y: float
    w: float
    h: float
    rotated: bool

class Sheet:
    def __init__(self, w, h, kerf, margin):
        self.w, self.h, self.kerf, self.margin = w, h, kerf, margin
        self.free = [(margin, margin, w-2*margin, h-2*margin)]
        self.placed = []

    def put(self, part, allow_rotate):
        # Best-short-side-fit over guillotine free rectangles.
        choices = []
        for i, (x,y,fw,fh) in enumerate(self.free):
            orientations = [(part.w, part.h, False)]
            if allow_rotate and not part.grain and abs(part.w-part.h) > 1e-6:
                orientations.append((part.h, part.w, True))
            for pw, ph, rot in orientations:
                if pw <= fw + 1e-6 and ph <= fh + 1e-6:
                    choices.append((min(fw-pw, fh-ph), max(fw-pw, fh-ph), i, x, y, fw, fh, pw, ph, rot))
        if not choices:
            return False
        _, _, idx, x,y,fw,fh,pw,ph,rot = min(choices)
        self.free.pop(idx)
        # Kerf is reserved between parts.
        rw, rh = fw-pw-self.kerf, fh-ph-self.kerf
        # Split into right and bottom rectangles; choose split that preserves larger usable area.
        right = (x+pw+self.kerf, y, max(0,rw), fh)
        bottom = (x, y+ph+self.kerf, pw, max(0,rh))
        if fw-pw > fh-ph:
            right = (x+pw+self.kerf, y, max(0,rw), fh)
            bottom = (x, y+ph+self.kerf, pw, max(0,rh))
        else:
            right = (x+pw+self.kerf, y, max(0,rw), ph)
            bottom = (x, y+ph+self.kerf, fw, max(0,rh))
        for r in (right,bottom):
            if r[2] > 0.5 and r[3] > 0.5:
                self.free.append(r)
        self.placed.append(Placed(part,x,y,pw,ph,rot))
        return True

def optimize(parts, sheet_w, sheet_h, kerf, margin, allow_rotate):
    # Expand quantities; largest-side/area first. Try both sheet orientations.
    items=[]
    for p in parts:
        for _ in range(p.qty):
            items.append(p)
    items.sort(key=lambda p: (max(p.w,p.h), p.w*p.h), reverse=True)
    sheets=[]
    for p in items:
        done=False
        for sh in sheets:
            if sh.put(p, allow_rotate):
                done=True; break
        if not done:
            candidates=[]
            for sw,shh in [(sheet_w,sheet_h),(sheet_h,sheet_w)]:
                sh=Sheet(sw,shh,kerf,margin)
                if sh.put(p,allow_rotate):
                    candidates.append(sh)
            if not candidates:
                raise ValueError(f'Detal "{p.name}" ({p.w} × {p.h} mm) lövhəyə yerləşmir.')
            sheets.append(candidates[0])
    return sheets

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1240x820")
        self.minsize(1050,700)
        self.parts=[]
        self.sheets=[]
        self._style()
        self._ui()

    def _style(self):
        s=ttk.Style(self)
        try: s.theme_use("clam")
        except: pass
        s.configure("Title.TLabel",font=("Segoe UI",19,"bold"),foreground="#17324d")
        s.configure("Sub.TLabel",font=("Segoe UI",10),foreground="#526579")
        s.configure("Treeview",rowheight=27,font=("Segoe UI",9))
        s.configure("Treeview.Heading",font=("Segoe UI",9,"bold"))

    def _ui(self):
        head=ttk.Frame(self,padding=(16,12)); head.pack(fill="x")
        ttk.Label(head,text="FOUR4 CUT PRO",style="Title.TLabel").pack(side="left")
        ttk.Label(head,text="Mebel lövhələrinin kəsim optimallaşdırılması",style="Sub.TLabel").pack(side="left",padx=18,pady=8)
        bar=ttk.Frame(self,padding=(12,4)); bar.pack(fill="x")
        ttk.Button(bar,text="Yeni layihə",command=self.new_project).pack(side="left",padx=3)
        ttk.Button(bar,text="Layihəni aç",command=self.open_project).pack(side="left",padx=3)
        ttk.Button(bar,text="Yadda saxla",command=self.save_project).pack(side="left",padx=3)
        ttk.Button(bar,text="Excel ixracı",command=self.export_excel).pack(side="left",padx=3)
        ttk.Button(bar,text="PDF hesabat",command=self.export_pdf).pack(side="left",padx=3)

        settings=ttk.LabelFrame(self,text="Lövhə və kəsim parametrləri (mm)",padding=10)
        settings.pack(fill="x",padx=14,pady=6)
        self.vars={}
        fields=[("sheet_w","Lövhə eni",2800),("sheet_h","Lövhə uzunluğu",2070),("kerf","Mişar payı",3.2),("margin","Kənar təhlükəsizlik payı",10)]
        for i,(key,label,val) in enumerate(fields):
            ttk.Label(settings,text=label).grid(row=0,column=i*2,sticky="w",padx=(4,5))
            v=tk.StringVar(value=str(val)); self.vars[key]=v
            ttk.Entry(settings,textvariable=v,width=11).grid(row=0,column=i*2+1,padx=(0,14))
        self.rotate=tk.BooleanVar(value=True)
        ttk.Checkbutton(settings,text="Lif istiqaməti olmayan detalları 90° döndərməyə icazə ver",variable=self.rotate).grid(row=1,column=0,columnspan=4,sticky="w",pady=(9,0))

        body=ttk.Panedwindow(self,orient="horizontal"); body.pack(fill="both",expand=True,padx=14,pady=8)
        left=ttk.Frame(body); right=ttk.Frame(body); body.add(left,weight=3); body.add(right,weight=4)
        form=ttk.LabelFrame(left,text="Detal əlavə et",padding=10); form.pack(fill="x")
        self.f={}
        rows=[("name","Detal adı","Yan panel"),("w","En (mm)","600"),("h","Uzunluq (mm)","720"),("qty","Say","2"),("material","Material","MDF"),("thickness","Qalınlıq","18")]
        for r,(k,l,d) in enumerate(rows):
            ttk.Label(form,text=l).grid(row=r,column=0,sticky="w",pady=3)
            v=tk.StringVar(value=d); self.f[k]=v
            if k=="material":
                wid=ttk.Combobox(form,textvariable=v,values=["MDF","DSP","Faner","Taxta","Digər"],width=18)
            else: wid=ttk.Entry(form,textvariable=v,width=21)
            wid.grid(row=r,column=1,sticky="ew",pady=3)
        self.edge_vars={}
        edgebox=ttk.LabelFrame(form,text="PVC kənar bantlama",padding=6); edgebox.grid(row=6,column=0,columnspan=2,sticky="ew",pady=6)
        for i,(k,l) in enumerate([("edge_top","Üst"),("edge_bottom","Alt"),("edge_left","Sol"),("edge_right","Sağ")]):
            v=tk.BooleanVar(); self.edge_vars[k]=v
            ttk.Checkbutton(edgebox,text=l,variable=v).grid(row=0,column=i,padx=4)
        self.grain=tk.BooleanVar(value=False)
        ttk.Checkbutton(form,text="Lif/dekor istiqaməti qorunsun",variable=self.grain).grid(row=7,column=0,columnspan=2,sticky="w",pady=4)
        ttk.Button(form,text="＋ Detalı siyahıya əlavə et",command=self.add_part).grid(row=8,column=0,columnspan=2,sticky="ew",pady=7)
        ttk.Button(form,text="Seçilmiş detalı sil",command=self.remove_part).pack(fill="x",pady=5)
        cols=("name","w","h","qty","material","thickness","edges","grain")
        self.tree=ttk.Treeview(left,columns=cols,show="headings",height=12)
        labels=("Detal","En","Uzunluq","Say","Material","Qalınlıq","PVC","Lif")
        widths=(115,55,70,40,65,65,60,45)
        for c,l,w in zip(cols,labels,widths):
            self.tree.heading(c,text=l); self.tree.column(c,width=w,anchor="center")
        self.tree.pack(fill="both",expand=True,pady=7)

        top=ttk.Frame(right); top.pack(fill="x")
        ttk.Button(top,text="RASKROY HESABLA",command=self.run).pack(side="left",fill="x",expand=True,padx=(0,5))
        ttk.Button(top,text="Kəsim sxemini çap et",command=self.print_hint).pack(side="left",fill="x",expand=True)
        self.summary=ttk.Label(right,text="Detalları daxil edin və “RASKROY HESABLA” düyməsini basın.",justify="left",padding=8)
        self.summary.pack(fill="x")
        self.canvas=tk.Canvas(right,bg="#f5f7fa",highlightthickness=1,highlightbackground="#ccd5df")
        self.canvas.pack(fill="both",expand=True,pady=5)
        self.sheet_select=ttk.Combobox(right,state="readonly",values=[])
        self.sheet_select.pack(fill="x")
        self.sheet_select.bind("<<ComboboxSelected>>",self.draw_sheet)
        self.status=tk.StringVar(value="Hazırdır.")
        ttk.Label(self,textvariable=self.status,relief="sunken",anchor="w",padding=5).pack(fill="x",side="bottom")

    def add_part(self):
        try:
            name=self.f["name"].get().strip() or "Detal"
            w=float(self.f["w"].get()); h=float(self.f["h"].get()); q=int(self.f["qty"].get())
            if w<=0 or h<=0 or q<1: raise ValueError()
            p=Part(name,w,h,q,self.f["material"].get(),self.f["thickness"].get(),
                   **{k:v.get() for k,v in self.edge_vars.items()},grain=self.grain.get())
            self.parts.append(p); self.refresh_tree(); self.status.set(f"{name} əlavə edildi.")
        except: messagebox.showerror("Xəta","En, uzunluq və say üçün düzgün müsbət rəqəmlər daxil edin.")

    def refresh_tree(self):
        for x in self.tree.get_children(): self.tree.delete(x)
        for i,p in enumerate(self.parts):
            edges=",".join(k.replace("edge_","") for k in ["edge_top","edge_bottom","edge_left","edge_right"] if getattr(p,k)) or "—"
            self.tree.insert("","end",iid=str(i),values=(p.name,p.w,p.h,p.qty,p.material,p.thickness,edges,"Bəli" if p.grain else "Xeyr"))

    def remove_part(self):
        sel=self.tree.selection()
        if not sel:return
        for i in sorted([int(x) for x in sel],reverse=True): self.parts.pop(i)
        self.refresh_tree()

    def params(self):
        sw=float(self.vars["sheet_w"].get()); sh=float(self.vars["sheet_h"].get())
        kerf=float(self.vars["kerf"].get()); margin=float(self.vars["margin"].get())
        if min(sw,sh)<=0 or kerf<0 or margin<0 or sw<=2*margin or sh<=2*margin: raise ValueError("Lövhə ölçüləri, mişar payı və kənar payını yoxlayın.")
        return sw,sh,kerf,margin

    def run(self):
        if not self.parts: messagebox.showwarning("Məlumat yoxdur","Əvvəlcə detallar əlavə edin."); return
        try:
            sw,sh,k,m=self.params()
            groups={}
            for p in self.parts: groups.setdefault((p.material,p.thickness),[]).append(p)
            self.sheets=[]; labels=[]
            for (mat,th),ps in groups.items():
                res=optimize(ps,sw,sh,k,m,self.rotate.get())
                for s in res: s.material=mat; s.thickness=th; self.sheets.append(s)
            for i,s in enumerate(self.sheets): labels.append(f"Lövhə {i+1} — {s.material} {s.thickness} mm")
            self.sheet_select["values"]=labels
            if labels:self.sheet_select.current(0)
            total_parts=sum(len(s.placed) for s in self.sheets)
            board_area=sum(s.w*s.h for s in self.sheets)
            used=sum(x.w*x.h for s in self.sheets for x in s.placed)
            waste=max(0,100*(board_area-used)/board_area) if board_area else 0
            self.summary.config(text=f"Lövhə sayı: {len(self.sheets)}    |    Detal sayı: {total_parts}\nMaterial sahəsi: {board_area/1e6:.2f} m²    |    Detal sahəsi: {used/1e6:.2f} m²\nTəxmini tullantı: {waste:.1f}%  (kəsim payı və kənar payı daxil deyil)")
            self.draw_sheet(); self.status.set("Raskroy hesablandı. Nəticəni yoxlayın.")
        except Exception as e: messagebox.showerror("Raskroy xətası",str(e))

    def draw_sheet(self,event=None):
        self.canvas.delete("all")
        idx=self.sheet_select.current()
        if idx<0 or idx>=len(self.sheets):return
        s=self.sheets[idx]
        cw=max(100,self.canvas.winfo_width()-30); ch=max(100,self.canvas.winfo_height()-30)
        scale=min(cw/s.w,ch/s.h)
        ox=(cw-s.w*scale)/2+10; oy=(ch-s.h*scale)/2+10
        self.canvas.create_rectangle(ox,oy,ox+s.w*scale,oy+s.h*scale,fill="#e4d2ad",outline="#795b35",width=2)
        for j,p in enumerate(s.placed):
            x=ox+p.x*scale; y=oy+p.y*scale; w=p.w*scale; h=p.h*scale
            self.canvas.create_rectangle(x,y,x+w,y+h,fill="#d5e7f4",outline="#245779",width=1)
            label=f"{p.part.name}\n{p.w:g}×{p.h:g}"
            if w>35 and h>24:self.canvas.create_text(x+w/2,y+h/2,text=label,font=("Segoe UI",8),fill="#17324d",justify="center")
        self.canvas.create_text(ox,oy-8,anchor="sw",text=f"{s.w:g} × {s.h:g} mm",font=("Segoe UI",9,"bold"))

    def project_data(self):
        return {"settings":{k:v.get() for k,v in self.vars.items()},"rotate":self.rotate.get(),"parts":[asdict(p) for p in self.parts]}

    def save_project(self):
        path=filedialog.asksaveasfilename(defaultextension=".f4cut",filetypes=[("FOUR4 layihəsi","*.f4cut"),("JSON","*.json")])
        if path:
            with open(path,"w",encoding="utf-8") as f:json.dump(self.project_data(),f,ensure_ascii=False,indent=2)
            self.status.set("Layihə yadda saxlanıldı.")

    def open_project(self):
        path=filedialog.askopenfilename(filetypes=[("FOUR4 layihəsi","*.f4cut *.json"),("Bütün fayllar","*.*")])
        if not path:return
        try:
            d=json.load(open(path,encoding="utf-8"))
            for k,v in d["settings"].items():self.vars[k].set(v)
            self.rotate.set(d.get("rotate",True)); self.parts=[Part(**p) for p in d["parts"]]
            self.refresh_tree(); self.sheets=[]; self.canvas.delete("all"); self.summary.config(text="Layihə açıldı. Raskroyu yenidən hesablayın.")
            self.status.set("Layihə açıldı.")
        except Exception as e:messagebox.showerror("Açılmadı",str(e))

    def export_excel(self):
        if not self.sheets:messagebox.showinfo("Məlumat","Əvvəlcə raskroy hesablayın.");return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path:return
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font,PatternFill,Alignment
            wb=Workbook(); ws=wb.active;ws.title="Kəsim siyahısı"
            ws.append(["Lövhə","Material","Qalınlıq","Detal","En (mm)","Uzunluq (mm)","X (mm)","Y (mm)","Döndərilib"])
            for i,s in enumerate(self.sheets,1):
                for p in s.placed:ws.append([i,s.material,s.thickness,p.part.name,p.w,p.h,round(p.x,1),round(p.y,1),"Bəli" if p.rotated else "Xeyr"])
            for c in ws[1]:
                c.font=Font(bold=True,color="FFFFFF");c.fill=PatternFill("solid",fgColor="17324D")
            ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width=min(24,max(12,max(len(str(c.value or "")) for c in col)+2))
            wb.save(path);self.status.set("Excel hesabatı yaradıldı.")
        except ImportError:messagebox.showerror("Kitabxana çatışmır","Excel ixracı üçün Command Prompt-da: pip install openpyxl")
        except Exception as e:messagebox.showerror("Excel xətası",str(e))

    def export_pdf(self):
        if not self.sheets:messagebox.showinfo("Məlumat","Əvvəlcə raskroy hesablayın.");return
        path=filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF","*.pdf")])
        if not path:return
        try:
            from reportlab.lib.pagesizes import A4,landscape
            from reportlab.pdfgen import canvas
            c=canvas.Canvas(path,pagesize=landscape(A4)); W,H=landscape(A4)
            c.setFont("Helvetica-Bold",16);c.drawString(35,H-35,"FOUR4 CUT PRO - KESIM HESABATI")
            c.setFont("Helvetica",9);c.drawString(35,H-52,f"Lövhə ölçüsü: {self.vars['sheet_w'].get()} × {self.vars['sheet_h'].get()} mm | Mişar payı: {self.vars['kerf'].get()} mm")
            y=H-80
            for i,s in enumerate(self.sheets,1):
                if y<100:c.showPage();y=H-45
                c.setFont("Helvetica-Bold",11);c.drawString(35,y,f"Lövhə {i}: {s.material} {s.thickness} mm — {s.w:g} × {s.h:g} mm");y-=17
                c.setFont("Helvetica",8)
                for p in s.placed:
                    c.drawString(45,y,f"{p.part.name}: {p.w:g} × {p.h:g} mm | X={p.x:g}, Y={p.y:g}");y-=12
                y-=10
            c.save();self.status.set("PDF hesabatı yaradıldı.")
        except ImportError:messagebox.showerror("Kitabxana çatışmır","PDF ixracı üçün Command Prompt-da: pip install reportlab")
        except Exception as e:messagebox.showerror("PDF xətası",str(e))

    def print_hint(self):
        messagebox.showinfo("Çap","Əvvəlcə PDF hesabatı yaradın. PDF faylını açaraq çap edə bilərsiniz.")

    def new_project(self):
        if messagebox.askyesno("Yeni layihə","Cari detallar silinsin?"):
            self.parts=[];self.sheets=[];self.refresh_tree();self.sheet_select["values"]=[];self.canvas.delete("all")
            self.summary.config(text="Yeni layihə hazırdır.");self.status.set("Yeni layihə.")

if __name__=="__main__":
    App().mainloop()
