import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass
import os
import json
import base64
import threading

# 로딩 속도 최적화: 대용량 라이브러리는 백그라운드 지연 로딩 또는 전역 선언
try:
    import fitz
except ImportError:
    pass

# =========================================================================================
# ANTIGRAVITY PDF-TO-HTML CONVERTER v34.1.16 (ULTIMATE FAST LOAD EDITION)
# Optimization:
# 1. Data Parsing: Use <script type='application/json'> to avoid JS parser lag on huge strings.
# 2. Priority Rendering: Render Page 1 immediately, then load others in background.
# 3. Binary Handling: Fast Base64 -> Uint8Array conversion.
# =========================================================================================

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>[[TITLE]]</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<style>
    body, html { margin:0; padding:0; height:100%; background:#525659; font-family:'Malgun Gothic',sans-serif; overflow:hidden; }
    .layout { display:flex; height:100%; position:relative; }
    .sb { width: 0px; background: #292929; border-right: 1px solid #444; display: flex; flex-direction: column; color: #ddd; z-index: 100; transition: width 0.3s ease; overflow: hidden; white-space: nowrap; }
    .sb.open { width: 300px; min-width: 300px; }
    .sb-head { padding:22px; background:#0078d4; font-weight:bold; display:flex; justify-content:space-between; align-items:center; font-size: 1.1em; }
    .sb-list { flex:1; overflow-y:auto; padding:10px; }
    .sb-item { padding:14px; background:#333; margin-bottom:10px; border-radius:4px; font-size:15px; cursor:pointer; display:flex; align-items:center; }
    .sb-item:hover { background:#444; color:white; }
    .main { flex:1; overflow-y: auto; display:flex; flex-direction:column; align-items:center; position:relative; scroll-behavior: smooth; background: #525659; scroll-snap-type: y mandatory; }
    .page-box { position: relative; margin: 10px 0; box-shadow: 0 4px 15px rgba(0,0,0,0.5); background: white; scroll-snap-align: start; scroll-snap-stop: always; }
    .page-box.loading { background: #fff; display:flex; align-items:center; justify-content:center; }
    .page-box.loading::before { content: "Loading..."; color: #ccc; font-size: 20px; }
    canvas { display:block; width:100%; height:100%; }
    .layer { position:absolute; inset:0; pointer-events:none; z-index:50; }
    .hit { position:absolute; cursor:pointer; pointer-events:auto; transition: all 0.2s; z-index:60; }
    .highlight-mode .hit.link { background:rgba(0,120,212,0.15); border:2px solid #0078d4; }
    .highlight-mode .hit.pin { background:rgba(255,165,0,0.2); border:2px solid orange; }
    .hit:hover { background:rgba(255,255,0,0.2) !important; box-shadow:0 0 8px rgba(0,0,0,0.2); }
    #tip { position:fixed; background:#ff9; border:1px solid #888; padding:8px 14px; font-size:15px; display:none; pointer-events:none; z-index:9999; color:black; }
    .sb-toggle { position: fixed; top: 50%; left: 33px; transform: translateY(-50%); z-index: 2000; background: rgba(50, 50, 50, 0.9); color: white; border: 1px solid #777; padding: 18px 31px; border-radius: 9px; cursor: pointer; display: flex; align-items: center; gap: 13px; box-shadow: 0 7px 14px rgba(0,0,0,0.4); font-size: 24px; font-weight: bold; transition: all 0.2s; }
    .sb-toggle:hover { background: rgba(60, 60, 60, 1.0); transform: translateY(-50%) scale(1.05); }
    .fab-highlight { position:fixed; bottom:44px; right:44px; width: 110px; height: 110px; border-radius:50%; background:#0078d4; color:white; border:none; box-shadow:0 9px 22px rgba(0,0,0,0.5); cursor:pointer; display:flex; align-items:center; justify-content:center; font-size: 50px; z-index:2000; transition: transform 0.2s; }
    .fab-highlight:hover { transform:scale(1.1); background:#0063b1; }
    .fab-highlight.active { background:#ff8c00; box-shadow:0 0 24px #ff8c00; }
    #loader { position:fixed; inset:0; background:rgb(30,30,30); color:white; display:flex; flex-direction:column; align-items:center; justify-content:center; z-index:10000; transition: opacity 0.5s; }
</style>
</head>
<body>
<div id="loader"><div style="font-size:30px; font-weight:bold; margin-bottom:10px;">Smart Viewer</div><div id="prog">초고속 로딩 중...</div></div>
<div class="layout">
    <div class="sb" id="sidebar">
        <div class="sb-head"><span>[Attachment List]</span><span style="font-size:24px; cursor:pointer;" onclick="toggleSidebar()">X</span></div>
        <div class="sb-list" id="sbList"></div>
    </div>
    <div class="main" id="viewer"></div>
</div>
<button class="sb-toggle" onclick="toggleSidebar()"><span>[List]</span><span>Attachment Files</span></button>
<button class="fab-highlight" id="toggleLinkBtn" title="Toggle Highlight">Eye</button>
<div id="tip"></div>

<!-- Data Isolation for Fast Parsing -->
<script id="pdf-data" type="application/json">
[[JSON_DATA]]
</script>

<script>
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    
    // Lazy Parse JSON
    let DATA = null;  
    
    const body=document.body, linkBtn=document.getElementById('toggleLinkBtn'), tip=document.getElementById('tip'), sidebar=document.getElementById('sidebar');
    let pdfDoc=null, observer=null;

    linkBtn.onclick=()=>{body.classList.toggle('highlight-mode');linkBtn.classList.toggle('active');};
    function toggleSidebar(){sidebar.classList.toggle('open');}
    
    // Optimized Base64 -> Uint8Array
    function base64ToUint8Array(base64) {
        var binary_string = window.atob(base64);
        var len = binary_string.length;
        var bytes = new Uint8Array(len);
        for (var i = 0; i < len; i++) { bytes[i] = binary_string.charCodeAt(i); }
        return bytes;
    }

    window.onresize=()=>updateLayout();
    
    function updateLayout(){
        const container=document.getElementById('viewer');
        const W=container.clientWidth, H=container.clientHeight;
        document.querySelectorAll('.page-box').forEach(box=>{
            const num=parseInt(box.dataset.pageNum);
            const meta=DATA.pages[num-1];
            // Safe fallback
            if(meta) {
                const scale=Math.min((W-40)/meta.width, (H-20)/meta.height);
                box.style.width=(meta.width*scale)+'px';
                box.style.height=(meta.height*scale)+'px';
            }
        });
    }

    async function init(){
        try{
            // 1. Parse Data Efficiently
            const raw=document.getElementById('pdf-data').textContent;
            DATA = JSON.parse(raw);
            
            // 2. Start PDF Loading
            const pdfBytes = base64ToUint8Array(DATA.pdf_base64);
            const loadingTask=pdfjsLib.getDocument({data:pdfBytes});
            
            loadingTask.onProgress=(p)=>{
                if(p.total){document.getElementById('prog').innerText="로딩: "+Math.round((p.loaded/p.total)*100)+"%";}
            };
            
            pdfDoc = await loadingTask.promise;
            
            // 3. Immediate Setup of Sidebar
            const sb=document.getElementById('sbList');
            if(DATA.attachments.length===0)sb.innerHTML="<div style='padding:15px; opacity:0.5'>첨부 없음</div>";
            else {
                // Fragment for speed
                const frag = document.createDocumentFragment();
                DATA.attachments.forEach(a=>{
                    const el=document.createElement('div');el.className='sb-item';
                    el.innerHTML="<span>[FILE]</span>"+a.name; el.onclick=()=>download(a); frag.appendChild(el);
                });
                sb.appendChild(frag);
            }

            const container=document.getElementById('viewer');
            const mainW=container.clientWidth, mainH=container.clientHeight;

            // 4. Setup Observer
            observer=new IntersectionObserver((entries)=>{
                entries.forEach(entry=>{
                    if(entry.isIntersecting){
                        const div=entry.target; 
                        const pageNum=parseInt(div.dataset.pageNum);
                        if(!div.dataset.rendered){
                            renderPage(pageNum, div); 
                            div.dataset.rendered="true";
                        }
                    }
                });
            },{root:container, rootMargin:"200% 0px", threshold:0.01});

            // 5. Render Page 1 IMMEDIATELY (Critical Path)
            await createPagePlaceholder(1, container, mainW, mainH);
            
            // Hide Loader ASAP
            document.getElementById('loader').style.opacity = '0';
            setTimeout(()=>document.getElementById('loader').style.display='none', 500);

            // 6. Defer creation of other pages to keep UI responsive
            if(pdfDoc.numPages > 1) {
                setTimeout(()=>{
                    createRemainingPages(container, mainW, mainH);
                }, 100);
            }

        }catch(e){alert("Error: "+e.message);}
    }

    function createPagePlaceholder(i, container, mainW, mainH) {
        const meta=DATA.pages[i-1]; 
        const box=document.createElement('div');
        box.className='page-box loading'; 
        box.id='page-'+i; 
        box.dataset.pageNum=i;
        
        const scale=Math.min((mainW-40)/meta.width, (mainH-20)/meta.height);
        box.style.width=(meta.width*scale)+'px'; 
        box.style.height=(meta.height*scale)+'px';
        
        container.appendChild(box);
        observer.observe(box);
        return box;
    }

    function createRemainingPages(container, mainW, mainH) {
        const frag = document.createDocumentFragment();
        // Just create div meta, observe later? No, must append to observe.
        // Let's create chunk by chunk if too many
        for(let i=2; i<=pdfDoc.numPages; i++){
            const meta=DATA.pages[i-1];
            const box=document.createElement('div');
            box.className='page-box loading'; 
            box.id='page-'+i; 
            box.dataset.pageNum=i;
            const scale=Math.min((mainW-40)/meta.width, (mainH-20)/meta.height);
            box.style.width=(meta.width*scale)+'px'; 
            box.style.height=(meta.height*scale)+'px';
            container.appendChild(box);
            observer.observe(box);
        }
    }

    async function renderPage(num,container){
        try{
            const page=await pdfDoc.getPage(num);
            const viewport=page.getViewport({scale:[[RENDER_SCALE]]});
            const cvs=document.createElement('canvas'); 
            cvs.width=viewport.width; 
            cvs.height=viewport.height;
            
            await page.render({canvasContext:cvs.getContext('2d'),viewport:viewport}).promise;
            
            container.classList.remove('loading'); 
            container.innerHTML=''; 
            container.appendChild(cvs);
            
            // Interaction Layer
            const layer=document.createElement('div'); layer.className='layer';
            const meta=DATA.pages[num-1];
            
            if(meta.links)meta.links.forEach(l=>{
                const el=document.createElement('div'); el.className='hit link';
                const p=l.rect_pct; 
                el.style.left=(p[0]*100)+'%'; el.style.top=(p[1]*100)+'%';
                el.style.width=((p[2]-p[0])*100)+'%'; el.style.height=((p[3]-p[1])*100)+'%';
                el.onclick=()=>{
                    if(l.uri)window.open(l.uri,'_blank');
                    else if(l.page!==undefined && l.page!==null){
                        const t=document.getElementById('page-'+(l.page+1));
                        if(t)t.scrollIntoView({behavior:'smooth',block:'start'});
                    }
                };
                layer.appendChild(el);
            });
            
            if(meta.pins)meta.pins.forEach(p=>{
                const el=document.createElement('div'); el.className='hit pin';
                const pct=p.rect_pct; 
                el.style.left=(pct[0]*100)+'%'; el.style.top=(pct[1]*100)+'%';
                el.style.width=((pct[2]-pct[0])*100)+'%'; el.style.height=((pct[3]-pct[1])*100)+'%';
                // Click to Download
                el.onclick=(e)=>{e.stopPropagation(); const at=DATA.attachments.find(x=>x.id===p.attId); if(at)download(at);};
                el.onmouseover=(e)=>{tip.style.display='block'; tip.innerHTML="💾 다운로드<br>"+p.name; moveTip(e);};
                el.onmousemove=moveTip; el.onmouseout=()=>tip.style.display='none';
                layer.appendChild(el);
            });
            container.appendChild(layer);
        }catch(e){console.error(e);}
    }
    
    function moveTip(e){tip.style.left=(e.clientX+15)+'px'; tip.style.top=(e.clientY+10)+'px';}
    function download(a){
        const bin=atob(a.data); const arr=new Uint8Array(bin.length);
        for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);
        const blob=new Blob([arr],{type:'application/octet-stream'});
        const url=URL.createObjectURL(blob); const el=document.createElement('a');
        el.href=url; el.download=a.name; el.click(); URL.revokeObjectURL(url);
    }
    
    // Start Init
    setTimeout(init, 10);
</script>
</body>
</html>
"""

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Ultimate PDF Converter")
        self.root.geometry("500x420")
        style = ttk.Style(); style.theme_use('clam')
        
        # 즉시 화면 표시를 위한 강제 업데이트
        self.root.update()
        
        lbl = tk.Label(root, text="[Ultimate PDF to HTML]", font=("Arial", 16, "bold"), fg="#d63031", pady=15)
        lbl.pack()
        
        frame = tk.Frame(root); frame.pack(pady=5, fill='x', padx=20)
        self.file_list = tk.Listbox(frame, height=5, selectmode='extended')
        self.file_list.pack(side='left', fill='both', expand=True)
        sb = tk.Scrollbar(frame, command=self.file_list.yview)
        sb.pack(side='right', fill='y'); self.file_list.config(yscrollcommand=sb.set)
        
        tk.Button(root, text="+ PDF 추가", command=self.select_files).pack(pady=5)
        
        frame_opts = tk.LabelFrame(root, text="Options", padx=10, pady=5)
        frame_opts.pack(pady=10, padx=20, fill='x')
        self.var_quality = tk.BooleanVar(value=True)
        self.var_homepatch = tk.BooleanVar(value=True)
        tk.Checkbutton(frame_opts, text="Hi-DPI", variable=self.var_quality).pack(side='left', padx=10)
        tk.Checkbutton(frame_opts, text="Home Patch", variable=self.var_homepatch).pack(side='left', padx=10)
        
        self.progress = ttk.Progressbar(root, orient='horizontal', length=100, mode='determinate')
        self.progress.pack(fill='x', padx=20, pady=5)
        
        self.btn_run = tk.Button(root, text="변환 시작", command=self.run_conversion, bg="#d63031", fg="white", font=("Arial", 12, "bold"))
        self.btn_run.pack(pady=10, fill='x', padx=50)
        self.lbl_status = tk.Label(root, text="Ready", fg="gray"); self.lbl_status.pack(pady=5)
        self.files = []

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files:
            for f in files: 
                if f not in self.files: self.files.append(f); self.file_list.insert(tk.END, os.path.basename(f))
            self.lbl_status.config(text=f"{len(self.files)} files")

    def run_conversion(self):
        if not self.files: return
        self.btn_run.config(state='disabled'); self.progress['value']=0
        t = threading.Thread(target=self.process_files); t.start()

    def process_files(self):
        try:
            # fitz가 전역에서 로드되지 않았을 경우를 대비한 2차 확인
            global fitz
            if 'fitz' not in globals():
                import fitz
            total = len(self.files)
            render_scale = "3.0" if self.var_quality.get() else "1.5"
            home_patch = self.var_homepatch.get()
            
            for idx, fpath in enumerate(self.files):
                fname = os.path.basename(fpath)
                self.update_ui(f"Processing: {fname}", (idx/total)*100)
                
                doc = fitz.open(fpath)
                data = {"pages": [], "attachments": []}
                saved_att = set()
                
                try: 
                    for i in range(doc.embfile_count()):
                        name = doc.embfile_info(i)["name"]
                        content = doc.embfile_get(i)
                        if name not in saved_att:
                            data["attachments"].append({"id": f"emb_{i}", "name": name, "data": base64.b64encode(content).decode('utf-8')})
                            saved_att.add(name)
                except: pass
                
                for i, page in enumerate(doc):
                    w, h = page.rect.width, page.rect.height
                    p_data = {"num":i+1, "width":w, "height":h, "links":[], "pins":[]}
                    has_home = False
                    for link in page.get_links():
                        r = link["from"]
                        if r.x0 > w*0.9 and r.y1 < h*0.15: has_home = True
                        p_data["links"].append({"rect_pct": [r.x0/w, r.y0/h, r.x1/w, r.y1/h], "uri": link.get("uri"), "page": link.get("page")})
                    if home_patch and i>0 and not has_home:
                        p_data["links"].append({"rect_pct":[0.94,0.0,1.0,0.08], "page":0, "is_virtual":True})
                    for annot in page.annots():
                        if annot.type[0] == 17:
                            try:
                                res = annot.get_file()
                                f_content = b""
                                f_name = "unknown.pdf"
                                if isinstance(res, bytes): f_content = res; f_name = annot.info.get("content") or annot.info.get("title") or f"file_{annot.xref}.pdf"
                                elif isinstance(res, dict): f_content = res.get("content", b""); f_name = res.get("filename") or annot.info.get("content") or "unknown.pdf"
                                if f_content:
                                    att_id = f"annot_{annot.xref}"
                                    existing = next((a for a in data["attachments"] if a["name"]==f_name), None)
                                    if not existing: data["attachments"].append({"id":att_id, "name":f_name, "data":base64.b64encode(f_content).decode('utf-8')}); saved_att.add(f_name); tgt_id = att_id
                                    else: tgt_id = existing["id"]
                                    r = annot.rect
                                    p_data["pins"].append({"attId": tgt_id, "name": f_name, "rect_pct": [r.x0/w, r.y0/h, r.x1/w, r.y1/h]})
                            except: pass
                    data["pages"].append(p_data)
                
                with open(fpath, "rb") as f: data["pdf_base64"] = base64.b64encode(f.read()).decode('utf-8')
                doc.close()
                
                out_name = os.path.splitext(fpath)[0] + ".html"
                final = HTML_TEMPLATE.replace("[[TITLE]]", os.path.basename(out_name)).replace("[[RENDER_SCALE]]", render_scale).replace("[[JSON_DATA]]", json.dumps(data, ensure_ascii=False))
                with open(out_name, "w", encoding="utf-8") as f: f.write(final)
            
            self.update_ui("Done!", 100); messagebox.showinfo("Success", f"{total} processed")
        except Exception as e: messagebox.showerror("Error", str(e))
        finally: self.root.after(0, lambda: self.btn_run.config(state='normal'))

    def update_ui(self, msg, progress):
        self.root.after(0, lambda: self.lbl_status.config(text=msg))
        self.root.after(0, lambda: self.progress.configure(value=progress))

if __name__ == "__main__":
    if sys.platform == 'win32':
        if sys.stdout: sys.stdout.reconfigure(encoding='utf-8')
        if sys.stderr: sys.stderr.reconfigure(encoding='utf-8')
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    app = App(root); root.mainloop()
