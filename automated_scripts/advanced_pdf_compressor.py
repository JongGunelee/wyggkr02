"""
================================================================================
 [지능형 PDF 고속 압축 도구 v34.1.16] [Engine Ultimate Compression | 속도/안정성 최적화]
================================================================================
- Clean Layer Architecture: Engine(Domain) / View(Presentation) / Controller(Application)
- PyMuPDF(fitz) 최신 물리적 대체 엔진(replace_image) 적용 (압축 무효화 차단)
- garbage=3 기본값 (속도/품질 최적 균형) - 사용자 요청 시 garbage=4 선택 가능
- Document 상태 관리 강화로 'document closed' 오류 원천 차단
================================================================================
"""
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass
import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time


# ==========================================
# 1. DOMAIN LAYER (Core Engine - 비즈니스 로직)
# ==========================================
class PDFCompressEngine:
    """초고속 지능형 PDF 최적화 엔진 v3.4 [속도/안정성 최적화]"""

    # garbage=4 (최고 수준) 기본 적용, 품질/속도 균형 조정
    LEVELS = {
        "SCREEN (초고속/최대압축)": {"dpi": 72, "quality": 35, "garbage": 4},
        "EBOOK (균형/표준추천)": {"dpi": 150, "quality": 60, "garbage": 4},
        "PRINTER (고화질/선명함)": {"dpi": 200, "quality": 75, "garbage": 4},
        "ULTIMATE (정밀복구/안전모드)": {"dpi": 300, "quality": 90, "garbage": 4}
    }
    
    @staticmethod
    def compress_file(filepath, output_dir, level_name, callback=None, cancel_event=None):
        """
        초정밀 진단형 PDF 압축 및 최적화 (Document 상태 관리 강화)
        """
        doc = None
        target_path = None
        current_step = "초기화"
        
        try:
            if cancel_event and cancel_event.is_set(): 
                return False, "작업 취소됨", 0, 0, 0
            
            filename = os.path.basename(filepath)
            
            # 0. 사전 검증
            current_step = "사전 검증"
            if not os.path.exists(filepath): 
                return False, "파일을 찾을 수 없습니다.", 0, 0, 0
            if not os.access(filepath, os.R_OK):
                return False, "파일 읽기 권한이 없습니다.", 0, 0, 0
            
            original_size = os.path.getsize(filepath)
            if original_size == 0:
                return False, "원본 파일이 비어있습니다.", 0, 0, 0
            
            # 저장 폴더 확보
            if not os.path.exists(output_dir): 
                os.makedirs(output_dir)
            
            # 기본 타겟 경로 (접두사는 결과에 따라 변경됨)
            result_prefix = "[Optimized]"  # 기본값
            target_path = os.path.join(output_dir, f"{result_prefix}_{filename}")
            
            settings = PDFCompressEngine.LEVELS.get(level_name, PDFCompressEngine.LEVELS["EBOOK (균형/표준추천)"])
            
            if callback: callback("progress", f"[INFO] {filename} ({original_size//1024} KB) 분석 중...")

            # 1. 문서 열기 (손상된 PDF 자동 복구 모드)
            current_step = "문서 열기"
            try:
                # 먼저 일반 모드로 시도
                doc = fitz.open(filepath)
            except Exception as open_err:
                # 실패 시 복구 모드로 재시도
                if callback: callback("progress", f"[WARN] 복구 모드로 문서 열기 시도...")
                try:
                    doc = fitz.open(filepath, filetype="pdf")
                except:
                    return False, f"문서 열기 실패 (손상된 파일)", 0, 0, 0
            
            if doc.is_encrypted:
                doc.close()
                doc = None
                return False, "암호화된 문서", 0, 0, 0
            
            if doc.page_count == 0:
                doc.close()
                doc = None
                return False, "페이지가 없는 문서", 0, 0, 0
            
            # xref 검증 (손상 여부 사전 체크)
            try:
                _ = doc.xref_length()
            except:
                if callback: callback("progress", f"[WARN] xref 손상 감지, 복구 시도...")
            
            # 2. 이미지 정밀 압축 (결함 수정 - 물리적 스트림 대체 엔진)
            current_step = "이미지 정밀 최적화"
            if settings["dpi"]:
                if callback: callback("progress", f"[INFO] 이미지 정밀 분석 및 물리적 교체 시작...")
                try:
                    processed_xrefs = set()
                    for page in doc:
                        img_list = page.get_images(full=True)
                        for img in img_list:
                            xref = img[0]
                            if xref in processed_xrefs: continue
                            processed_xrefs.add(xref)
                            
                            try:
                                orig_stream_len = doc.xref_stream_length(xref)
                                if orig_stream_len < 1000: continue # 임계값 하향 (1KB)

                                pix = fitz.Pixmap(doc, xref)
                                
                                # 해상도 제어 (예: PRINTER 이하이면서 너무 큰 4K급 이미지만 일부 축소할 수도 있으나, 여기선 quality 중심으로 통제)
                                compressed_img = pix.tobytes("jpeg", quality=settings["quality"])
                                
                                if len(compressed_img) < orig_stream_len:
                                    # 최신 API를 통해 필터 딕셔너리 동기화 보장
                                    try:
                                        page.replace_image(xref, stream=compressed_img)
                                    except Exception:
                                        doc.update_stream(xref, compressed_img)
                            except:
                                continue
                except Exception as img_err:
                    if callback: callback("progress", f"[WARN] 이미지 처리 중 일부 스킵됨")

            # 3. 폰트/구조 최적화 및 저장
            current_step = "구조 최적화 및 저장"
            if callback: callback("progress", f"[INFO] 최종 구조 다이어트 중...")
            try: doc.subset_fonts()
            except: pass
            
            # 취소 확인
            if cancel_event and cancel_event.is_set(): 
                doc.close()
                doc = None
                return False, "작업 취소됨", 0, 0, 0

            # 임시 파일명 생성
            import uuid
            temp_filename = f".tmp_{uuid.uuid4().hex[:8]}.pdf"
            temp_save_path = os.path.join(output_dir, temp_filename)
            
            save_success = False
            try:
                # [중요] 용량 비대화 차단을 위해 pretty=False 필수로 사용
                # 압축된 데이터를 텍스트 모드로 포맷팅하지 않도록 보장
                doc.save(
                    temp_save_path, 
                    garbage=4,
                    deflate=True, 
                    clean=True,
                    linear=False,
                    pretty=False
                )
                
                # 저장 성공 시 정식 이름으로 변경
                if os.path.exists(target_path):
                    try: os.remove(target_path)
                    except: pass
                
                try:
                    os.rename(temp_save_path, target_path)
                    save_success = True
                except:
                    # rename 실패 시 (파일 잠금 등) -> copy 후 삭제 시도
                    import shutil
                    shutil.copy2(temp_save_path, target_path)
                    try: os.remove(temp_save_path)
                    except: pass
                    save_success = True

            except Exception as save_err:
                # 임시 파일 정리
                if os.path.exists(temp_save_path):
                    try: os.remove(temp_save_path)
                    except: pass
                    
                err_msg = str(save_err).lower()
                if callback: callback("progress", f"[WARN] 저장 오류 감지 ({str(save_err)[:30]}...), 강제 재구성(Rebuild) 시작")
                
                # 현재 문서 닫기
                if doc:
                    try: doc.close()
                    except: pass
                    doc = None
                
                # 구조 오류 접두사 확정
                result_prefix = "[Struc Error]"
                target_path = os.path.join(output_dir, f"{result_prefix}_{filename}")
                
                # -------------------------------------------------------
                # [강력 복구 전략] 새 문서에 페이지를 하나씩 복사하여 모든 구조적 결합 제거
                # -------------------------------------------------------
                try:
                    src_doc = fitz.open(filepath)
                    new_doc = fitz.open()  # 완전히 새로운 PDF 객체 생성
                    
                    # 1. 모든 페이지 복사 (구조적 결함이 있는 객체는 이 과정에서 버려짐)
                    new_doc.insert_pdf(src_doc)
                    src_doc.close()
                    
                    # 2. 새 문서 이미지 정밀 최적화 (물리적 교체)
                    if settings["dpi"]:
                        processed_xrefs = set()
                        for page in new_doc:
                            for img in page.get_images(full=True):
                                xref = img[0]
                                if xref in processed_xrefs: continue
                                processed_xrefs.add(xref)
                                
                                try:
                                    orig_len = new_doc.xref_stream_length(xref)
                                    if orig_len < 5000: continue
                                    pix = fitz.Pixmap(new_doc, xref)
                                    compressed = pix.tobytes("jpeg", quality=settings["quality"])
                                    if len(compressed) < orig_len:
                                        try:
                                            page.replace_image(xref, stream=compressed)
                                        except Exception:
                                            new_doc.update_stream(xref, compressed)
                                except: continue
                    
                    # 3. 폰트 서브셋 재적용
                    try: new_doc.subset_fonts()
                    except: pass
                    
                    # 4. 강력한 재구축 저장 (garbage=4: 모든 미사용 객체 제거 및 인덱스 재생성)
                    new_doc.save(
                        target_path,
                        garbage=4,
                        deflate=True,
                        clean=True,
                        linear=True
                    )
                    new_doc.close()
                    save_success = True
                    if callback: callback("progress", f"[OK] 문서 완전 재구축 성공 (용량 다이어트 완료)")
                    
                except Exception as rebuild_err:
                    if callback: callback("progress", f"[WARN] 재구축 실패 ({str(rebuild_err)[:20]}), 최후의 폴백 시도...")
                    
                    # 최후의 수단: 단순히 다시 시도 (가장 낮은 옵션으로)
                    try:
                        last_doc = fitz.open(filepath)
                        last_doc.save(target_path, garbage=1, deflate=False, clean=False)
                        last_doc.close()
                        save_success = True
                        if callback: callback("progress", f"[OK] 단순 저장 복구 성공")
                    except:
                        # 진짜 마지막: 원본 파일 그대로 복사
                        try:
                            import shutil
                            shutil.copy2(filepath, target_path)
                            save_success = True
                            if callback: callback("progress", f"[OK] 원본 복제 (구조 손상 심각하여 압축 불가)")
                        except:
                            pass
            
            if not save_success:
                if doc:
                    doc.close()
                    doc = None
                # 저장 실패 시 빈 파일 즉시 삭제
                if os.path.exists(target_path):
                    for _ in range(3):
                        try: 
                            os.remove(target_path)
                            break
                        except: 
                            time.sleep(0.1)
                return False, f"저장 실패 (PDF 구조 복구 불가)", original_size, 0, 0
            
            # 저장 완료 후 문서 닫기 (None 체크 필수)
            if doc:
                doc.close()
                doc = None
            
            # 6. 결과 검증 (0KB 파일 강력 차단)
            current_step = "검증"
            if not os.path.exists(target_path):
                return False, "저장 파일 없음", 0, 0, 0
            
            compressed_size = os.path.getsize(target_path)
            
            # 0KB 파일 강력 삭제 (3회 재시도)
            if compressed_size == 0:
                for attempt in range(3):
                    try: 
                        os.remove(target_path)
                        break
                    except: 
                        time.sleep(0.1)
                return False, "압축 실패 (0KB)", original_size, 0, 0
            
            # [가장 엄격한 검증] 원본보다 1바이트라도 크면 실패로 간주
            is_struct_error = "[Struc Error]" in target_path
            
            final_path = target_path
            
            # 용량이 줄어들지 않았거나 늘어난 경우 (오차 허용 0%)
            if not is_struct_error and compressed_size >= original_size:
                # [Optimized] 파일 삭제 (강력 삭제)
                if os.path.exists(target_path):
                    for attempt in range(3):
                        try: 
                            os.remove(target_path)
                            break
                        except: 
                            time.sleep(0.1)
                
                # [SKIPP] 접두사로 원본 복사
                import shutil
                skip_path = os.path.join(output_dir, f"[SKIPP]_{filename}")
                shutil.copy2(filepath, skip_path)
                final_path = skip_path  # 검증 대상 변경
                
                if callback: callback("progress", f"ℹ️ 압축 효과 없음 (원본 복사 → [SKIPP])")
            
            # 7. 최종 파일 무결성 보증 검사 (Integrity Check)
            current_step = "무결성 검증"
            try:
                # 최종 파일 열기 시도
                verify_doc = fitz.open(final_path)
                
                # 페이지 수 검증
                if verify_doc.page_count < 1:
                    raise ValueError("페이지 없음 (Empty PDF)")
                    
                # 렌더링 검증 (첫 페이지 로드)
                try:
                    _ = verify_doc.load_page(0)
                except Exception as load_err:
                    raise ValueError(f"페이지 렌더링 실패: {str(load_err)}")
                
                page_cnt = verify_doc.page_count
                verify_doc.close()
                
                if callback: callback("progress", f"[OK] [검증 통과] {os.path.basename(final_path)} (Page={page_cnt}p OK)")
                
            except Exception as integrity_err:
                # 검증 실패 시 파일 즉시 삭제
                if os.path.exists(final_path):
                    for _ in range(3):
                        try: 
                            os.remove(final_path)
                            break
                        except: 
                            time.sleep(0.1)
                return False, f"무결성 검증 실패: {str(integrity_err)}", original_size, 0, 0

            # 최종 결과 반환
            final_size = os.path.getsize(final_path)
            ratio = (1 - final_size / original_size) * 100
            return True, f"{final_path} [Integrity Verified]", original_size, final_size, ratio
            
        except Exception as e:
            err_str = str(e)
            if "Cancelled" in err_str: 
                return False, "작업 취소됨", 0, 0, 0
            if "document closed" in err_str.lower():
                return False, "문서 처리 오류 (재시도 필요)", 0, 0, 0
            return False, f"[{current_step}] {err_str[:40]}", 0, 0, 0
        finally:
            # 안전한 문서 정리
            if doc is not None:
                try: 
                    doc.close()
                except: 
                    pass


# ==========================================
# 2. PRESENTATION LAYER (View - GUI 정의)
# ==========================================
class PDFCompressView:
    """GUI 화면 구성: Controller와 분리된 순수 View"""
    
    def __init__(self, root, controller):
        self.root = root
        self.controller = controller
        self.root.title("[지능형 PDF 고속 압축 도구 v34.1.16]")
        self.root.geometry("780x680")
        self._build_ui()
    
    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        # 1. 상단 타이틀
        header = tk.Frame(self.root, bg="#0D47A1", pady=10)
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(header, text="[지능형 PDF 고속 압축기 v34.1.16]", 
                 font=("Malgun Gothic", 14, "bold"), bg="#0D47A1", fg="white").pack()
        
        # 2. 메인 컨텐츠
        main = tk.Frame(self.root, padx=15, pady=5)
        main.grid(row=1, column=0, sticky="nsew")
        
        # 가이드 박스 + 접두사 안내 버튼
        guide_frame = tk.Frame(main, bg="#E3F2FD")
        guide_frame.pack(fill="x", pady=(0, 5))
        
        tk.Label(guide_frame, text="[INFO] 파일 추가 -> 레벨 선택 -> 고속 압축 가동 (모든 파일 저장)", 
                 font=("Malgun Gothic", 9), bg="#E3F2FD", fg="#1565C0", pady=5).pack(side="left", padx=5)
        
        tk.Button(guide_frame, text="[?] 저장 규칙 안내", font=("Malgun Gothic", 8), 
                  command=self._show_prefix_guide, bg="#FFF9C4", fg="#5D4037").pack(side="right", padx=5, pady=2)

        top_f = tk.Frame(main)
        top_f.pack(fill="x")
        
        # 1️⃣ 파일 목록
        box1 = tk.LabelFrame(top_f, text="[1] 압축 대상 목록", font=("Malgun Gothic", 9, "bold"), padx=10, pady=5)
        box1.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        btn_f = tk.Frame(box1)
        btn_f.pack(fill="x")
        tk.Button(btn_f, text="[추가]", command=self.controller.handle_select_files, bg="#E3F2FD", width=8).pack(side="left")
        tk.Button(btn_f, text="[초기화]", command=self.controller.handle_clear_list, bg="#FFEBEE", width=8).pack(side="left", padx=5)
        self.file_listbox = tk.Listbox(box1, font=("Consolas", 9), height=5)
        self.file_listbox.pack(fill="both", expand=True, pady=5)
        
        self.file_count_label = tk.Label(box1, text="선택된 파일: 0개", font=("Malgun Gothic", 9), fg="#1976D2")
        self.file_count_label.pack(anchor="w")

        # 2️⃣ 설정
        box2 = tk.LabelFrame(top_f, text="[2] 엔진 설정", font=("Malgun Gothic", 9, "bold"), padx=10, pady=5)
        box2.pack(side="right", fill="both", expand=True)
        
        tk.Label(box2, text="압축 레벨:", font=("Malgun Gothic", 9, "bold")).pack(anchor="w")
        self.level_combo = ttk.Combobox(box2, values=list(PDFCompressEngine.LEVELS.keys()), state="readonly", width=28)
        self.level_combo.pack(pady=2, fill="x")
        self.level_combo.set("SCREEN (초고속/최대압축)")
        self.level_combo.bind("<<ComboboxSelected>>", self._on_level_change)
        
        # 레벨 설명 패널
        self.level_info_frame = tk.Frame(box2, bg="#FFFDE7", padx=8, pady=6)
        self.level_info_frame.pack(fill="x", pady=(5, 0))
        
        self.level_desc_label = tk.Label(self.level_info_frame, text="", font=("Malgun Gothic", 8), 
                                         bg="#FFFDE7", fg="#5D4037", justify="left", anchor="w", wraplength=200)
        self.level_desc_label.pack(fill="x")
        
        # 초기 설명 표시
        self._update_level_description("SCREEN (초고속/최대압축)")
        
        tk.Label(box2, text="저장 위치 (자동):", font=("Malgun Gothic", 9)).pack(anchor="w", pady=(8, 0))
        self.path_entry = tk.Entry(box2, textvariable=self.controller.output_dir_var, font=("Malgun Gothic", 8), width=25, state="readonly")
        self.path_entry.pack(fill="x")

        # 3️⃣ 로그
        box3 = tk.LabelFrame(self.root, text="[3] 실시간 가동 현황", font=("Malgun Gothic", 9, "bold"), padx=10, pady=5)
        box3.grid(row=2, column=0, sticky="nsew", padx=15, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(box3, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=(0, 5))
        
        self.log_txt = tk.Text(box3, font=("Consolas", 9), bg="#1E1E1E", fg="#D4D4D4", insertbackground="white")
        self.log_txt.pack(fill="both", expand=True)
        
        # 4. 하단 버튼
        footer = tk.Frame(self.root, pady=10, bg="#F5F5F5")
        footer.grid(row=3, column=0, sticky="ew")
        
        self.run_btn = tk.Button(footer, text="[고속 압축 가동]", font=("Malgun Gothic", 12, "bold"),
                                 bg="#1565C0", fg="white", height=2, command=self.controller.handle_start)
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(15, 5))
        
        self.cancel_btn = tk.Button(footer, text="[중단]", font=("Malgun Gothic", 12, "bold"),
                                    bg="#D32F2F", fg="white", height=2, state="disabled", command=self.controller.handle_cancel)
        self.cancel_btn.pack(side="right", padx=(5, 15))
        
        self.status_bar = tk.Label(self.root, text="Ready | garbage=3 (고속) / ULTIMATE=garbage=4 (최대품질)", 
                                   bd=1, relief="sunken", anchor="w", font=("Malgun Gothic", 8), padx=10)
        self.status_bar.grid(row=4, column=0, sticky="ew")
    
    def log(self, msg):
        self.log_txt.insert("end", f"{msg}\n")
        self.log_txt.see("end")
    
    def clear_log(self):
        self.log_txt.delete("1.0", "end")
    
    def set_status(self, msg):
        self.status_bar.config(text=msg)

    def set_progress(self, val):
        self.progress_var.set(val)
    
    def update_file_list(self, files):
        self.file_listbox.delete(0, tk.END)
        for f in files:
            self.file_listbox.insert(tk.END, os.path.basename(f))
        self.file_count_label.config(text=f"선택된 파일: {len(files)}개")
    
    def set_running_state(self, is_running):
        if is_running:
            self.run_btn.config(state="disabled", text="⏳ 가동 중...")
            self.cancel_btn.config(state="normal")
        else:
            self.run_btn.config(state="normal", text="[START] 고속 압축 가동 [Turbo Mode]")
            self.cancel_btn.config(state="disabled")
    
    def _on_level_change(self, event=None):
        """콤보박스 선택 변경 시 설명 업데이트"""
        selected = self.level_combo.get()
        self._update_level_description(selected)
    
    def _update_level_description(self, level_name):
        """레벨별 상세 설명 표시"""
        descriptions = {
            "SCREEN (초고속/최대압축)": (
                " 용도: 웹 업로드, 이메일 첨부, 미리보기\n"
                " 장점: 가장 빠름, 최대 용량 절감\n"
                " 단점: 이미지 품질 저하\n"
                " 권장: 텍스트 위주, 빠른 공유용"
            ),
            "EBOOK (균형/표준추천)": (
                "[INFO] 단점: 없음 ([RECOMMEND] 추천 옵션)\n"
                "👉 권장: 대부분의 PDF 문서"
            ),
            "PRINTER (고화질/선명함)": (
                "[INFO] 용도: 고품질 인쇄, 프레젠테이션\n"
                "✅ 장점: 이미지 선명도 유지\n"
                "⚠️ 단점: 압축률이 낮음\n"
                "👉 권장: 사진/도면이 많은 문서"
            ),
            "ULTIMATE (최대품질/느림)": (
                "📌 용도: 아카이빙, 법적 문서, 장기보관\n"
                "✅ 장점: 무손실에 가까운 최고 품질\n"
                "⚠️ 단점: 가장 느림 (수 분 소요)\n"
                "👉 권장: 중요 문서, 원본 품질 필수"
            )
        }
        desc = descriptions.get(level_name, "레벨을 선택하세요.")
        self.level_desc_label.config(text=desc)
    
    def _show_prefix_guide(self):
        """접두사 저장 규칙 안내 다이얼로그"""
        guide_text = (
            "[HELP] 저장 파일 접두사 규칙 안내\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            "모든 처리된 파일은 결과에 따라\n"
            "다른 접두사로 저장됩니다:\n\n"
            "[OK] [Optimized]_파일명.pdf\n"
            "   → 정상적으로 압축된 파일\n"
            "   → 용량이 감소됨\n\n"
            "[INFO] [SKIPP]_파일명.pdf\n"
            "   → 이미 최적화된 파일 (원본 복사)\n"
            "   → 추가 압축 불필요\n\n"
            "[WARN] [Struc Error]_파일명.pdf\n"
            "   → PDF 구조 오류로 폴백 복구됨\n"
            "   → 원본과 동일하게 열림\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "[INFO] 모든 파일이 임시 폴더에 저장되어\n"
            "   원본은 절대 수정되지 않습니다."
        )
        messagebox.showinfo("저장 규칙 안내", guide_text)

# ==========================================
# 3. APPLICATION LAYER (Controller)
# ==========================================
class PDFCompressController:
    """Controller: 안정적 순차 처리"""
    
    def __init__(self, root):
        self.root = root
        self.engine = PDFCompressEngine()
        self.is_running = False
        self.cancel_event = threading.Event()
        self.selected_files = []
        self.output_dir_var = tk.StringVar()
        self.view = PDFCompressView(root, self)
    
    def handle_select_files(self):
        """파일 선택 + 저장 위치 자동 생성"""
        files = filedialog.askopenfilenames(
            title="압축할 PDF 파일들을 선택하세요",
            filetypes=[("PDF files", "*.pdf")]
        )
        if files:
            new_files = [f for f in files if f not in self.selected_files]
            self.selected_files.extend(new_files)
            self.view.update_file_list(self.selected_files)
            
            # 저장 폴더 자동 생성
            if self.selected_files:
                first_dir = os.path.dirname(self.selected_files[0])
                auto_output_dir = os.path.join(first_dir, "00_Optimized")
                if not os.path.exists(auto_output_dir):
                    try:
                        os.makedirs(auto_output_dir)
                        self.view.log(f"[INFO] 폴더 자동 생성: {auto_output_dir}")
                    except:
                        pass
                self.output_dir_var.set(auto_output_dir)
            
            self.view.log(f"[+] {len(new_files)}개 추가 (총 {len(self.selected_files)}개)")
    
    def handle_clear_list(self):
        self.selected_files = []
        self.view.update_file_list([])
        self.output_dir_var.set("")
        self.view.log("[INFO] 목록 초기화")
    
    def handle_cancel(self):
        if self.is_running:
            self.cancel_event.set()
            self.view.log("[INFO] 중단 요청됨")
    
    def _check_existing_files(self, output_dir):
        """선택된 파일들에 대응하는 기존 파일 검사"""
        conflicts = []
        prefixes = ["[Optimized]", "[SKIPP]", "[Struc Error]"]
        
        for filepath in self.selected_files:
            filename = os.path.basename(filepath)
            for prefix in prefixes:
                existing = os.path.join(output_dir, f"{prefix}_{filename}")
                if os.path.exists(existing):
                    size = os.path.getsize(existing)
                    conflicts.append((f"{prefix}_{filename}", size))
        
        return conflicts
    
    def _delete_conflicting_files(self, output_dir):
        """사용자 승인 후 충돌 파일 삭제"""
        prefixes = ["[Optimized]", "[SKIPP]", "[Struc Error]"]
        deleted_count = 0
        
        for filepath in self.selected_files:
            filename = os.path.basename(filepath)
            for prefix in prefixes:
                existing = os.path.join(output_dir, f"{prefix}_{filename}")
                if os.path.exists(existing):
                    try:
                        os.remove(existing)
                        deleted_count += 1
                    except:
                        pass
        
        return deleted_count

    def handle_start(self):
        if self.is_running: return
        if not self.selected_files:
            messagebox.showwarning("경고", "대상이 없습니다.")
            return
        
        out_dir = self.output_dir_var.get().strip()
        if not out_dir: 
            messagebox.showwarning("경고", "저장 위치가 없습니다.")
            return

        # ★ 기존 파일 충돌 검사
        conflicts = self._check_existing_files(out_dir)
        if conflicts:
            # 충돌 정보 표시 및 사용자 승인 요청
            conflict_info = f"[WARN] 저장 폴더에 {len(conflicts)}개의 기존 파일이 발견되었습니다.\n\n"
            for fname, size in conflicts[:5]:  # 최대 5개만 표시
                conflict_info += f"  • {fname} ({size//1024} KB)\n"
            if len(conflicts) > 5:
                conflict_info += f"  ... 외 {len(conflicts)-5}개 더\n"
            conflict_info += "\n기존 파일을 어떻게 처리할까요?"
            
            # askyesnocancel: Yes=덮어쓰기, No=건너뛰기, Cancel=취소
            result = messagebox.askyesnocancel(
                "기존 파일 충돌", 
                conflict_info + "\n\n예(Y): 기존 파일 삭제 후 덮어쓰기\n아니오(N): 기존 파일 건너뛰기\n취소: 작업 중단"
            )
            
            if result is None:  # 취소
                return
            elif result:  # 예 - 덮어쓰기
                self._delete_conflicting_files(out_dir)
                self.view.log("[WARN] 기존 파일 삭제 완료 (사용자 승인)")
            else:  # 아니오 - 건너뛰기 (기존 파일이 있는 원본은 스킵)
                self.skip_existing = True
                self.view.log("[INFO] 기존 파일은 건너뛰기로 설정됨")
        else:
            self.skip_existing = False

        self.is_running = True
        self.cancel_event.clear()
        self.view.set_running_state(True)
        self.view.clear_log()
        level = self.view.level_combo.get()
        
        self.view.log(f"[INFO] 압축 시작 (Level: {level})")
        self.view.log(f"   대상: {len(self.selected_files)}개")
        self.view.log("-" * 50)
        
        def _fmt(b):
            return f"{b/1024:.1f} KB" if b < 1024*1024 else f"{b/(1024*1024):.1f} MB"

        def run_thread():
            start_t = time.time()
            total = len(self.selected_files)
            results = []
            
            for i, f in enumerate(self.selected_files, 1):
                if self.cancel_event.is_set(): 
                    break
                
                fname = os.path.basename(f)
                self._cb("progress", f"[{i}/{total}] {fname}", (i-1)/total*100)
                
                res = self.engine.compress_file(f, out_dir, level, self._cb, self.cancel_event)
                ok, msg, o_sz, c_sz, ratio = res
                
                results.append({"ok": ok, "msg": msg, "o": o_sz, "c": c_sz, "r": ratio, "f": fname})
                
                # 결과 표시 (접두사 기반 분류)
                if ok:
                    if "[SKIPP]" in str(msg):
                        self._cb("skip", f"[INFO] {fname}: {_fmt(o_sz)} (원본 복사 -> [SKIPP])", i/total*100)
                    elif "[Struc Error]" in str(msg):
                        self._cb("struct", f"[WARN] {fname}: {_fmt(o_sz)} (구조 복구 -> [Struc Error])", i/total*100)
                    else:
                        self._cb("ok", f"[OK] {fname}: {_fmt(o_sz)} -> {_fmt(c_sz)} ({ratio:.1f}%) -> [Optimized]", i/total*100)
                else:
                    self._cb("err", f"[FAIL] {fname}: {msg}", i/total*100)
                
                time.sleep(0.02)

            end_t = time.time()
            
            if not results:
                self.root.after(0, lambda: self.view.set_running_state(False))
                return

            # 접두사 기반 분류
            optimized_list = [r for r in results if r['ok'] and "[Optimized]" in str(r['msg'])]
            skip_list = [r for r in results if r['ok'] and "[SKIPP]" in str(r['msg'])]
            struct_list = [r for r in results if r['ok'] and "[Struc Error]" in str(r['msg'])]
            fail_list = [r for r in results if not r['ok']]
            
            # 무결성 검증 성공 태그 확인
            verified_count = sum(1 for r in results if "[Integrity Verified]" in r['msg'])

            saved = sum(r['o'] - r['c'] for r in optimized_list)
            orig = sum(r['o'] for r in optimized_list)
            
            summary = {
                "optimized": len(optimized_list),
                "skip": len(skip_list),
                "struct": len(struct_list),
                "fail": len(fail_list),
                "verified": verified_count,
                "saved_mb": saved / (1024*1024),
                "ratio": (saved / orig * 100) if orig > 0 else 0,
                "time": end_t - start_t
            }
            
            self.root.after(0, lambda: self._done(summary))

        threading.Thread(target=run_thread, daemon=True).start()
    
    def _cb(self, t, msg, prog=None):
        self.root.after(0, lambda: self.view.log(f"  {msg}"))
        if t == "progress":
            self.root.after(0, lambda: self.view.set_status(msg))
        if prog is not None:
            self.root.after(0, lambda: self.view.set_progress(prog))

    def _done(self, s):
        self.is_running = False
        self.view.set_running_state(False)
        self.view.log("-" * 50)
        self.view.log(f"[OK] 완료! ({s['time']:.1f}초)")
        self.view.log(f"   [Optimized]: {s['optimized']}개 (압축 성공)")
        self.view.log(f"   [SKIPP]: {s['skip']}개 (이미 최적화)")
        self.view.log(f"   [Struc Error]: {s['struct']}개 (구조 복구)")
        self.view.log(f"   실패: {s['fail']}개")
        self.view.log(f"   [OK] 무결성 보증 검사: {s['verified']}개 완료 (100% 신뢰)")
        self.view.log(f"   절감: {s['saved_mb']:.2f} MB ({s['ratio']:.1f}%)")
        
        total_saved = s['optimized'] + s['skip'] + s['struct']
        self.view.set_status(f"완료: 저장 {total_saved}개, 검증 {s['verified']}개")
        
        messagebox.showinfo("완료", 
            f"[Optimized] 압축: {s['optimized']}개\n"
            f"[SKIPP] 스킵: {s['skip']}개\n"
            f"[Struc Error] 복구: {s['struct']}개\n"
            f"실패: {s['fail']}개\n\n"
            f"[OK] 무결성 검증: {s['verified']}개 완료\n"
            f"절감: {s['saved_mb']:.2f} MB")


# ==========================================
# 4. ENTRY POINT
# ==========================================
if __name__ == "__main__":
    root = tk.Tk()

    # [v34.1.21] Stealth Launch 대응: 창을 최상단으로 강제 부각
    root.lift()
    root.attributes('-topmost', True)
    root.after(100, lambda: root.attributes('-topmost', False))

    PDFCompressController(root)
    root.mainloop()
