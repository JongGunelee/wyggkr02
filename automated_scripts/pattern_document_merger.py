# -*- coding: utf-8 -*-
"""
================================================================================
 [MERGER] 패턴 기반 문서 통합 병합기 (Pattern Document Merger)
 버전: v34.1.32 (Force Cleanup & Direct Binding)
 마지막 업데이트: 2026-03-11
================================================================================
 - 권장 사항: sys.stdout UTF-8 강제 설정 및 이모지 제거 (CP949 호환)
 - 아키텍처: Clean Architecture (Engine / View / Controller)
 - 무결성: Phase-based Cleanup & Extreme Moniker Bypass (GetObject Armor)
================================================================================
"""
import sys, os
if __name__ == "__main__" or "merger" in __name__:
    print(f"--- [BOOT] Pattern Document Merger v34.1.32 Loaded from {__file__} ---", flush=True)
    print(f"--- [ENV] System Locale: {os.environ.get('LANG', 'Default')}, Python: {sys.version.split()[0]} ---", flush=True)
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass
import sys
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
except: pass

import os
import sys
import re
import time
import subprocess
import threading
import tempfile
import fitz  # PyMuPDF
import shutil
from pathlib import Path
from collections import defaultdict
import pythoncom
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor, as_completed

# ═══════════════════════════════════════════════════════════════════════════════
# LAYER 1: DOMAIN (Engine - 순수 비즈니스 로직)
# - GUI 의존성 0%, 단독 테스트 가능
# ═══════════════════════════════════════════════════════════════════════════════

class PatternMergerEngine:
    """
    이기종 문서(PPT, Excel, PDF)를 패턴(접두사)별로 그룹화하여 PDF로 병합하는 핵심 엔진.
    
    Attributes:
        SUPPORTED_EXTENSIONS: 지원되는 파일 확장자 목록
        OUTPUT_FOLDER_NAME: 병합 결과물이 저장될 하위 폴더명
    """
    
    SUPPORTED_EXTENSIONS = {'.pdf', '.pptx', '.ppt', '.xlsx', '.xls'}
    OUTPUT_FOLDER_NAME = "00_Merged_PDFs"
    
    def __init__(self):
        self._xl_app = None
        self._ppt_app = None
        self._temp_dir = None
        self._lock = threading.Lock()  # COM 객체 접근 동기화용
    
    def scan_files(self, source: str, recursive: bool = True) -> list:
        """
        소스(파일 경로 또는 폴더 경로)에서 지원되는 문서 파일을 스캔합니다.
        
        Args:
            source: 파일 또는 폴더 경로 (다중 파일 시 ';'로 구분)
            recursive: 하위 폴더 탐색 여부
            
        Returns:
            발견된 파일 경로 리스트 (Path 객체)
        """
        files = []
        
        # 다중 파일 선택 지원 (';' 구분자)
        if ';' in source:
            paths = [Path(p.strip()) for p in source.split(';') if p.strip()]
            for p in paths:
                if p.is_file() and p.suffix.lower() in self.SUPPORTED_EXTENSIONS:
                    files.append(p)
            return sorted(files, key=lambda x: x.name.lower())
        
        source_path = Path(source)
        
        if source_path.is_file():
            if source_path.suffix.lower() in self.SUPPORTED_EXTENSIONS:
                return [source_path]
            return []
        
        if source_path.is_dir():
            if recursive:
                for root, _, filenames in os.walk(source_path):
                    for filename in filenames:
                        file_path = Path(root) / filename
                        if file_path.suffix.lower() in self.SUPPORTED_EXTENSIONS:
                            files.append(file_path)
            else:
                for item in source_path.iterdir():
                    if item.is_file() and item.suffix.lower() in self.SUPPORTED_EXTENSIONS:
                        files.append(item)
        
        return sorted(files, key=lambda x: x.name.lower())
    
    def detect_patterns(self, files: list, mode: str = 'number', delimiter: str = '_', custom_pattern: str = None) -> dict:
        """
        파일명에서 접두사 패턴을 추출하고 그룹화합니다.
        
        Args:
            files: 파일 경로 리스트
            mode: 'number' (연속 숫자), 'delimiter' (구분자 기준), 'auto' (자동 감지), 'regex' (정규식)
            delimiter: 구분자 모드 시 사용할 구분자
            custom_pattern: regex 모드 시 사용할 정규식 패턴
            
        Returns:
            dict[str, list]: {'01': [file1, file2], '02': [file3], ...}
        """
        groups = defaultdict(list)
        
        for file_path in files:
            filename = file_path.stem  # 확장자 제외한 파일명
            
            if mode == 'number':
                # 파일명 시작부의 연속된 숫자 추출
                match = re.match(r'^(\d+)', filename)
                prefix = match.group(1) if match else 'UNGROUPED'
            elif mode == 'delimiter':
                # 구분자 이전까지 추출
                parts = filename.split(delimiter, 1)
                prefix = parts[0] if len(parts) > 1 else 'UNGROUPED'
            elif mode == 'regex' and custom_pattern:
                # 사용자 정규식 패턴 적용
                try:
                    match = re.match(custom_pattern, filename)
                    prefix = match.group(1) if match and match.groups() else 'UNGROUPED'
                except:
                    prefix = 'UNGROUPED'
            else:
                # auto 모드: 숫자 우선, 실패 시 첫 번째 토큰
                match = re.match(r'^(\d+)', filename)
                if match:
                    prefix = match.group(1)
                else:
                    # 공통 구분자로 토큰화 시도
                    for sep in ['_', '-', ' ', '.']:
                        if sep in filename:
                            parts = filename.split(sep, 1)
                            if parts[0]:
                                prefix = parts[0]
                                break
                    else:
                        prefix = 'UNGROUPED'
            
            groups[prefix].append(file_path)
        
        # 그룹 내 파일들을 이름순 정렬
        for prefix in groups:
            groups[prefix] = sorted(groups[prefix], key=lambda x: x.name.lower())
        
        return dict(sorted(groups.items()))
    
    def auto_detect_patterns(self, files: list) -> dict:
        """
        파일 목록을 심층 분석하여 접두사/접미사/중앙 패턴을 자동으로 탐지합니다.
        
        분석 알고리즘:
        1. 접두사(Prefix) 패턴: 파일명 시작 부분의 공통 패턴
        2. 접미사(Suffix) 패턴: 파일명 끝 부분의 공통 패턴
        3. 중앙(Middle) 패턴: 파일명 중간에 반복되는 토큰
        4. 구분자 기반 토큰화 분석
        
        Args:
            files: 파일 경로 리스트
            
        Returns:
            dict: {
                'pattern_position': str,  # 'prefix', 'suffix', 'middle'
                'detected_mode': str,
                'detected_delimiter': str,
                'patterns_found': list,
                'group_preview': dict,
                'analysis': str
            }
        """
        if not files:
            return {'pattern_position': 'prefix', 'detected_mode': 'number', 'analysis': '분석할 파일이 없습니다.'}
        
        filenames = [f.stem for f in files]  # 확장자 제외
        analysis_lines = []
        
        # ═══════════════════════════════════════════════════════════════
        # 1. 접두사(PREFIX) 패턴 분석
        # ═══════════════════════════════════════════════════════════════
        prefix_results = self._analyze_prefix_patterns(filenames)
        
        # ═══════════════════════════════════════════════════════════════
        # 2. 접미사(SUFFIX) 패턴 분석
        # ═══════════════════════════════════════════════════════════════
        suffix_results = self._analyze_suffix_patterns(filenames)
        
        # ═══════════════════════════════════════════════════════════════
        # 3. 중앙(MIDDLE) 패턴 분석
        # ═══════════════════════════════════════════════════════════════
        middle_results = self._analyze_middle_patterns(filenames)
        
        # ═══════════════════════════════════════════════════════════════
        # 4. 최적 패턴 위치 결정 (점수 기반)
        # ═══════════════════════════════════════════════════════════════
        scores = {
            'prefix': prefix_results.get('score', 0),
            'suffix': suffix_results.get('score', 0),
            'middle': middle_results.get('score', 0)
        }
        
        best_position = max(scores, key=scores.get)
        
        # 분석 결과 텍스트 생성
        analysis_lines.append("[INFO] 심층 패턴 분석 결과")
        analysis_lines.append("=" * 50)
        
        # 접두사 분석 결과
        analysis_lines.append(f"\n[-] 접두사(Prefix) 분석: 점수 {prefix_results['score']:.0f}점")
        if prefix_results['patterns']:
            analysis_lines.append(f"   발견: {', '.join(prefix_results['patterns'][:8])}")
        
        # 접미사 분석 결과
        analysis_lines.append(f"\n[-] 접미사(Suffix) 분석: 점수 {suffix_results['score']:.0f}점")
        if suffix_results['patterns']:
            analysis_lines.append(f"   발견: {', '.join(suffix_results['patterns'][:8])}")
        
        # 중앙 패턴 분석 결과
        analysis_lines.append(f"\n[-] 중앙(Middle) 분석: 점수 {middle_results['score']:.0f}점")
        if middle_results['patterns']:
            analysis_lines.append(f"   발견: {', '.join(middle_results['patterns'][:8])}")
        
        # 최종 권장 위치 안내
        position_names = {'prefix': '접두사', 'suffix': '접미사', 'middle': '중앙'}
        analysis_lines.append(f"\n[OK] 권장 패턴 위치: {position_names[best_position]} ({scores[best_position]:.0f}점)")
        
        # 해당 위치의 결과 사용
        if best_position == 'prefix':
            result_data = prefix_results
        elif best_position == 'suffix':
            result_data = suffix_results
        else:
            result_data = middle_results
        
        # 그룹화 미리보기 생성
        group_preview = self._create_group_preview(files, best_position, result_data)
        
        analysis_lines.append(f"\n[-] 예상 그룹 수: {len(group_preview)}개")
        
        return {
            'pattern_position': best_position,
            'detected_mode': result_data.get('mode', 'auto'),
            'detected_delimiter': result_data.get('delimiter', '_'),
            'patterns_found': result_data.get('patterns', []),
            'group_preview': group_preview,
            'scores': scores,
            'analysis': '\n'.join(analysis_lines)
        }
    
    def _analyze_prefix_patterns(self, filenames: list) -> dict:
        """접두사 패턴을 분석합니다."""
        if not filenames:
            return {'score': 0, 'patterns': [], 'mode': 'number'}
        
        # 숫자 접두사 분석
        number_prefixes = []
        for name in filenames:
            match = re.match(r'^(\d+)', name)
            if match:
                number_prefixes.append(match.group(1))
        
        number_ratio = len(number_prefixes) / len(filenames)
        
        # 구분자 기반 접두사 분석
        delimiter_results = {}
        for sep in ['_', '-', ' ', '.']:
            tokens = []
            for name in filenames:
                if sep in name:
                    token = name.split(sep, 1)[0]
                    if token:
                        tokens.append(token)
            if tokens:
                unique_tokens = set(tokens)
                delimiter_results[sep] = {
                    'ratio': len(tokens) / len(filenames),
                    'tokens': list(unique_tokens),
                    'group_count': len(unique_tokens)
                }
        
        # 점수 계산 (0-100)
        score = 0
        patterns = []
        mode = 'auto'
        delimiter = '_'
        
        if number_ratio >= 0.8:
            score = number_ratio * 100
            patterns = sorted(set(number_prefixes))
            mode = 'number'
        elif delimiter_results:
            best = max(delimiter_results.items(), key=lambda x: x[1]['ratio'])
            if best[1]['ratio'] >= 0.5:
                score = best[1]['ratio'] * 80
                patterns = best[1]['tokens']
                mode = 'delimiter'
                delimiter = best[0]
        
        return {'score': score, 'patterns': patterns, 'mode': mode, 'delimiter': delimiter}
    
    def _analyze_suffix_patterns(self, filenames: list) -> dict:
        """접미사 패턴을 분석합니다."""
        if not filenames:
            return {'score': 0, 'patterns': [], 'mode': 'suffix'}
        
        # 숫자 접미사 분석
        number_suffixes = []
        for name in filenames:
            match = re.search(r'(\d+)$', name)
            if match:
                number_suffixes.append(match.group(1))
        
        number_ratio = len(number_suffixes) / len(filenames)
        
        # 구분자 기반 접미사 분석
        delimiter_results = {}
        for sep in ['_', '-', ' ', '.']:
            tokens = []
            for name in filenames:
                if sep in name:
                    parts = name.rsplit(sep, 1)
                    if len(parts) > 1 and parts[1]:
                        tokens.append(parts[1])
            if tokens:
                unique_tokens = set(tokens)
                delimiter_results[sep] = {
                    'ratio': len(tokens) / len(filenames),
                    'tokens': list(unique_tokens),
                    'group_count': len(unique_tokens)
                }
        
        # 점수 계산
        score = 0
        patterns = []
        mode = 'suffix_number'
        delimiter = '_'
        
        if number_ratio >= 0.8:
            score = number_ratio * 100
            patterns = sorted(set(number_suffixes))
            mode = 'suffix_number'
        elif delimiter_results:
            best = max(delimiter_results.items(), key=lambda x: x[1]['ratio'])
            if best[1]['ratio'] >= 0.5:
                score = best[1]['ratio'] * 80
                patterns = best[1]['tokens']
                mode = 'suffix_delimiter'
                delimiter = best[0]
        
        return {'score': score, 'patterns': patterns, 'mode': mode, 'delimiter': delimiter}
    
    def _analyze_middle_patterns(self, filenames: list) -> dict:
        """중앙 패턴을 분석합니다."""
        if not filenames:
            return {'score': 0, 'patterns': [], 'mode': 'middle'}
        
        # 모든 파일에서 공통으로 나타나는 토큰 추출
        all_tokens = defaultdict(int)
        for name in filenames:
            # 다양한 구분자로 토큰화
            tokens = re.split(r'[_\-\s\.\(\)\[\]]', name)
            for token in tokens[1:-1]:  # 첫번째와 마지막 제외 (중앙만)
                if token and len(token) >= 2:
                    all_tokens[token] += 1
        
        # 여러 파일에서 반복되는 토큰 찾기
        repeated_tokens = {k: v for k, v in all_tokens.items() if v >= 2}
        
        score = 0
        patterns = []
        
        if repeated_tokens:
            # 가장 많이 반복되는 토큰들
            sorted_tokens = sorted(repeated_tokens.items(), key=lambda x: -x[1])
            patterns = [t[0] for t in sorted_tokens[:10]]
            
            # 점수: 가장 많이 반복되는 토큰의 비율
            best_count = sorted_tokens[0][1]
            score = (best_count / len(filenames)) * 70  # 중앙 패턴은 약간 낮은 가중치
        
        return {'score': score, 'patterns': patterns, 'mode': 'middle'}
    
    def _create_group_preview(self, files: list, position: str, result_data: dict) -> dict:
        """패턴 위치에 따라 그룹화 미리보기를 생성합니다."""
        groups = defaultdict(list)
        
        for file_path in files:
            filename = file_path.stem
            group_key = 'UNGROUPED'
            
            if position == 'prefix':
                mode = result_data.get('mode', 'auto')
                delimiter = result_data.get('delimiter', '_')
                
                if mode == 'number':
                    match = re.match(r'^(\d+)', filename)
                    group_key = match.group(1) if match else 'UNGROUPED'
                elif mode == 'delimiter':
                    parts = filename.split(delimiter, 1)
                    group_key = parts[0] if len(parts) > 1 and parts[0] else 'UNGROUPED'
                else:
                    # auto: 숫자 우선
                    match = re.match(r'^(\d+)', filename)
                    if match:
                        group_key = match.group(1)
                    else:
                        for sep in ['_', '-', ' ', '.']:
                            if sep in filename:
                                group_key = filename.split(sep, 1)[0]
                                break
                                
            elif position == 'suffix':
                mode = result_data.get('mode', 'suffix_number')
                delimiter = result_data.get('delimiter', '_')
                
                if mode == 'suffix_number':
                    match = re.search(r'(\d+)$', filename)
                    group_key = match.group(1) if match else 'UNGROUPED'
                else:
                    parts = filename.rsplit(delimiter, 1)
                    group_key = parts[1] if len(parts) > 1 and parts[1] else 'UNGROUPED'
                    
            elif position == 'middle':
                patterns = result_data.get('patterns', [])
                for pattern in patterns:
                    if pattern in filename:
                        group_key = pattern
                        break
            
            groups[group_key].append(file_path)
        
        # 정렬
        for key in groups:
            groups[key] = sorted(groups[key], key=lambda x: x.name.lower())
        
        return dict(sorted(groups.items()))
    
    def _compress_single_pdf(self, pdf_path: Path, quality: int = 50, resize: bool = True) -> Path:
        """
        [v2.1 Bulletproof Pre-Merge Compression] 개별 PDF 이미지 압축.
        PPT→PDF 변환 직후 호출되어 3MB PPT → 15MB PDF 비대화를 원천 차단.
        
        v2.0 대비 수정사항:
        - Alpha 채널 제거 후 JPEG 변환 (RGBA → RGB 안전 변환)
        - 임시 파일 저장 후 성공 시에만 원본 교체 (원본 손상 방지)
        - 불안정한 Pixmap 리사이즈 로직 제거 (JPEG quality만으로 충분)
        
        Args:
            pdf_path: 압축할 PDF 파일 경로
            quality: JPEG 압축 품질 (30-90)
            resize: (예약됨, 현재 미사용 - JPEG quality로 대체)
            
        Returns:
            압축된 PDF 파일 경로
        """
        try:
            import fitz
            doc = fitz.open(str(pdf_path))
            processed_xrefs = set()
            modified = False
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                for img in page.get_images(full=True):
                    xref = img[0]
                    if xref in processed_xrefs:
                        continue
                    processed_xrefs.add(xref)
                    
                    try:
                        orig_len = doc.xref_stream_length(xref)
                        if orig_len < 5000:  # 5KB 미만 아이콘 무시
                            continue
                        
                        pix = fitz.Pixmap(doc, xref)
                        
                        # [v2.1 핵심 수정] Alpha 채널 제거 (JPEG는 투명도 미지원)
                        # WHY: PPT 내 PNG 이미지에 Alpha가 있으면 tobytes("jpeg") 크래시
                        if pix.alpha:
                            pix = fitz.Pixmap(pix, 0)  # 0 = Alpha 채널 드롭
                        
                        # CMYK/CMYKA → RGB 변환 (JPEG 호환성 보장)
                        # WHY: pix.n > 4 체크는 RGBA(n=4)를 놓침 → colorspace 직접 비교
                        if pix.colorspace and pix.colorspace.n >= 4:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        
                        compressed = pix.tobytes("jpeg", quality=quality)
                        
                        # 원본보다 5% 이상 작아질 때만 교체 (무의미한 교체 방지)
                        if len(compressed) < orig_len * 0.95:
                            try:
                                page.replace_image(xref, stream=compressed)
                                modified = True
                            except Exception:
                                try:
                                    doc.update_stream(xref, compressed)
                                    modified = True
                                except Exception:
                                    pass  # 이 이미지는 건너뜀
                    except Exception:
                        continue
            
            # [v2.1 핵심 수정] 안전 저장: 임시파일에 먼저 저장 → 성공 시에만 교체
            # WHY: 같은 경로에 직접 save()하면 실패 시 원본 PDF가 손상됨
            if modified:
                temp_out = pdf_path.parent / f"_tmp_{pdf_path.name}"
                try:
                    doc.save(str(temp_out), garbage=4, deflate=True, clean=True,
                             deflate_images=True, deflate_fonts=True)
                    doc.close()
                    
                    # 원본을 안전하게 교체
                    if temp_out.exists() and temp_out.stat().st_size > 0:
                        temp_out.replace(pdf_path)
                    else:
                        # 임시파일이 비정상이면 삭제하고 원본 유지
                        if temp_out.exists():
                            temp_out.unlink()
                except Exception:
                    doc.close()
                    if temp_out.exists():
                        try: temp_out.unlink()
                        except: pass
            else:
                doc.close()
            
        except Exception as e:
            print(f"개별 PDF 압축 오류 ({pdf_path.name}): {e}")
        
        return pdf_path

    def _kill_office_zombies(self, callback=None):
        """[v2.9.22] 기존 오피스 좀비 프로세스 강제 소거 (UAC Active-Recovery 핵심)"""
        import time, ctypes, subprocess
        try: ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x0002 | 0x8000)
        except: pass
        
        if callback: callback('progress', "[v2.9.22] 기존 Office 프로세스 네이티브 정리 중 (UAC 실드 가동)...")
        targets = ["POWERPNT.EXE", "EXCEL.EXE"]
        killed = False
        
        # 1. taskkill을 통한 물리적 강제 처단 (가장 강력)
        for target in targets:
            try:
                subprocess.run(["taskkill", "/F", "/IM", target, "/T"], 
                               capture_output=True, creationflags=0x08000000)
                killed = True
            except: pass
            
        # 2. psutil을 통한 정밀 소거 (이미 taskkill로 처리되었겠지만 2차 검증)
        try:
            import psutil
            for p in psutil.process_iter(['name']):
                if p.info['name'] and p.info['name'].upper() in targets:
                    try: p.kill(); killed = True
                    except: pass
        except:
            # psutil 없을 경우 WMI 폴백
            try:
                import pythoncom, win32com.client
                pythoncom.CoInitialize()
                wmi = win32com.client.GetObject("winmgmts:")
                for p in wmi.InstancesOf("Win32_Process"):
                    name = p.Properties_('Name').Value
                    if name and name.upper() in targets:
                        try: p.Terminate(); killed = True
                        except: pass
            except: pass
            
        if killed:
            time.sleep(1.0)

    def merge_to_pdf(self, files: list, output_path: Path, callback=None, compress_options=None) -> dict:
        """
        [v2.0 Incremental Merge Architecture]
        여러 문서 파일을 단일 PDF로 병합합니다.
        
        핵심 전략 변경 (v1.9 → v2.0):
        - Before: 전체 변환 → 전체 메모리 적재 → 병합 → 후처리 전체 순회 압축
        - After:  전체 변환+개별압축 → 1개씩 읽기/삽입/해제(점진적) → 저장
        
        Args:
            files: 병합할 파일 경로 리스트 (순서대로 병합)
            output_path: 출력 PDF 파일 경로
            callback: 진행 상황 콜백 (type, message) 형식
            compress_options: {'enabled': bool, 'quality': int, 'resize': bool}
            
        Returns:
            dict: {'success': bool, 'message': str, 'page_count': int}
        """
        # [v3.1.3] fitz is now imported at top level
        
        # 압축 옵션 기본값
        if compress_options is None:
            compress_options = {'enabled': False, 'quality': 70, 'resize': True}
        
        # 임시 폴더 생성 (PDF 변환용)
        self._temp_dir = Path(tempfile.mkdtemp(prefix='merger_'))
        temp_pdfs_map = {}  # {원본경로: 임시PDF경로}
        
        try:
            # 1. 병렬 변환 대상 추출 (이미 PDF인 것은 제외)
            to_convert = []
            for i, f in enumerate(files):
                f_path = Path(f)
                # [v3.0.7] 임시/잠금 파일 (~$) 무시 필터링 강화
                if f_path.name.startswith("~$"):
                    continue
                    
                ext = f_path.suffix.lower()
                if ext in {'.pptx', '.ppt', '.xlsx', '.xls'}:
                    to_convert.append((f, i))
                else:
                    temp_pdfs_map[f] = f  # 이미 PDF인 경우
            
            # 2. 순차 변환 + 즉시 개별 압축 (Single-Thread COM Sequential Conversion)
            # WHY: Dispatch("PowerPoint.Application")는 같은 프로세스에 연결됨
            # 멀티스레드에서 각 스레드가 Quit()하면 다른 스레드의 COM 커넥션이 파괴됨
            # → 실제 테스트 결과: Thread 1 "Object does not exist" 에러 확인
            # → 해결: 단일 스레드 + 단일 COM 인스턴스로 순차 처리
            if to_convert:
                self._kill_office_zombies(callback)
                
                if callback:
                    callback('progress', f"  [📂 Step 1/2] {len(to_convert)}개 Office 문서 PDF로 변환 중...")
                
                converted = self._convert_all_sequential(to_convert, compress_options, callback)
                for orig_path, pdf_path in converted:
                    if pdf_path:
                        temp_pdfs_map[orig_path] = pdf_path
                    else:
                        if callback: callback('warning', f"[주의] 변환 실패 (건너뜀): {orig_path.name}")

            # 3. 점진적 병합 (Incremental Merge - 메모리 안전)
            file_count = len(files)
            if callback:
                callback('progress', f"  [🔗 Step 2/2] 총 {file_count}개 파일 최종 PDF 병합 및 압축 중...")
            
            merged = fitz.open()
            total_pages = 0
            
            for idx, f in enumerate(files):
                pdf_path = temp_pdfs_map.get(f)
                if not pdf_path:
                    continue
                    
                try:
                    # 파일을 바이트로 읽어서 즉시 열기
                    pdf_bytes = pdf_path.read_bytes()
                    doc = fitz.open("pdf", pdf_bytes)
                    merged.insert_pdf(doc)
                    total_pages += len(doc)
                    
                    # 즉시 해제: 해당 파일의 메모리를 즉시 반환
                    doc.close()
                    del pdf_bytes
                    
                    # 진행률 로그 (10개마다 또는 마지막 파일)
                    if callback and ((idx + 1) % 10 == 0 or idx == file_count - 1):
                        callback('progress', f"    • 병합 진행: {idx + 1}/{file_count} ({total_pages}페이지 완료)")
                        
                except Exception as e:
                    if callback:
                        callback('warning', f"[경고] PDF 삽입 실패: {f.name} - {e}")
            
            if total_pages == 0:
                return {'success': False, 'message': '병합할 유효한 페이지가 없습니다.', 'page_count': 0}
            
            # 4. 출력 폴더 생성 및 고속 저장
            # WHY: 개별 PDF가 이미 Pre-Merge 압축 완료 상태이므로
            # 후처리 전체 이미지 순회 압축은 불필요 → 순수 재구성만 수행
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            save_config = {
                'garbage': 4,  # 완전 공간 절약 (Pre-Merge 압축으로 속도 부담 제거됨)
                'deflate': True,
                'clean': True,
                'deflate_images': True,
                'deflate_fonts': True
            }
            
            if callback:
                callback('progress', f"최종 PDF 저장 중 ({total_pages}페이지)...")
            
            merged.save(str(output_path), **save_config)
            merged.close()
            
            return {'success': True, 'message': f'병합 완료: {output_path.name}', 'page_count': total_pages}
            
        except Exception as e:
            return {'success': False, 'message': f'병합 프로세스 중 치명적 오류: {str(e)}', 'page_count': 0}
        
        finally:
            # 임시 파일 정리
            self._cleanup_temp()
    
    def _convert_all_sequential(self, to_convert: list, compress_options: dict = None, callback=None) -> list:
        """
        [v2.4 Out-of-Process COM Isolation]
        Python UI(GUI) 환경이 가진 스레드-보안 컨텍스트(E_ELEVATION_REQUIRED) 문제를 
        원천 회피하기 위해, 변환 로직을 일회용 Python 자식 프로세스로 격리 실행합니다.
        
        WHY:
        GUI가 Unelevated인데 COM(PowerPoint)이 백그라운드에서 Elevated 상태로 좀비화되어 있거나
        서로 권한이 틀어질 경우 pythoncom/Dispatch에서 영구적인 -2147024156 에러가 발생함.
        아예 새로운 독립 프로세스를 생성하여 실행하면 이 권한 상실/충돌 문제를 100% 피해갈 수 있음.
        """
        results = []
        
        do_compress = compress_options and compress_options.get('enabled', False)
        quality = compress_options.get('quality', 50) if compress_options else 50
        resize = compress_options.get('resize', True) if compress_options else True
        
        import subprocess
        import sys
        import time
        
        # 일회용 독립 프로세스가 실행할 일괄 컨버터 스크립트 작성 (v34.1.32 Force Cleanup Armor)
        script_code = """import sys, traceback, time, os, json, ctypes
try:
    ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x0002 | 0x8000)
except: pass

def run_main():
    try:
        import pythoncom, win32com.client, win32api
        from win32com.client import Dispatch, DispatchEx
        from win32com.client.dynamic import Dispatch as DynDispatch
        pythoncom.CoInitialize()

        with open(sys.argv[1], 'r', encoding='utf-8') as f:
            tasks = json.load(f)

        def get_app(prog_id):
            import win32com.client
            from win32com.client import Dispatch, DispatchEx, gencache
            from win32com.client.dynamic import Dispatch as DynDispatch
            import pythoncom
            
            def finalize(a):
                try:
                    if "Excel" in prog_id:
                        a.DisplayAlerts = False
                        a.Visible = False
                    else:
                        a.DisplayAlerts = 1 # msoFalse (ppAlertsNone)
                        a.Visible = 0 # msoFalse
                except: pass
                return a

            errors = []
            clsid = "{91493441-5A91-11CF-8700-00AA0060263B}" if "PowerP" in prog_id else "{00024500-0000-0000-C000-000000000046}"

            try: return finalize(DispatchEx(prog_id))
            except Exception as e: errors.append(f"Ex:{str(e)[:50]}")

            try: return finalize(Dispatch(prog_id))
            except Exception as e: errors.append(f"Std:{str(e)[:50]}")

            try: return finalize(DynDispatch(prog_id))
            except Exception as e: errors.append(f"Dyn:{str(e)[:50]}")

            try:
                unknown = pythoncom.CoCreateInstance(clsid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IUnknown)
                return finalize(Dispatch(unknown))
            except Exception as e: errors.append(f"Clsid:{str(e)[:50]}")

            try: return finalize(gencache.EnsureDispatch(prog_id))
            except Exception as e: errors.append(f"Ensure:{str(e)[:50]}")

            # Level 6: Direct Object Binding (Active Object Capture)
            try:
                # 이미 실행 중인 인스턴스가 있을 가능성 대비 (UAC 벽 너머 낚시)
                return finalize(win32com.client.GetObject(None, prog_id))
            except Exception as e: errors.append(f"Obj:{str(e)[:50]}")

            print(f"DEBUG|FAIL_INIT|{prog_id}|{' / '.join(errors)}", flush=True)
            return None

        ppt_app = None
        xls_app = None

        for task in tasks:
            ipt = os.path.normpath(task['in_path']); opt = os.path.normpath(task['out_path'])
            app_type = task['app_type']
            print(f"PROGRESS|{tasks.index(task)+1}|{len(tasks)}|{os.path.basename(ipt)}", flush=True)
            
            doc_obj = None
            try:
                if app_type == 'ppt':
                    if not ppt_app: ppt_app = get_app("PowerPoint.Application")
                    
                    if ppt_app:
                        try:
                            doc_obj = ppt_app.Presentations.Open(ipt, -1, 0, 0)
                        except Exception as e1:
                            if "-2147024156" in str(e1): doc_obj = win32com.client.GetObject(ipt)
                            else: raise e1
                    else:
                        doc_obj = win32com.client.GetObject(ipt)
                    
                    if not doc_obj: raise Exception("Engine Init Failed (UAC/Permission)")
                    doc_obj.SaveAs(opt, 32)
                    doc_obj.Close()
                    
                elif app_type == 'xls':
                    if not xls_app: xls_app = get_app("Excel.Application")
                    
                    if xls_app:
                        try:
                            doc_obj = xls_app.Workbooks.Open(ipt, ReadOnly=True, UpdateLinks=False)
                        except Exception as e1:
                            if "-2147024156" in str(e1): doc_obj = win32com.client.GetObject(ipt)
                            else: raise e1
                    else:
                        doc_obj = win32com.client.GetObject(ipt)
                    
                    if not doc_obj: raise Exception("Engine Init Failed (UAC/Permission)")
                    
                    # GetObject로 가져온 경우 Worksheets 접근을 위해 application 확인
                    parent_app = doc_obj.Application if hasattr(doc_obj, 'Application') else None
                    if parent_app:
                        parent_app.DisplayAlerts = False
                        parent_app.Visible = False
                    
                    items = doc_obj.Worksheets if hasattr(doc_obj, 'Worksheets') else [doc_obj]
                    for s in items:
                        try: s.Visible = -1
                        except: pass
                    try: doc_obj.Worksheets.Select()
                    except: pass
                    
                    doc_obj.ExportAsFixedFormat(0, opt, Quality=0, IgnorePrintAreas=True)
                    doc_obj.Close(False)

                if os.path.exists(opt): print(f"DEBUG|SUCCESS|{app_type.upper()}|{os.path.basename(ipt)}", flush=True)
                else: print(f"DEBUG|FAIL|{app_type.upper()}|{os.path.basename(ipt)}|ExportResultMissing", flush=True)
                
            except Exception as e:
                msg = f"DEBUG|FAIL|{app_type.upper()}|{os.path.basename(ipt)}|[v34.1.32] {str(e)}"
                if "-2147024156" in str(e):
                    msg += " (UAC 정지: 오피스의 관리자 권한 설정이 강제되어 있습니다. 오피스를 수동으로 실행한 뒤 변환해 보십시오.)"
                print(msg, flush=True)
                if doc_obj:
                    try: doc_obj.Close() if app_type == 'ppt' else doc_obj.Close(False)
                    except: pass

        if ppt_app:
            try: ppt_app.Quit()
            except: pass
        if xls_app:
            try: xls_app.Quit()
            except: pass
        pythoncom.CoUninitialize()
        print("SUCCESS", flush=True)

    except Exception as e:
        print(f"DEBUG|FATAL_INTERNAL|{str(e)}", flush=True)
        traceback.print_exc()

if __name__ == "__main__":
    try:
        run_main()
    except Exception as e:
        print(f"DEBUG|FATAL_CRASH|{str(e)}", flush=True)
        traceback.print_exc()
"""
        script_path = self._temp_dir / "com_wrapper.py"
        script_path.write_text(script_code, encoding="utf-8")

        # No Window 플래그 (터미널 팝업 방지)
        creationflags = 0x08000000 if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0

        import json
        batch_tasks = []
        for file_idx, (file_path, index) in enumerate(to_convert):
            ext = file_path.suffix.lower()
            if ext in {'.pptx', '.ppt', '.xlsx', '.xls'}:
                app_type = 'ppt' if ext in {'.pptx', '.ppt'} else 'xls'
                output_pdf = self._temp_dir / f"{index}_{file_path.stem}_at.pdf"
                batch_tasks.append({
                    'file_path': str(file_path.absolute()),
                    'out_path': str(output_pdf.absolute()),
                    'in_path': str(file_path.absolute()),
                    'app_type': app_type,
                    'index': index
                })
        
        if batch_tasks:
            json_path = self._temp_dir / "batch_tasks.json"
            json_path.write_text(json.dumps(batch_tasks, ensure_ascii=False), encoding="utf-8")
            
            # [v3.1.0] 자식 프로세스의 출력을 UTF-8로 강제하여 CP949 에러 방지
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"

            # [v34.1.32] 초강력 환경 청소: 기존에 꼬여있는 오피스 프로세스 강제 종료
            # WHY: Admin 권한으로 뜬 오피스가 있으면 Normal 권한 스크립트가 접근 불가
            if callback: callback('progress', "  [🧹 Step 0/2] 기존 오피스 프로세스 강제 정리 중 (Clean Slate)...")
            os.system('taskkill /f /im EXCEL.EXE /t >nul 2>&1')
            os.system('taskkill /f /im POWERPNT.EXE /t >nul 2>&1')
            time.sleep(1)

            try:
                process = subprocess.Popen(
                    [sys.executable, str(script_path), str(json_path.absolute())],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                    creationflags=creationflags, encoding='utf-8', env=env
                )
                
                # 실시간 진행률 리딩 루프 (타임아웃 보호 탑재)
                last_activity = time.time()
                while True:
                    try:
                        line = process.stdout.readline()
                    except UnicodeDecodeError:
                        continue
                        
                    if not line:
                        if process.poll() is not None: break
                        if time.time() - last_activity > 60: # 60초간 응답 없으면 강제 종료
                            process.terminate()
                            if callback: callback('error', "[FAIL] 60초간 응답이 없어 변환 엔진을 강제 종료했습니다.")
                            break
                        time.sleep(0.1)
                        continue
                        
                    last_activity = time.time()
                    line = line.strip()
                    if line.startswith("PROGRESS|"):
                        parts = line.split("|")
                        if len(parts) >= 4 and callback:
                            callback('progress', f"   실시간 변환 중: {parts[1]}/{parts[2]} ({parts[3]})")
                    elif line.startswith("DEBUG|"):
                        if callback: callback('warning', f"[알림] {line}")
                    else:
                        # [v34.1.30] Swallow 방지: 형식이 맞지 않는 로그도 투명하게 출력
                        if callback and line: callback('warning', f"[SUB] {line}")

                # 실행 에러 및 최종 출력 캡처
                err_out = process.stderr.read()
                
                # 결과 수집
                for task in batch_tasks:
                    file_path = Path(task['file_path'])
                    output_pdf = Path(task['out_path'])
                    pdf_path = output_pdf if (output_pdf.exists() and output_pdf.stat().st_size > 0) else None
                    
                    if pdf_path and do_compress:
                        self._compress_single_pdf(output_pdf, quality, resize)
                    
                    if not pdf_path:
                        err_msg = err_out.strip() if err_out else "상세 오류 없음 (Office 응답 지연)"
                        if "-2147024156" in err_msg: err_msg = "권한 충돌 (Office 관리자 권한 설정됨)"
                        if callback: callback('warning', f"[주의] 변환 실패 ({file_path.name}): {err_msg[:120]}")
                    
                    results.append((file_path, pdf_path))
                        
            except Exception as e:
                 if callback: callback('warning', f"[FAIL] 시스템 오류: {str(e)[:100]}")
                 for task in batch_tasks:
                     results.append((Path(task['file_path']), None))
                     
        return results


    def _convert_ppt_to_pdf(self, ppt_path: Path, index: int = 0) -> Path:
        """PowerPoint 파일을 PDF로 변환합니다. (COM Automation)"""
        try:
            import win32com.client
            
            if self._ppt_app is None:
                # 메인 프로세스에서도 동일한 고밀도 엔진 초기화 적용 (Fallback 안전장치)
                self._ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                if self._ppt_app is None: raise Exception("PowerPoint Engine Init Failed (NoneType)")
                try:
                    self._ppt_app.Visible = 0 # msoFalse
                except:
                    try: self._ppt_app.WindowState = 2
                    except: pass
                try: self._ppt_app.DisplayAlerts = 1 # ppAlertsNone
                except: pass
            
            # 파일명이 겹칠 수 있으므로 고유 인덱스 부여 (마지막 파일만 변환되는 이슈 해결)
            output_pdf = self._temp_dir / f"{index}_{ppt_path.stem}_converted.pdf"
            
            presentation = self._ppt_app.Presentations.Open(
                str(ppt_path.absolute()), 
                ReadOnly=True, 
                Untitled=False, 
                WithWindow=False
            )
            presentation.SaveAs(str(output_pdf.absolute()), 32)  # 32 = ppSaveAsPDF
            presentation.Close()
            
            return output_pdf
            
        except Exception as e:
            print(f"PPT 변환 오류 ({ppt_path.name}): {e}")
            return None
    
    def _convert_excel_to_pdf(self, excel_path: Path, index: int = 0) -> Path:
        """Excel 파일을 PDF로 변환합니다. (COM Automation)"""
        try:
            import win32com.client
            
            if self._xl_app is None:
                self._xl_app = win32com.client.Dispatch("Excel.Application")
                try:
                    self._xl_app.Visible = False
                    self._xl_app.DisplayAlerts = False
                    self._xl_app.ScreenUpdating = False
                    self._xl_app.Interactive = False
                except: pass
            
            # 고유 인덱스 부여
            output_pdf = self._temp_dir / f"{index}_{excel_path.stem}_converted.pdf"
            
            workbook = self._xl_app.Workbooks.Open(str(excel_path.absolute()), ReadOnly=True, UpdateLinks=False)
            
            # [v2.9.18 Improvement] 모든 시트 선택하여 일괄 PDF 변환 (레이아웃 준수)
            try:
                for sheet in workbook.Worksheets:
                    sheet.Visible = -1
                workbook.Worksheets.Select()
            except: pass
            
            # IgnorePrintAreas=False (페이지 레이아웃 설정값 존중)
            workbook.ExportAsFixedFormat(0, str(output_pdf.absolute()), Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False)
            workbook.Close(False)
            
            return output_pdf
            
        except Exception as e:
            print(f"Excel 변환 오류: {e}")
            return None
    
    def _cleanup_temp(self):
        """임시 폴더 및 파일을 정리합니다."""
        if self._temp_dir and self._temp_dir.exists():
            try:
                shutil.rmtree(self._temp_dir)
            except:
                pass
        self._temp_dir = None
    
    def confirm_merger_results(self, source_path: Path, patterns: list = None, callback=None) -> dict:
        """
        [v2.9.15/16 확정 기능] 병합된 파일을 상위로 이동하고 원본 작업 파일을 정리합니다.
        
        Logic:
        1. 00_Merged_PDFs 내 파일들을 source_path로 이동 (덮어쓰기 허용)
        2. 00_Merged_PDFs 폴더 삭제
        3. 사용자 정의 패턴 또는 기본값(*작업요청서*.pp* , *특기_시방서*.pd*) 파일 자동 삭제
        4. 최종 파일 목록 반환
        """
        if not source_path or not source_path.is_dir():
            return {'success': False, 'message': '유효한 입력 소스 폴더가 아닙니다.'}
            
        merged_dir = source_path / self.OUTPUT_FOLDER_NAME
        moved_count = 0
        deleted_count = 0
        
        # 1. 파일 이동
        if merged_dir.exists() and merged_dir.is_dir():
            if callback: callback('progress', f"📂 병합 결과물 이동 시작: {self.OUTPUT_FOLDER_NAME} → {source_path.name}")
            for f in merged_dir.iterdir():
                if f.is_file():
                    target = source_path / f.name
                    try:
                        # 이미 존재하면 삭제 후 이동 (덮어쓰기 강제)
                        if target.exists():
                            target.unlink()
                        shutil.move(str(f), str(target))
                        moved_count += 1
                    except Exception as e:
                        if callback: callback('warning', f"⚠️ 이동 실패 ({f.name}): {e}")
            
            # 2. 폴더 삭제
            try:
                shutil.rmtree(merged_dir)
                if callback: callback('progress', f"🗑️ 임시 폴더 삭제 완료: {self.OUTPUT_FOLDER_NAME}")
            except Exception as e:
                if callback: callback('warning', f"⚠️ 폴더 삭제 실패: {e}")
        else:
            if callback: callback('warning', "⚠️ 확정할 병합 폴더(00_Merged_PDFs)가 존재하지 않습니다.")

        # 3. 원본 작업 파일 삭제 (사용자 정의 또는 기본값)
        if patterns is None:
            patterns = ["*작업요청서*.pp*", "*특기_시방서*.pd*"]
            
        patterns_str = ", ".join(patterns)
        if callback: callback('progress', f"🧹 원본 작업 파일 정밀 소거 중 ({patterns_str})...")
        
        import fnmatch
        for f in source_path.iterdir():
            if f.is_file():
                for p in patterns:
                    if fnmatch.fnmatch(f.name.lower(), p.lower().strip()):
                        try:
                            f.unlink()
                            deleted_count += 1
                            if callback: callback('progress', f"   [삭제완료] {f.name}")
                        except Exception as e:
                            if callback: callback('warning', f"   [삭제실패] {f.name}: {e}")
                            
        # 4. 최종 파일 목록 스캔
        final_files = sorted([f.name for f in source_path.iterdir() if f.is_file()])
        
        return {
            'success': True,
            'moved': moved_count,
            'deleted': deleted_count,
            'final_list': final_files,
            'message': f"확정 완료: {moved_count}개 이동, {deleted_count}개 작업파일 소거됨."
        }
    
    def cleanup_com(self):
        """COM 객체를 정리합니다."""
        if self._xl_app:
            try:
                self._xl_app.Quit()
            except:
                pass
            self._xl_app = None
        
        if self._ppt_app:
            try:
                self._ppt_app.Quit()
            except:
                pass
            self._ppt_app = None
        
        import gc
        gc.collect()


# ═══════════════════════════════════════════════════════════════════════════════
# LAYER 2: PRESENTATION (View - GUI 정의)
# - 위젯 배치만 담당, 로직 절대 금지
# ═══════════════════════════════════════════════════════════════════════════════

class PatternMergerView:
    """
    Tkinter 기반 GUI 레이아웃을 정의합니다.
    비즈니스 로직 없이 순수하게 위젯 배치만 담당합니다.
    """
    
    def __init__(self, root: tk.Tk, controller):
        self.root = root
        self.controller = controller
        
        self.root.title("[패턴 기반 문서 통합 병합기 v34.1.32]")
        self.root.geometry("1000x900")
        self.root.minsize(900, 700)
        
        # 변수 초기화
        self.source_var = tk.StringVar()
        self.mode_var = tk.StringVar(value='pattern')  # batch | pattern
        self.pattern_position_var = tk.StringVar(value='prefix')  # prefix | suffix | middle
        self.recursive_var = tk.BooleanVar(value=True)
        self.output_name_var = tk.StringVar()
        self.use_first_name_var = tk.BooleanVar(value=True)
        
        # 압축 옵션 변수
        self.quality_var = tk.IntVar(value=50)  # 이미지 품질 (30-90), 기본값 Web(50)
        self.resize_var = tk.BooleanVar(value=True)  # 큰 이미지 리사이즈
        self.compress_enabled_var = tk.BooleanVar(value=True)  # 압축 활성화
        
        self._build_ui()
    
    def _build_ui(self):
        """[v1.6] 메인 레이아웃을 스크롤 가능하게 구성하여 모든 위젯 가시성 확보."""
        # ─── 하단 상태 표시 영역 (버튼은 미리보기 타이틀 라인으로 이동) ───
        bottom_frame = ttk.Frame(self.root, padding="10")
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = ttk.Label(bottom_frame, text="[INFO] 폴더 또는 파일을 선택하세요", font=('맑은 고딕', 10))
        self.status_label.pack(side=tk.LEFT)

        # ─── 메인 스크롤 영역 (핵심 개선) ───
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        
        # 내부 프레임
        self.scroll_content_frame = ttk.Frame(canvas, padding="10")
        self.scroll_content_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scroll_content_frame, anchor="nw", width=960)
        canvas.configure(yscrollcommand=scrollbar.set)

        # 마우스 휠 지원
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        content_frame = self.scroll_content_frame
        
        # ─── 1. 입력 소스 영역 ───
        source_frame = ttk.LabelFrame(content_frame, text="[1] 입력 소스", padding="10")
        source_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Entry(source_frame, textvariable=self.source_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(source_frame, text="폴더 선택", command=self.controller.handle_select_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(source_frame, text="파일 선택", command=self.controller.handle_select_files).pack(side=tk.LEFT, padx=2)
        
        ttk.Checkbutton(source_frame, text="하위 폴더 포함", variable=self.recursive_var).pack(side=tk.LEFT, padx=(10, 0))
        
        # ─── 2. 병합 모드 선택 ───
        mode_frame = ttk.LabelFrame(content_frame, text="[2] 병합 모드", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        left_mode = ttk.Frame(mode_frame)
        left_mode.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Radiobutton(left_mode, text="일괄 병합 (모든 파일 → 1개 PDF)", variable=self.mode_var, value='batch').pack(anchor=tk.W)
        ttk.Radiobutton(left_mode, text="패턴 분할 병합 (감지된 패턴별 → 그룹별 PDF)", variable=self.mode_var, value='pattern').pack(anchor=tk.W)
        
        # ─── 3. 심층 패턴 설정 ───
        pattern_frame = ttk.LabelFrame(content_frame, text="[3] 패턴 위치 선택", padding="10")
        pattern_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 상단: 패턴 분석 버튼
        analyze_row = ttk.Frame(pattern_frame)
        analyze_row.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(analyze_row, text="[패턴 자동 분석]", command=self.controller.handle_analyze).pack(side=tk.LEFT)
        ttk.Label(analyze_row, text="← 파일 선택 후 클릭하면 최적 패턴을 자동 감지합니다").pack(side=tk.LEFT, padx=10)
        
        # 패턴 위치 선택 라디오버튼
        position_row = ttk.Frame(pattern_frame)
        position_row.pack(fill=tk.X)
        
        ttk.Label(position_row, text="그룹화 기준:").pack(side=tk.LEFT, padx=(0, 10))
        
        # 점수 표시 라벨 (분석 후 업데이트됨)
        self.prefix_radio = ttk.Radiobutton(position_row, text="[접두사] (Prefix)", 
                                             variable=self.pattern_position_var, value='prefix')
        self.prefix_radio.pack(side=tk.LEFT, padx=5)
        self.prefix_score_label = ttk.Label(position_row, text="")
        self.prefix_score_label.pack(side=tk.LEFT)
        
        self.suffix_radio = ttk.Radiobutton(position_row, text="[접미사] (Suffix)", 
                                             variable=self.pattern_position_var, value='suffix')
        self.suffix_radio.pack(side=tk.LEFT, padx=5)
        self.suffix_score_label = ttk.Label(position_row, text="")
        self.suffix_score_label.pack(side=tk.LEFT)
        
        self.middle_radio = ttk.Radiobutton(position_row, text="[중앙] (Middle)", 
                                             variable=self.pattern_position_var, value='middle')
        self.middle_radio.pack(side=tk.LEFT, padx=5)
        self.middle_score_label = ttk.Label(position_row, text="")
        self.middle_score_label.pack(side=tk.LEFT)
        
        # ─── 4. 압축 설정 ───
        compress_frame = ttk.LabelFrame(content_frame, text="[4] 이미지 압축 설정", padding="10")
        compress_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 압축 활성화 체크박스
        compress_top = ttk.Frame(compress_frame)
        compress_top.pack(fill=tk.X)
        ttk.Checkbutton(compress_top, text="PDF 병합 시 이미지 압축 적용", 
                        variable=self.compress_enabled_var).pack(side=tk.LEFT)
        
        # 품질 슬라이더
        quality_row = ttk.Frame(compress_frame)
        quality_row.pack(fill=tk.X, pady=(10, 5))
        ttk.Label(quality_row, text="이미지 품질:").pack(side=tk.LEFT)
        tk.Scale(quality_row, from_=30, to=90, orient=tk.HORIZONTAL, 
                 variable=self.quality_var, showvalue=True, length=150).pack(side=tk.LEFT, padx=10)
        
        # 프리셋 버튼
        ttk.Button(quality_row, text="Web (50)", command=lambda: self.quality_var.set(50)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quality_row, text="Standard (70)", command=lambda: self.quality_var.set(70)).pack(side=tk.LEFT, padx=2)
        ttk.Button(quality_row, text="Print (85)", command=lambda: self.quality_var.set(85)).pack(side=tk.LEFT, padx=2)
        
        # 리사이즈 옵션
        resize_row = ttk.Frame(compress_frame)
        resize_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Checkbutton(resize_row, text="큰 이미지 자동 축소 (4K → QHD 2560px)", 
                        variable=self.resize_var).pack(side=tk.LEFT)
        ttk.Label(resize_row, text="  [SIZE] 파일 크기 대폭 감소", foreground="gray").pack(side=tk.LEFT)
        
        # ─── 5. 출력 이름 설정 ───
        output_frame = ttk.LabelFrame(content_frame, text="[5] 출력 파일 이름", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Checkbutton(output_frame, text="첫 번째 파일 이름 사용 (기본값)", variable=self.use_first_name_var).pack(anchor=tk.W)
        
        custom_row = ttk.Frame(output_frame)
        custom_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(custom_row, text="사용자 지정 이름:").pack(side=tk.LEFT)
        ttk.Entry(custom_row, textvariable=self.output_name_var, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Label(custom_row, text="(일괄 병합 시 적용)").pack(side=tk.LEFT)
        
        # ─── 5. 미리보기 영역 + 액션 버튼 (타이틀 라인에 배치) ───
        preview_frame = ttk.LabelFrame(content_frame, text="[6] 패턴 분석 미리보기 (대규모 그룹 정보 확인용)", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 버튼을 미리보기 타이틀 라인에 배치 (하단 고정에서 이동)
        action_row = ttk.Frame(preview_frame)
        action_row.pack(fill=tk.X, pady=(0, 5))
        
        self.analyze_btn = ttk.Button(action_row, text="[패턴 분석]", command=self.controller.handle_analyze)
        self.analyze_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.run_btn = ttk.Button(action_row, text="[병합 실행]", command=self.controller.handle_run)
        self.run_btn.pack(side=tk.LEFT, padx=5, ipadx=20, ipady=3)
        
        self.confirm_btn = ttk.Button(action_row, text="[확정] (파일 정리)", command=self.controller.handle_confirm)
        self.confirm_btn.pack(side=tk.LEFT, padx=5, ipadx=10, ipady=3)
        
        preview_scroll = ttk.Scrollbar(preview_frame)
        preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preview_text = tk.Text(preview_frame, height=12, wrap=tk.WORD, yscrollcommand=preview_scroll.set, 
                                     font=('Consolas', 10), state='disabled', background="#f1f5f9")
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        preview_scroll.config(command=self.preview_text.yview)
        
        # ─── 6. 로그 영역 (가독성 확장) ───
        log_frame = ttk.LabelFrame(content_frame, text="[7] 실시간 실행 로그 (병목 지점 확인)", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, yscrollcommand=log_scroll.set, 
                                 font=('Consolas', 9), state='disabled', background="#fdfdfd")
        self.log_text.pack(fill=tk.BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)
    
    def set_preview(self, text: str):
        """미리보기 영역에 텍스트를 표시합니다."""
        self.preview_text.config(state='normal')
        self.preview_text.delete('1.0', tk.END)
        self.preview_text.insert('1.0', text)
        self.preview_text.config(state='disabled')
    
    def append_log(self, message: str):
        """로그 영역에 메시지를 추가합니다."""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
    
    def clear_log(self):
        """로그 영역을 초기화합니다."""
        self.log_text.config(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state='disabled')
    
    def set_status(self, text: str):
        """상태 라벨을 업데이트합니다."""
        self.status_label.config(text=text)
    
    def set_running(self, is_running: bool):
        """실행 상태에 따라 버튼 상태를 변경합니다."""
        state = 'disabled' if is_running else 'normal'
        self.run_btn.config(state=state)
        self.confirm_btn.config(state=state)
        self.analyze_btn.config(state=state)
    
    def update_pattern_scores(self, scores: dict, best_position: str):
        """
        패턴 분석 점수를 업데이트하고 최고 점수 위치를 자동 선택합니다.
        
        Args:
            scores: {'prefix': 점수, 'suffix': 점수, 'middle': 점수}
            best_position: 'prefix', 'suffix', 'middle' 중 하나
        """
        # 점수 라벨 업데이트
        prefix_score = scores.get('prefix', 0)
        suffix_score = scores.get('suffix', 0)
        middle_score = scores.get('middle', 0)
        
        self.prefix_score_label.config(text=f"[{prefix_score:.0f}점]" if prefix_score > 0 else "")
        self.suffix_score_label.config(text=f"[{suffix_score:.0f}점]" if suffix_score > 0 else "")
        self.middle_score_label.config(text=f"[{middle_score:.0f}점]" if middle_score > 0 else "")
        
        # 최고 점수 위치 자동 선택
        self.pattern_position_var.set(best_position)

    def show_confirm_dialog(self, source_path: str, default_patterns: str) -> dict:
        """[v2.9.17] 확정 및 파일 정리 설정을 위한 커스텀 다이얼로그 (경로 선택 가능)"""
        dialog = tk.Toplevel(self.root)
        dialog.title("확정 및 파일 정리 설정")
        dialog.geometry("550x280")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 중앙 배치
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 275
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 140
        dialog.geometry(f"+{x}+{y}")
        
        content = ttk.Frame(dialog, padding="20")
        content.pack(fill=tk.BOTH, expand=True)
        
        # 📁 대상 폴더 경로 (수정 가능하게 변경)
        ttk.Label(content, text="[경로] 대상 폴더 경로:", font=('맑은 고딕', 9, 'bold')).pack(anchor=tk.W)
        
        path_frame = ttk.Frame(content)
        path_frame.pack(fill=tk.X, pady=(5, 15))
        
        path_entry = ttk.Entry(path_frame, font=('Consolas', 9))
        path_entry.insert(0, source_path)
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        def browse_folder():
            folder = filedialog.askdirectory(title="대상 폴더 변경", initialdir=path_entry.get())
            if folder:
                path_entry.config(state='normal')
                path_entry.delete(0, tk.END)
                path_entry.insert(0, folder)
        
        ttk.Button(path_frame, text="찾기...", width=8, command=browse_folder).pack(side=tk.RIGHT, padx=(5, 0))
        
        # 🧹 삭제할 파일 패턴
        ttk.Label(content, text="[CLEAN] 삭제할 파일 패턴 (쉼표 구분):", font=('맑은 고딕', 9, 'bold')).pack(anchor=tk.W)
        pattern_entry = ttk.Entry(content, font=('Consolas', 10))
        pattern_entry.insert(0, default_patterns)
        pattern_entry.pack(fill=tk.X, pady=(5, 10))
        pattern_entry.focus_set()
        
        ttk.Label(content, text="※ 확정 시 해당 폴더의 '00_Merged_PDFs' 결과물은 상위로 이동되며 지정된 원본은 삭제됩니다.", 
                  foreground="#64748b", font=('맑은 고딕', 8)).pack(anchor=tk.W)
        
        btn_frame = ttk.Frame(content)
        btn_frame.pack(fill=tk.X, pady=(15, 0))
        
        result = {'path': None, 'patterns': None}
        
        def on_ok():
            p = path_entry.get().strip()
            if not p or not Path(p).is_dir():
                messagebox.showerror("오류", "유효한 폴더 경로를 입력하거나 선택해 주세요.")
                return
            result['path'] = p
            result['patterns'] = pattern_entry.get()
            dialog.destroy()
            
        def on_cancel():
            dialog.destroy()
            
        ttk.Button(btn_frame, text="취소", command=on_cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="[확정 실행]", command=on_ok).pack(side=tk.RIGHT, padx=5)
        
        dialog.bind('<Return>', lambda e: on_ok())
        dialog.bind('<Escape>', lambda e: on_cancel())
        
        self.root.wait_window(dialog)
        return result if result['path'] else None


# ═══════════════════════════════════════════════════════════════════════════════
# LAYER 3: APPLICATION (Controller - 이벤트 중재)
# - View와 Engine 연결, 스레드 관리
# ═══════════════════════════════════════════════════════════════════════════════

class PatternMergerController:
    """
    사용자 이벤트를 처리하고 Engine과 View를 연결합니다.
    스레드 안전성을 보장하여 UI 프리징을 방지합니다.
    """
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.engine = PatternMergerEngine()
        self.view = PatternMergerView(root, self)
        
        self._current_files = []
        self._current_groups = {}
        self._is_running = False
        self._detected_mode = 'number'
        self._detected_delimiter = '_'
        self._pattern_position = 'prefix'  # 'prefix', 'suffix', 'middle'
        
        # 윈도우 종료 시 COM 정리
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _on_closing(self):
        """윈도우 종료 시 리소스를 정리합니다."""
        self.engine.cleanup_com()
        self.root.destroy()
    
    def handle_select_folder(self):
        """폴더 선택 다이얼로그를 엽니다."""
        folder = filedialog.askdirectory(title="병합할 파일이 있는 폴더 선택")
        if folder:
            self.view.source_var.set(folder)
            self._refresh_files()
    
    def handle_select_files(self):
        """파일 다중 선택 다이얼로그를 엽니다."""
        filetypes = [
            ("지원 문서", "*.pdf;*.pptx;*.ppt;*.xlsx;*.xls"),
            ("PDF 파일", "*.pdf"),
            ("PowerPoint", "*.pptx;*.ppt"),
            ("Excel", "*.xlsx;*.xls"),
        ]
        files = filedialog.askopenfilenames(title="병합할 파일 선택", filetypes=filetypes)
        if files:
            self.view.source_var.set(';'.join(files))
            self._refresh_files()
    
    def _refresh_files(self):
        """선택된 소스에서 파일을 새로 스캔합니다."""
        source = self.view.source_var.get()
        if not source:
            return
        
        recursive = self.view.recursive_var.get()
        self._current_files = self.engine.scan_files(source, recursive)
        
        self.view.set_status(f"발견된 파일: {len(self._current_files)}개")
        self.handle_analyze()
    
    def handle_analyze(self):
        """현재 파일 목록에서 패턴을 심층 분석하고 미리보기를 표시합니다."""
        if not self._current_files:
            self.view.set_preview("선택된 파일이 없습니다. 폴더 또는 파일을 선택해주세요.")
            return
        
        # 심층 패턴 분석 수행
        auto_result = self.engine.auto_detect_patterns(self._current_files)
        
        # 결과 저장
        self._current_groups = auto_result.get('group_preview', {})
        self._detected_mode = auto_result.get('detected_mode', 'number')
        self._detected_delimiter = auto_result.get('detected_delimiter', '_')
        self._pattern_position = auto_result.get('pattern_position', 'prefix')
        
        # 분석에서 점수 추출 (Engine에서 반환)
        scores = auto_result.get('scores', {})
        if not scores:
            # scores가 없으면 분석 텍스트에서 추정 (fallback)
            scores = {'prefix': 0, 'suffix': 0, 'middle': 0}
            analysis_text = auto_result.get('analysis', '')
            import re
            for pos in ['prefix', 'suffix', 'middle']:
                # 점수 파싱 시도
                pos_kr = {'prefix': '접두사', 'suffix': '접미사', 'middle': '중앙'}[pos]
                match = re.search(rf'{pos_kr}.*?점수\s*(\d+)', analysis_text)
                if match:
                    scores[pos] = int(match.group(1))
        
        # View에 점수 업데이트 및 자동 선택
        self.view.update_pattern_scores(scores, self._pattern_position)
        
        # ─── 파일 통계 정보 생성 ───
        stats_lines = []
        stats_lines.append("[INFO] 대상 파일 정보")
        stats_lines.append("=" * 50)
        
        # 폴더 수 계산
        unique_folders = set(f.parent for f in self._current_files)
        stats_lines.append(f"[-] 대상 폴더: {len(unique_folders)}개")
        
        # 파일 수
        stats_lines.append(f"[-] 대상 파일: {len(self._current_files)}개")
        
        # 유형별 분류
        type_count = {}
        for f in self._current_files:
            ext = f.suffix.lower()
            type_count[ext] = type_count.get(ext, 0) + 1
        
        type_str = ", ".join([f"{ext.upper()[1:]}: {count}개" for ext, count in sorted(type_count.items())])
        stats_lines.append(f"📋 유형별: {type_str}")
        
        # 폴더 목록 (최대 5개)
        if len(unique_folders) > 1:
            stats_lines.append("\n[-] 폴더 구조:")
            for folder in list(unique_folders)[:5]:
                stats_lines.append(f"   • {folder.name}/")
            if len(unique_folders) > 5:
                stats_lines.append(f"   • ... 외 {len(unique_folders) - 5}개 폴더")
        
        stats_lines.append("")
        
        # 미리보기 텍스트 생성
        lines = []
        lines.extend(stats_lines)
        lines.append(auto_result.get('analysis', ''))
        lines.append("")
        lines.append(f"📊 총 {len(self._current_files)}개 파일 → {len(self._current_groups)}개 그룹\n")
        
        # 그룹별 파일 목록 표시
        for prefix, files in self._current_groups.items():
            lines.append(f"\n[GROUP] [{prefix}]: {len(files)}개 파일")
            for f in files:
                lines.append(f"   • {f.name}")
            lines.append(f"   → 출력: {files[0].stem}.pdf")
        
        self.view.set_preview('\n'.join(lines))
    
    def handle_run(self):
        """병합 작업을 실행합니다."""
        if self._is_running:
            return
        
        if not self._current_files:
            messagebox.showwarning("경고", "병합할 파일이 없습니다.")
            return
        
        self._is_running = True
        self.view.set_running(True)
        self.view.clear_log()
        
        # 백그라운드 스레드에서 실행
        threading.Thread(target=self._run_merge_with_com, daemon=True).start()
    
    def _run_merge_with_com(self):
        """COM 초기화와 함께 병합 작업을 실행합니다."""
        try:
            pythoncom.CoInitialize()
            self._run_merge()
        finally:
            pythoncom.CoUninitialize()
    
    def _run_merge(self):
        """실제 병합 작업을 수행합니다. (백그라운드 스레드)"""
        try:
            mode = self.view.mode_var.get()
            source = self.view.source_var.get()
            
            # 출력 폴더 결정
            if ';' in source:
                first_file = Path(source.split(';')[0].strip())
                output_dir = first_file.parent / PatternMergerEngine.OUTPUT_FOLDER_NAME
            else:
                output_dir = Path(source) / PatternMergerEngine.OUTPUT_FOLDER_NAME
            
            output_dir.mkdir(parents=True, exist_ok=True)
            
            self._log("progress", f"[INFO] 출력 폴더: {output_dir}")
            
            if mode == 'batch':
                # 일괄 병합
                self._run_batch_merge(output_dir)
            else:
                # 패턴 분할 병합
                self._run_pattern_merge(output_dir)
            
            self._log("status", "[OK] 모든 작업이 완료되었습니다.")
            
        except Exception as e:
            self._log("error", f"[FAIL] 오류 발생: {str(e)}")
        
        finally:
            # COM 객체 정리
            self.engine.cleanup_com()
            
            # UI 업데이트 (메인 스레드)
            self.root.after(0, self._finalize)
    
    def _run_batch_merge(self, output_dir: Path):
        """일괄 병합을 수행합니다."""
        # 출력 파일명 결정
        if self.view.use_first_name_var.get() or not self.view.output_name_var.get():
            output_name = self._current_files[0].stem + "_병합.pdf"
        else:
            output_name = self.view.output_name_var.get()
            if not output_name.endswith('.pdf'):
                output_name += '.pdf'
        
        output_path = output_dir / output_name
        
        self._log("progress", f"[INFO] 일괄 병합 시작: {len(self._current_files)}개 파일 → {output_name}")
        
        # 압축 옵션 수집
        compress_options = {
            'enabled': self.view.compress_enabled_var.get(),
            'quality': self.view.quality_var.get(),
            'resize': self.view.resize_var.get()
        }
        
        if compress_options['enabled']:
            self._log("progress", f"🗜️ 압축 설정: 품질 {compress_options['quality']}%, 리사이즈 {'ON' if compress_options['resize'] else 'OFF'}")
        
        result = self.engine.merge_to_pdf(
            self._current_files, 
            output_path, 
            callback=self._log,
            compress_options=compress_options
        )
        
        if result['success']:
            self._log("success", f"[OK] {result['message']} ({result['page_count']}페이지)")
        else:
            self._log("error", f"[FAIL] {result['message']}")
    
    def _kill_office_processes(self, callback=None):
        """[v2.9.22] 기존 오피스 좀비 프로세스 강제 청소 (Engine 로직 호출 최적화)"""
        self.engine._kill_office_zombies(callback)

    def _run_pattern_merge(self, output_dir: Path):
        """패턴 분할 병합을 수행합니다."""
        total_groups = len(self._current_groups)
        
        # 압축 옵션 수집
        compress_options = {
            'enabled': self.view.compress_enabled_var.get(),
            'quality': self.view.quality_var.get(),
            'resize': self.view.resize_var.get()
        }
        
        if compress_options['enabled']:
            self._log("progress", f"[OPT] 압축 설정: 품질 {compress_options['quality']}%, 리사이즈 {'ON' if compress_options['resize'] else 'OFF'}")
        
        for i, (prefix, files) in enumerate(self._current_groups.items(), 1):
            self._log("progress", f"\n[{i}/{total_groups}] 그룹 '{prefix}' 처리 중 ({len(files)}개 파일)")
            
            # 출력 파일명: 첫 번째 파일명 사용
            output_name = files[0].stem + ".pdf"
            output_path = output_dir / output_name
            
            result = self.engine.merge_to_pdf(
                files, 
                output_path, 
                callback=self._log,
                compress_options=compress_options
            )
            
            if result['success']:
                self._log("success", f"   [OK] {output_name} 생성 완료 ({result['page_count']}페이지)")
            else:
                self._log("error", f"   [FAIL] {result['message']}")
    
    def _log(self, log_type: str, message: str):
        """스레드 안전하게 로그를 출력합니다."""
        self.root.after(0, lambda: self.view.append_log(message))
        
        if log_type == 'progress':
            self.root.after(0, lambda: self.view.set_status(message[:50] + "..."))
    
    def _finalize(self):
        """작업 완료 후 UI 상태를 복원합니다."""
        self._is_running = False
        self.view.set_running(False)
        self.view.set_status("작업 완료")

    def handle_confirm(self):
        """[v2.9.15/17] 확정 프로세스 실행: 대상 경로 변경 및 사용자 정의 상위 정리"""
        source = self.view.source_var.get()
        # 단일 폴더 모드 체크
        if not source or ';' in source or not Path(source).is_dir():
            messagebox.showwarning("경고", "확정 기능은 단일 폴더 선택 모드에서만 지원됩니다.")
            return
            
        # 1. 확정 설정 다이얼로그 표시 (경로 및 패턴 포함)
        default_pats = "*작업요청서*.pp*, *특기_시방서*.pd*"
        user_config = self.view.show_confirm_dialog(source, default_pats)
        
        if not user_config: # 취소 버튼 또는 잘못된 경로
            return
            
        final_source = Path(user_config['path'])
        cleanup_patterns = [p.strip() for p in user_config['patterns'].split(',') if p.strip()]
        
        self._is_running = True
        self.view.set_running(True)
        self.view.append_log("\n" + "═"*50)
        self.view.append_log(f"[START] 최종 확정 및 파일 정리 프로세스 가동 (v34.1.32)")
        self.view.append_log(f"   [DIR] 대상: {final_source}")
        self.view.append_log("═"*50)
        
        def run():
            try:
                # 선택된 최신 경로와 패턴 전달
                result = self.engine.confirm_merger_results(final_source, cleanup_patterns, self._log)
                self.root.after(0, lambda: self._finalize_confirm(result))
            except Exception as e:
                self.root.after(0, lambda: self.view.append_log(f"[FAIL] 확정 오류: {e}"))
                self.root.after(0, lambda: self.view.set_running(False))
                
        threading.Thread(target=run, daemon=True).start()

    def _finalize_confirm(self, result: dict):
        """확정 프로세스 완료 후 결과 표시"""
        self._is_running = False
        self.view.set_running(False)
        
        if result['success']:
            self.view.append_log("\n[OK] 확정 완료 현황")
            self.view.append_log(f"   - 이동된 병합 파일: {result['moved']}개")
            self.view.append_log(f"   - 삭제된 작업 파일: {result['deleted']}개")
            self.view.append_log("\n[FILE] [현재 폴더 파일 목록]")
            for f in result['final_list']:
                self.view.append_log(f"   • {f}")
            
            self.view.set_status("확정 완료")
            messagebox.showinfo("확정 완료", result['message'])
        else:
            messagebox.showerror("오류", result['message'])


# ═══════════════════════════════════════════════════════════════════════════════
# 메인 진입점
# ═══════════════════════════════════════════════════════════════════════════════

def is_admin():
    """현재 프로세스가 관리자 권한으로 실행 중인지 확인합니다."""
    import ctypes
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

def run_as_admin():
    """프로그램을 관리자 권한으로 자동으로 재발행(Elevation)합니다."""
    import ctypes, sys, os
    if is_admin():
        return True
    
    # [v2.7.1] 경로에 공백이 포함된 경우를 대비해 쿼테이션 처리
    params = f'"{sys.argv[0]}" ' + " ".join([f'"{arg}"' for arg in sys.argv[1:]])
    try:
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, params, None, 1
        )
    except Exception as e:
        print(f"관리자 권한 승격 실패: {e}")
    sys.exit(0)

def main():
    """애플리케이션 진입점"""
    import ctypes
    try:
        # 시스템 에러 팝업(잘못된 이미지, DLL 로드 에러 등) 억제 글로벌 적용
        ctypes.windll.kernel32.SetErrorMode(0x0001 | 0x0002 | 0x8000)
    except: pass
    
    root = tk.Tk()
    
    # [v2.8.1] 창 활성화 및 포커스 강제 (대시보드 실행 시 화면 뒤에 숨는 현상 방지)
    root.lift()
    root.focus_force()
    root.attributes('-topmost', True)
    root.after(500, lambda: root.attributes('-topmost', False))
    
    # 스타일 설정
    style = ttk.Style()
    style.configure('TButton', font=('맑은 고딕', 10))
    style.configure('TLabel', font=('맑은 고딕', 10))
    style.configure('TLabelframe.Label', font=('맑은 고딕', 10, 'bold'))
    
    app = PatternMergerController(root)
    root.mainloop()


if __name__ == "__main__":
    # [v2.9.24 Root Cause Fix] 관리자 권한 자동 승격 제거
    # MS Office COM은 Elevated(관리자 권한) 프로세스에서 
    # 일반(Medium Integrity) Office App을 제어하려 할 때 
    # ERROR_ELEVATION_REQUIRED (-2147024156)를 발생시킵니다.
    # UAC 충돌을 완벽히 해결하기 위해 스크립트를 일반 권한으로 실행합니다.
    # run_as_admin()  <-- 절대 활성화하지 마세요.
    main()
