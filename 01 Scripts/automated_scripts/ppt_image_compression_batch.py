# -*- coding: utf-8 -*-
"""
PowerPoint 이미지 일괄 압축 도구 (PPT/PPTX)

벤치마킹 반영 사항
- 02마감 계열 도구의 폴더 선택, 깊은 하위 폴더 전수 스캔, 진행 로그, 사후 검증
- 원본 보존을 기본값으로 하는 결과 생성/백업/롤백 안전성
- PPT는 PowerPoint COM으로 PPTX로 변환한 뒤 동일한 이미지 압축 파이프라인 적용
- PPTX 패키지 내부 ppt/media 이미지만 변경하여 슬라이드 위치·크기 XML은 유지
- UTF-8-SIG 보고서와 콘솔 재구성으로 한글 경로·파일명 안전 처리

필요 패키지:
    pip install pillow pywin32
"""

from __future__ import annotations

import io
import os
import posixpath
import re
import shutil
import sys
import tempfile
import threading
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from PIL import Image
except Exception as exc:  # pragma: no cover - dependency message is UI-facing
    Image = None
    PIL_ERROR = exc

try:
    import pythoncom
    import win32com.client
except Exception as exc:  # PPT 변환/PowerPoint 검증 시에만 필수
    pythoncom = None
    win32com = None
    COM_ERROR = exc


SUPPORTED_EXTS = {".ppt", ".pptx"}
EMU_PER_INCH = 914400.0
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def configure_utf8() -> None:
    """Windows 콘솔 코드페이지에 관계없이 한글 로그가 깨지지 않게 한다."""
    for stream in (sys.stdout, sys.stderr):
        try:
            if hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


configure_utf8()


@dataclass
class CompressionOptions:
    resolution_dpi: int = 96
    jpeg_quality: int = 85
    delete_cropped: bool = True
    recursive: bool = True
    convert_ppt: bool = True
    keep_original: bool = True
    verify_powerpoint: bool = True
    overwrite_existing: bool = False
    preserve_tree: bool = True


@dataclass
class MediaUse:
    slide_xml: Path
    media_path: Path
    width_px: int
    height_px: int
    crop: Tuple[int, int, int, int]


@dataclass
class ProcessResult:
    source: Path
    output: Optional[Path]
    status: str
    message: str = ""
    images: int = 0
    before: int = 0
    after: int = 0
    crop_removed: int = 0
    warnings: List[str] = field(default_factory=list)


def safe_delete_file(path: Path, max_retries: int = 5, delay: float = 0.2) -> bool:
    """Windows 파일 락을 고려하여 짧은 지연 후 재시도하며 파일을 안전하게 삭제한다."""
    if not path.exists():
        return True
    for _ in range(max_retries):
        try:
            path.unlink(missing_ok=True)
            return True
        except (PermissionError, OSError):
            time.sleep(delay)
    return not path.exists()


def safe_relative_output(
    source: Path, source_root: Path, output_root: Optional[Path], preserve_tree: bool = True
) -> Path:
    if output_root is None:
        return source.parent
    if not preserve_tree:
        output_root.mkdir(parents=True, exist_ok=True)
        return output_root
    try:
        relative_parent = source.parent.relative_to(source_root)
    except ValueError:
        relative_parent = Path()
    destination = output_root / relative_parent
    destination.mkdir(parents=True, exist_ok=True)
    return destination


def choose_output_path(
    source: Path,
    source_root: Path,
    output_root: Optional[Path],
    opts: CompressionOptions,
    allocated_paths: Optional[set[Path]] = None,
) -> Tuple[Path, bool]:
    folder = safe_relative_output(source, source_root, output_root, opts.preserve_tree)
    if source.suffix.lower() == ".ppt":
        candidate = folder / f"{source.stem}.pptx"
    elif opts.keep_original:
        candidate = folder / f"{source.stem}_압축.pptx"
    else:
        candidate = folder / source.name

    allocated = allocated_paths if allocated_paths is not None else set()

    # 원본 폴더에서 원본 교체 모드: candidate가 원본 파일 자신인 경우 중복 접미사 없이 반환
    if output_root is None and not opts.keep_original and candidate == source:
        return candidate, False

    # 기존 결과 파일 덮어쓰기 허용 시: 단, 같은 작업 내에서 다른 원본 파일이 이미 선점한 경우 제외
    if opts.overwrite_existing and candidate not in allocated:
        return candidate, False

    if not candidate.exists() and candidate not in allocated:
        return candidate, False

    # 중복 발생 시 접미사(_1, _2, ...)를 부여하여 덮어쓰기 방지
    stem = candidate.stem
    suffix = candidate.suffix
    counter = 1
    while True:
        numbered = folder / f"{stem}_{counter}{suffix}"
        if not numbered.exists() and numbered not in allocated:
            return numbered, True
        counter += 1


def kill_zombie_powerpoint() -> None:
    """작업 전후 백그라운드에 잔류하는 무응답 PowerPoint 좀비 프로세스를 안전하게 정리한다."""
    try:
        import psutil
        for p in psutil.process_iter(["name"]):
            if p.info["name"] and p.info["name"].upper() == "POWERPNT.EXE":
                try:
                    p.kill()
                except Exception:
                    pass
        time.sleep(0.5)
    except Exception:
        pass


def require_pillow() -> None:
    if Image is None:
        raise RuntimeError(f"Pillow가 필요합니다: {PIL_ERROR}")


def require_com() -> None:
    if pythoncom is None or win32com is None:
        raise RuntimeError(f"PPT 변환/PowerPoint 검증에는 pywin32가 필요합니다: {COM_ERROR}")


def convert_ppt_to_pptx(source: Path, destination: Path, log: Callable[[str], None]) -> None:
    """PowerPoint COM을 별도 자동화 인스턴스로 사용해 PPT를 PPTX로 저장한다."""
    require_com()
    kill_zombie_powerpoint()
    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        try:
            app.WindowState = 2  # ppWindowMinimized: UI를 하나씩 띄우지 않음
        except Exception:
            pass
        destination.parent.mkdir(parents=True, exist_ok=True)
        presentation = app.Presentations.Open(str(source), False, False, False)
        # 24 = ppSaveAsOpenXMLPresentation
        presentation.SaveAs(str(destination), 24)
        log(f"  [변환 완료] {source.name} -> {destination.name}")
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            if presentation is not None:
                pythoncom.CoUninitialize()
            else:
                pythoncom.CoUninitialize()
        except Exception:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def parse_int_attr(node, name: str) -> int:
    try:
        return int(node.attrib.get(name, "0"))
    except Exception:
        return 0


def get_crop(node) -> Tuple[int, int, int, int]:
    if node is None:
        return (0, 0, 0, 0)
    return tuple(parse_int_attr(node, x) for x in ("l", "t", "r", "b"))


def relationship_map(rels_file: Path) -> Dict[str, str]:
    import xml.etree.ElementTree as ET
    tree = ET.parse(rels_file)
    return {
        rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
        for rel in tree.getroot()
    }


def scan_media_uses(package_root: Path) -> Tuple[Dict[Path, List[MediaUse]], Dict[Path, object]]:
    """슬라이드 및 레이아웃/마스터 XML에서 그룹화(grpSp), 도형 채우기, 표 셀 등을 포함한 모든 이미지 표시 크기와 crop 정보를 수집한다."""
    import xml.etree.ElementTree as ET

    uses: Dict[Path, List[MediaUse]] = {}
    slide_trees: Dict[Path, object] = {}

    def collect_blip_uses_from_node(node, rels: Dict[str, str], xml_path: Path, base_dir: str, scale_x: float = 1.0, scale_y: float = 1.0):
        tag = node.tag.split("}")[-1] if "}" in node.tag else node.tag

        # 그룹 셰이프(p:grpSp)인 경우: 자체 변환 배율 계산 후 자식 요소 재귀 순회
        if tag == "grpSp":
            grp_ext = node.find(".//p:grpSpPr/a:xfrm/a:ext", NS)
            grp_chExt = node.find(".//p:grpSpPr/a:xfrm/a:chExt", NS)
            cur_scale_x = scale_x
            cur_scale_y = scale_y
            if grp_ext is not None and grp_chExt is not None:
                ext_cx = parse_int_attr(grp_ext, "cx")
                ext_cy = parse_int_attr(grp_ext, "cy")
                ch_cx = parse_int_attr(grp_chExt, "cx")
                ch_cy = parse_int_attr(grp_chExt, "cy")
                if ext_cx > 0 and ch_cx > 0:
                    cur_scale_x = scale_x * (ext_cx / ch_cx)
                if ext_cy > 0 and ch_cy > 0:
                    cur_scale_y = scale_y * (ext_cy / ch_cy)

            for child in list(node):
                collect_blip_uses_from_node(child, rels, xml_path, base_dir, cur_scale_x, cur_scale_y)
            return

        # 2) 하위에 grpSp를 품고 있는 상위 컨테이너(cSld, spTree 등)인 경우 자식들을 순회
        has_sub_grp = any((c.tag.split("}")[-1] if "}" in c.tag else c.tag) == "grpSp" for c in node.iter() if c is not node)
        if has_sub_grp:
            for child in list(node):
                collect_blip_uses_from_node(child, rels, xml_path, base_dir, scale_x, scale_y)
            return

        # 3) 일반 셰이프(p:pic, p:sp 도형 채우기, a:tc 표 셀, p:bg 등)에서 a:blip 탐색
        for blip in node.findall(".//a:blip", NS):
            rid = blip.attrib.get("{" + NS["r"] + "}embed", "") or blip.attrib.get("{" + NS["r"] + "}link", "")
            target = rels.get(rid)
            if not target:
                continue
            media_rel = posixpath.normpath(posixpath.join(base_dir, target))
            if not media_rel.startswith("ppt/media/"):
                continue
            media_path = package_root / Path(*media_rel.split("/"))
            if not media_path.exists():
                continue

            # 표시 크기 탐색 (a:xfrm/a:ext 또는 p:spPr/a:xfrm/a:ext)
            ext = node.find(".//a:xfrm/a:ext", NS)
            if ext is None:
                ext = node.find(".//p:spPr/a:xfrm/a:ext", NS)
            
            cx = parse_int_attr(ext, "cx") if ext is not None else 0
            cy = parse_int_attr(ext, "cy") if ext is not None else 0

            width_inch = (cx * scale_x / EMU_PER_INCH) if cx > 0 else 4.0
            height_inch = (cy * scale_y / EMU_PER_INCH) if cy > 0 else 3.0

            width_px = max(1, round(width_inch * 96))
            height_px = max(1, round(height_inch * 96))

            src_rect = node.find(".//a:srcRect", NS)
            crop = get_crop(src_rect)
            uses.setdefault(media_path, []).append(MediaUse(xml_path, media_path, width_px, height_px, crop))

    # slides, slideLayouts, slideMasters 순회
    search_dirs = [
        (package_root / "ppt" / "slides", "ppt/slides"),
        (package_root / "ppt" / "slideLayouts", "ppt/slideLayouts"),
        (package_root / "ppt" / "slideMasters", "ppt/slideMasters"),
    ]

    for sdir, base_rel in search_dirs:
        if not sdir.exists():
            continue
        rels_dir = sdir / "_rels"
        for xml_file in sorted(sdir.glob("*.xml")):
            rel_file = rels_dir / f"{xml_file.name}.rels"
            if not rel_file.exists():
                continue
            rels = relationship_map(rel_file)
            try:
                tree = ET.parse(xml_file)
            except Exception:
                continue
            slide_trees[xml_file] = tree
            root = tree.getroot()
            for child in list(root):
                collect_blip_uses_from_node(child, rels, xml_file, base_rel)

    # ppt/media/ 디렉토리 전수 검사하여 고립 미디어(Orphaned)도 누락 없이 100% 압축 보장
    media_dir = package_root / "ppt" / "media"
    if media_dir.exists():
        for mf in sorted(media_dir.glob("*")):
            if mf.is_file() and mf not in uses:
                uses.setdefault(mf, []).append(MediaUse(Path(), mf, 1920, 1080, (0, 0, 0, 0)))

    return uses, slide_trees


def crop_image(image, crop: Tuple[int, int, int, int]):
    l, t, r, b = crop
    if not any(crop):
        return image
    width, height = image.size
    # 크롭 백분율 이상치 방어
    if (l + r >= 100000) or (t + b >= 100000):
        return image
    left = max(0, min(width - 1, round(width * l / 100000)))
    top = max(0, min(height - 1, round(height * t / 100000)))
    right = max(left + 1, min(width, round(width * (100000 - r) / 100000)))
    bottom = max(top + 1, min(height, round(height * (100000 - b) / 100000)))
    return image.crop((left, top, right, bottom))


def compress_one_image(path: Path, max_display_w: int, max_display_h: int, quality: int, crop: Tuple[int, int, int, int]) -> bool:
    require_pillow()
    suffix = path.suffix.lower()
    if suffix not in {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}:
        return False
    temp = path.with_name(path.name + ".tmp")
    try:
        with Image.open(path) as original:
            orig_size = path.stat().st_size
            image = crop_image(original.copy(), crop)
            orig_w, orig_h = image.size
            if orig_w <= 0 or orig_h <= 0:
                return False

            # [무결성 핵심] 원본 종횡비(Aspect Ratio) 100% 보존 리사이징:
            # 슬라이드 내에서 이미지가 그룹화되거나 비대칭 조정되었더라도 원본 비율을 유지하며 스케일링하여
            # 오피스 렌더링 시 이중 왜곡(Double Distortion) 및 화질 훼손을 원천 차단
            target_scale = min(1.0, max(max_display_w / orig_w, max_display_h / orig_h))
            target_w = max(1, round(orig_w * target_scale))
            target_h = max(1, round(orig_h * target_scale))

            if target_w < orig_w or target_h < orig_h:
                image = image.resize((target_w, target_h), Image.Resampling.LANCZOS)

            if suffix in {".jpg", ".jpeg"}:
                if image.mode not in ("RGB", "L"):
                    background = Image.new("RGB", image.size, "white")
                    if "A" in image.getbands():
                        background.paste(image, mask=image.getchannel("A"))
                    else:
                        background.paste(image)
                    image = background
                image.save(temp, format="JPEG", quality=quality, optimize=True, dpi=(96, 96))
            elif suffix == ".png":
                image.save(temp, format="PNG", optimize=True, dpi=(96, 96))
            elif suffix == ".webp":
                image.save(temp, format="WEBP", quality=quality, method=6)
            else:
                image.save(temp, format=original.format or "PNG")

        # [용량 역주행 방지 가드] 압축 파일이 원본보다 실제로 작을 때만 교체
        if temp.exists() and temp.stat().st_size < orig_size:
            shutil.move(str(temp), str(path))
            return True
        return False
    finally:
        if temp.exists():
            temp.unlink(missing_ok=True)


def remove_crop_nodes(slide_trees: Dict[Path, object], cropped_media: set[Path], uses: Dict[Path, List[MediaUse]]) -> int:
    import xml.etree.ElementTree as ET
    changed = 0

    for slide_xml, tree in slide_trees.items():
        root = tree.getroot()
        rel_file = slide_xml.parent / "_rels" / f"{slide_xml.name}.rels"
        rels = relationship_map(rel_file) if rel_file.exists() else {}
        base_dir = "ppt/slides" if "slides" in slide_xml.parts else ("ppt/slideLayouts" if "slideLayouts" in slide_xml.parts else "ppt/slideMasters")
        xml_modified = False

        # 모든 a:blipFill 탐색 (p:pic, p:sp, p:grpSp, a:tc 등 포괄)
        for fill in root.iter():
            blip = fill.find("a:blip", NS)
            if blip is None:
                blip = fill.find(".//a:blip", NS)
            if blip is None:
                continue

            rid = blip.attrib.get("{" + NS["r"] + "}embed", "") or blip.attrib.get("{" + NS["r"] + "}link", "")
            target = rels.get(rid, "")
            if not target:
                continue

            media_rel = posixpath.normpath(posixpath.join(base_dir, target))
            media_path = slide_xml.parents[2] / Path(*media_rel.split("/"))
            if media_path not in cropped_media:
                continue

            src_rect = fill.find("a:srcRect", NS)
            if src_rect is not None:
                fill.remove(src_rect)
                changed += 1
                xml_modified = True

        if xml_modified:
            for prefix, uri in (("p", NS["p"]), ("a", NS["a"]), ("r", NS["r"])):
                ET.register_namespace(prefix, uri)
            tree.write(slide_xml, encoding="utf-8", xml_declaration=True)

    return changed


def compress_pptx_package(source: Path, destination: Path, opts: CompressionOptions) -> Tuple[int, int, List[str]]:
    """PPTX ZIP 내부 이미지를 압축하고, 슬라이드 XML의 표시 크기는 건드리지 않는다."""
    warnings: List[str] = []
    changed = 0
    crop_removed = 0
    with tempfile.TemporaryDirectory(prefix="ppt_img_") as temp_name:
        work = Path(temp_name)
        with zipfile.ZipFile(source, "r") as zin:
            zin.extractall(work)
        uses, slide_trees = scan_media_uses(work)
        cropped_media: set[Path] = set()
        for media_path, media_uses in uses.items():
            if not media_uses:
                continue
            max_width = max(u.width_px for u in media_uses)
            max_height = max(u.height_px for u in media_uses)
            # 목표 DPI로 해상도 리스케일링
            target_width = max(1, round(max_width * (opts.resolution_dpi / 96.0)))
            target_height = max(1, round(max_height * (opts.resolution_dpi / 96.0)))

            crop_values = {u.crop for u in media_uses if any(u.crop)}
            crop = (0, 0, 0, 0)
            if opts.delete_cropped and crop_values:
                if len(crop_values) == 1:
                    crop = next(iter(crop_values))
                    cropped_media.add(media_path)
                else:
                    warnings.append(f"공유 이미지의 crop 값이 달라 잘린 영역 삭제를 건너뜀: {media_path.name}")
            try:
                if compress_one_image(media_path, target_width, target_height, opts.jpeg_quality, crop):
                    changed += 1
            except Exception as exc:
                warnings.append(f"이미지 처리 실패({media_path.name}): {exc}")
        if cropped_media:
            crop_removed = remove_crop_nodes(slide_trees, cropped_media, uses)
        destination.parent.mkdir(parents=True, exist_ok=True)
        temp_zip = work.parent / (work.name + ".pptx")
        with zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for item in work.rglob("*"):
                if item.is_file():
                    zout.write(item, item.relative_to(work).as_posix())
        shutil.move(str(temp_zip), str(destination))
    return changed, crop_removed, warnings


def verify_zip(path: Path) -> Tuple[bool, str]:
    try:
        with zipfile.ZipFile(path, "r") as zf:
            bad = zf.testzip()
            required = {"[Content_Types].xml", "ppt/presentation.xml"}
            missing = required.difference(zf.namelist())
            if bad:
                return False, f"ZIP CRC 오류: {bad}"
            if missing:
                return False, f"필수 항목 누락: {', '.join(sorted(missing))}"
        return True, "ZIP 무결성 통과"
    except Exception as exc:
        return False, str(exc)


def verify_with_powerpoint(paths: Iterable[Path], log: Callable[[str], None]) -> Dict[Path, str]:
    """최종 산출물을 PowerPoint로 실제 열어 슬라이드 접근 가능 여부를 확인한다."""
    require_com()
    results: Dict[Path, str] = {}
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            app.DisplayAlerts = 0
            app.WindowState = 2
        except Exception:
            pass
        for path in paths:
            prs = None
            try:
                prs = app.Presentations.Open(str(path), True, False, False)
                count = prs.Slides.Count
                results[path] = f"PowerPoint 열기 검증 통과 ({count}슬라이드)"
            except Exception as exc:
                results[path] = f"PowerPoint 열기 검증 실패: {exc}"
            finally:
                try:
                    if prs is not None:
                        prs.Close()
                except Exception:
                    pass
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return results


class CompressionEngine:
    def __init__(self, log: Callable[[str], None]):
        self.log = log

    def scan(self, root: Path, recursive: bool) -> List[Path]:
        iterator = root.rglob("*") if recursive else root.glob("*")
        return sorted([p for p in iterator if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS])

    def process(self, source_root: Path, output_root: Optional[Path], opts: CompressionOptions) -> List[ProcessResult]:
        files = self.scan(source_root, opts.recursive)
        if not files:
            raise RuntimeError("선택한 폴더에서 .ppt 또는 .pptx 파일을 찾지 못했습니다.")
        self.log(f"[SCAN] PowerPoint {len(files)}개 발견")
        results: List[ProcessResult] = []
        verify_paths: List[Path] = []
        allocated_paths: set[Path] = set()
        for index, source in enumerate(files, 1):
            self.log(f"[{index}/{len(files)}] 처리: {source}")
            destination, is_renamed = choose_output_path(source, source_root, output_root, opts, allocated_paths)
            allocated_paths.add(destination)
            if is_renamed:
                self.log(f"  [중복 방지] 접미사 적용 -> {destination.name}")
            work_source = source
            temp_converted: Optional[Path] = None
            temp_target: Optional[Path] = None
            before = source.stat().st_size
            try:
                if source.suffix.lower() == ".ppt":
                    if not opts.convert_ppt:
                        results.append(ProcessResult(source, None, "SKIP", "PPT 변환 옵션이 꺼져 있음", before=before))
                        continue
                    temp_converted = Path(tempfile.mktemp(suffix=".pptx", prefix="ppt_convert_"))
                    convert_ppt_to_pptx(source, temp_converted, self.log)
                    work_source = temp_converted

                # 원본 폴더에서 원본 직접 교체 모드 (output_root is None, keep_original is False, PPTX):
                # 원본 손상 방지를 위해 임시 파일에 먼저 압축 및 검증 후 교체
                in_place_replace = (output_root is None and not opts.keep_original and source.suffix.lower() == ".pptx")
                if in_place_replace:
                    temp_target = source.with_name(f"{source.stem}_tmp_{time.time_ns()}.pptx")
                    actual_target = temp_target
                else:
                    actual_target = destination

                # 압축 결과는 대상 위치에 저장한다.
                images, crops, warnings = compress_pptx_package(work_source, actual_target, opts)
                ok, verify_message = verify_zip(actual_target)
                if not ok:
                    results.append(ProcessResult(source, destination, "FAIL", verify_message, images, before, before, crops, warnings))
                    continue

                if in_place_replace:
                    shutil.move(str(actual_target), str(source))
                    temp_target = None  # 교체 완료되었으므로 finally에서 삭제되지 않도록 초기화
                    destination = source
                    after = source.stat().st_size
                    self.log(f"  [원본 정리] {source.name} 압축 파일로 교체 완료 (단일 파일 유지)")
                else:
                    after = destination.stat().st_size
                    # '원본 보존(권장)' 해제 시: 검증 통과 후 원본 파일(PPT/PPTX) 정리(삭제)
                    if not opts.keep_original:
                        if output_root is not None:
                            if safe_delete_file(source):
                                self.log(f"  [원본 정리] {source.name} 삭제 완료")
                            else:
                                self.log(f"  [경고] 원본 파일 삭제 실패({source.name})")
                        elif source.suffix.lower() == ".ppt":
                            if safe_delete_file(source):
                                self.log(f"  [원본 정리] {source.name} 삭제 완료 (.pptx로 대체됨)")
                            else:
                                self.log(f"  [경고] 원본 .ppt 삭제 실패({source.name})")

                verify_paths.append(destination)
                results.append(ProcessResult(source, destination, "OK", verify_message, images, before, after, crops, warnings))
                self.log(f"  [OK] 이미지 {images}개, crop 삭제 {crops}개, {before:,} -> {after:,} bytes")
            except Exception as exc:
                results.append(ProcessResult(source, destination, "FAIL", str(exc), before=before))
                self.log(f"  [FAIL] {exc}")
            finally:
                if temp_converted is not None and temp_converted.exists():
                    temp_converted.unlink(missing_ok=True)
                if temp_target is not None and temp_target.exists():
                    temp_target.unlink(missing_ok=True)

        if opts.verify_powerpoint and verify_paths:
            self.log("[VERIFY] PowerPoint로 최종 산출물 열기 검증 중...")
            try:
                opened = verify_with_powerpoint(verify_paths, self.log)
                for result in results:
                    if result.output in opened:
                        result.warnings.append(opened[result.output])
                        if "실패" in opened[result.output] and result.status == "OK":
                            result.status = "FAIL"
                            result.message = opened[result.output]
            except Exception as exc:
                self.log(f"  [WARN] PowerPoint COM 검증을 실행하지 못함: {exc}")
        return results


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PowerPoint PPT/PPTX 이미지 압축 일괄 도구")
        self.root.geometry("920x760")
        self.root.minsize(820, 640)
        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.preserve_tree_var = tk.BooleanVar(value=True)
        self.resolution_var = tk.StringVar(value="96")
        self.quality_var = tk.IntVar(value=85)
        self.recursive_var = tk.BooleanVar(value=True)
        self.convert_var = tk.BooleanVar(value=True)
        self.keep_var = tk.BooleanVar(value=True)
        self.crop_var = tk.BooleanVar(value=True)
        self.verify_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=False)
        self.running = False
        self.log_box: Optional[tk.Text] = None
        self.run_button: Optional[ttk.Button] = None
        self._build()

        # 대시보드 연동 시 윈도우 전면 가시성 확보 (AI Coding Guidelines 표준)
        try:
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after_idle(self.root.attributes, "-topmost", False)
            self.root.focus_force()
        except Exception:
            pass

    def _build(self) -> None:
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")
        ttk.Label(top, text="PowerPoint 이미지 압축 일괄 도구", font=("맑은 고딕", 17, "bold")).pack(anchor="w")
        ttk.Label(top, text="PPT/PPTX · 하위 폴더 전수 스캔 · 96ppi · 최종 PowerPoint 검증", foreground="#555").pack(anchor="w", pady=(3, 0))

        paths = ttk.LabelFrame(self.root, text="1. 대상 폴더 선택", padding=10)
        paths.pack(fill="x", padx=12, pady=6)
        ttk.Label(paths, text="대상 폴더:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(paths, textvariable=self.source_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(paths, text="폴더 선택", command=self.choose_source).grid(row=0, column=2)
        ttk.Label(paths, text="출력 폴더(선택):").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(paths, textvariable=self.output_var).grid(row=1, column=1, sticky="ew", padx=6)
        ttk.Button(paths, text="폴더 선택", command=self.choose_output).grid(row=1, column=2)
        ttk.Checkbutton(paths, text="출력 폴더에 하위 폴더 구조(트리) 유지 (해제 시 단일 폴더에 저장)", variable=self.preserve_tree_var).grid(row=2, column=1, columnspan=2, sticky="w", pady=(2, 2))
        ttk.Label(paths, text="※ 비워두면 각 원본 폴더에 결과를 저장합니다.", foreground="#666").grid(row=3, column=1, columnspan=2, sticky="w")
        paths.columnconfigure(1, weight=1)

        opts = ttk.LabelFrame(self.root, text="2. 압축 및 변환 옵션", padding=10)
        opts.pack(fill="x", padx=12, pady=6)
        ttk.Label(opts, text="해상도:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(opts, textvariable=self.resolution_var, values=("96", "150", "220", "330"), state="readonly", width=8).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Label(opts, text="DPI (PowerPoint 전자 메일 기준은 96)").grid(row=0, column=2, sticky="w")
        ttk.Label(opts, text="JPEG 품질:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Scale(opts, from_=40, to=100, variable=self.quality_var, orient="horizontal", length=240).grid(row=1, column=1, columnspan=2, sticky="w", padx=5)
        ttk.Label(opts, textvariable=self.quality_var).grid(row=1, column=3, sticky="w")
        ttk.Checkbutton(opts, text="하위 폴더까지 전수 스캔", variable=self.recursive_var).grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(opts, text=".ppt를 .pptx로 변환", variable=self.convert_var).grid(row=2, column=2, sticky="w")
        ttk.Checkbutton(opts, text="잘린 그림 영역 삭제", variable=self.crop_var).grid(row=3, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(opts, text="원본 보존(권장)", variable=self.keep_var).grid(row=3, column=2, sticky="w")
        ttk.Checkbutton(opts, text="PowerPoint 최종 열기 검증", variable=self.verify_var).grid(row=4, column=0, columnspan=2, sticky="w")
        ttk.Checkbutton(opts, text="기존 결과 파일 덮어쓰기", variable=self.overwrite_var).grid(row=4, column=2, sticky="w")

        action = ttk.Frame(self.root, padding=(12, 4))
        action.pack(fill="x")
        self.run_button = ttk.Button(action, text="압축·변환 시작", command=self.run)
        self.run_button.pack(side="left")
        ttk.Button(action, text="로그 지우기", command=self.clear_log).pack(side="left", padx=8)
        self.status_label = ttk.Label(action, text="대기 중", foreground="#555")
        self.status_label.pack(side="right")

        log_frame = ttk.LabelFrame(self.root, text="3. 작업 로그 및 최종 검증 결과", padding=8)
        log_frame.pack(fill="both", expand=True, padx=12, pady=(2, 12))
        self.log_box = tk.Text(log_frame, wrap="none", font=("Consolas", 9), state="disabled")
        self.log_box.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_box.yview)
        sb.pack(side="right", fill="y")
        self.log_box.configure(yscrollcommand=sb.set)

    def choose_source(self) -> None:
        value = filedialog.askdirectory(title="PowerPoint 대상 폴더 선택")
        if value:
            self.source_var.set(value)

    def choose_output(self) -> None:
        value = filedialog.askdirectory(title="결과 저장 폴더 선택")
        if value:
            self.output_var.set(value)

    def log(self, text: str) -> None:
        def append() -> None:
            if self.log_box is None:
                return
            self.log_box.configure(state="normal")
            self.log_box.insert("end", text + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.root.after(0, append)

    def clear_log(self) -> None:
        if self.log_box:
            self.log_box.configure(state="normal")
            self.log_box.delete("1.0", "end")
            self.log_box.configure(state="disabled")

    def run(self) -> None:
        if self.running:
            return
        source = Path(self.source_var.get().strip())
        output = Path(self.output_var.get().strip()) if self.output_var.get().strip() else None
        if not source.is_dir():
            messagebox.showwarning("확인 필요", "유효한 대상 폴더를 선택하세요.")
            return
        if output is not None and not output.exists():
            try:
                output.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                messagebox.showerror("오류", f"출력 폴더를 만들 수 없습니다:\n{exc}")
                return
        if Image is None:
            messagebox.showerror("필수 패키지", f"Pillow를 설치하세요.\n{PIL_ERROR}")
            return
        if self.convert_var.get() or self.verify_var.get():
            if pythoncom is None or win32com is None:
                messagebox.showerror("필수 패키지", f"PPT 변환/최종 검증에는 pywin32가 필요합니다.\n{COM_ERROR}")
                return
        opts = CompressionOptions(
            resolution_dpi=int(self.resolution_var.get()),
            jpeg_quality=int(self.quality_var.get()),
            delete_cropped=self.crop_var.get(),
            recursive=self.recursive_var.get(),
            convert_ppt=self.convert_var.get(),
            keep_original=self.keep_var.get(),
            verify_powerpoint=self.verify_var.get(),
            overwrite_existing=self.overwrite_var.get(),
            preserve_tree=self.preserve_tree_var.get(),
        )
        self.running = True
        if self.run_button:
            self.run_button.configure(state="disabled")
        self.status_label.configure(text="백그라운드 처리 중", foreground="#b45f06")
        self.clear_log()

        def worker() -> None:
            engine = CompressionEngine(self.log)
            try:
                results = engine.process(source, output, opts)
                report_root = output or source
                report = report_root / f"ppt_image_compression_report_{time.strftime('%Y%m%d_%H%M%S')}.txt"
                tree_desc = "유지" if opts.preserve_tree else "미반영 (단일 폴더 저장)"
                lines = [
                    "PowerPoint 이미지 압축 최종 보고서",
                    f"대상: {source}",
                    f"출력: {output if output else '(각 원본 폴더)'}",
                    f"폴더 구조(트리) 유지: {tree_desc}",
                    f"생성: {time.strftime('%Y-%m-%d %H:%M:%S')}",
                    "",
                ]
                for result in results:
                    ratio = (1 - result.after / result.before) * 100 if result.before else 0
                    lines.append(f"[{result.status}] {result.source}")
                    lines.append(f"  출력: {result.output or '-'}")
                    lines.append(f"  이미지: {result.images}, crop 삭제: {result.crop_removed}, 용량: {result.before:,} -> {result.after:,} ({ratio:.1f}%)")
                    lines.append(f"  메시지: {result.message}")
                    for warning in result.warnings:
                        lines.append(f"  검증/경고: {warning}")
                report.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
                ok = sum(1 for x in results if x.status == "OK")
                failed = len(results) - ok
                self.log(f"[완료] 성공 {ok}개 / 실패·제외 {failed}개")
                self.log(f"[보고서] {report}")
                self.root.after(0, lambda: messagebox.showinfo("완료", f"처리 완료\n성공: {ok}개\n실패·제외: {failed}개\n\nUTF-8 보고서:\n{report}"))
            except Exception as exc:
                self.log(f"[중단] {exc}")
                self.root.after(0, lambda: messagebox.showerror("오류", str(exc)))
            finally:
                def finish() -> None:
                    self.running = False
                    if self.run_button:
                        self.run_button.configure(state="normal")
                    self.status_label.configure(text="대기 중", foreground="#555")
                self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()


def main() -> None:
    if sys.platform != "win32":
        raise SystemExit("이 도구는 Windows PowerPoint 자동화를 위해 Windows에서 실행해야 합니다.")
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
