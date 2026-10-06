#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================"
HỆ SINH THÁI KHÉP KÍN BÓC TÁCH ĐỀ THI HÓA HỌC CHUẨN GDPT 2018
MODULE 3: main.py (FastAPI Serverless Service - 0 Token)
--------------------------------------------------------------------------------
Tính năng cốt lõi:
1. REST API tích hợp trọn gói Module 1 (ex_text_engine), Module 2 (mtef_parser)
2. Endpoint bóc tách:   và Module 4 (docx_export_engine).


4. Tự động dịch OLE MathType nhị phân sang LaTeX $\mathrm{...}$ 0-Token.   - POST /api/parse-docx: tiếp nhận tệp DOCX trực tiếp qua multipart.
   - POST /api/parse-docx-base64: tiếp nhận chuỗi Base64 từ Google Apps Script.
3. Endpoint xuất ngược DOCX:
   - POST /api/export-docx-base64: tiếp nhận dữ liệu câu hỏi, hình ảnh và sinh
     ngược tệp DOCX chuẩn Bộ GD&ĐT kèm công thức Native Word OMML.


5. Lọc sạch 100% tệp rác WMF/WDP, xuất kèm ảnh PNG/JPG mã hóa Base64 cho GAS tải Drive.
6. Bảo vệ tuyệt đối câu cuối cùng (End[N] = Total) và bóc tách Delimiter Slicing.
================================================================================
"""


import os
import io
import re
import json
import base64
from dataclasses import asdict
import zipfile
from typing import Optional, Dict, Any, List
import xml.etree.ElementTree as ET


# Nhập khẩu các Module nền tảng đã nghiệm thu
from ex_text_engine import (
    ExTextEngine, DocxParagraph, QuestionItem, XML_NAMESPACES,
    VALID_IMAGE_EXTENSIONS, JUNK_IMAGE_EXTENSIONS, DEFAULT_SLUG
)
from mtef_parser import OLEExtractor, MTEFParser, convert_ole_to_latex
from docx_export_engine import (
    export_exam_to_docx_bytes,
    export_exam_to_docx_base64,
    ChemistryDocxExporter,
    run_export_stress_test
)




# ==============================================================================
# 1. BỘ TÍCH HỢP TOÀN DIỆN (INTEGRATED PIPELINE ENGINE)
# ==============================================================================
class IntegratedChemistryParser:
    """
    Điều phối luồng xử lý khép kín:
    DOCX -> Lọc WMF -> Dịch MathType OLE sang LaTeX -> Two-Pointer Slicer -> JSON + Markdown.
    """


    def __init__(self, slug: str = DEFAULT_SLUG):
        self.slug = re.sub(r'[^a-zA-Z0-9_]', '', slug.strip()) or DEFAULT_SLUG
        self.engine = ExTextEngine(slug=self.slug)
        self.ole_map: Dict[str, str] = {}  # rId -> LaTeX string


    def parse_docx_bytes(self, docx_bytes: bytes) -> Dict[str, Any]:
        try:
            with zipfile.ZipFile(io.BytesIO(docx_bytes), 'r') as zf:
                # Bước 1: Quét quan hệ ảnh và áp dụng bộ lọc phần cứng (Module 1)
                self.engine._process_relationships(zf)


                # Bước 2: Quét và giải mã tất cả các khối MathType OLE Object (Module 2)
                self._process_ole_objects(zf)


                # Bước 3: Đọc và phân tích document.xml tích hợp cả ảnh và công thức LaTeX
                doc_xml_name = 'word/document.xml'
                if doc_xml_name not in zf.namelist():
                    raise FileNotFoundError("Không tìm thấy word/document.xml trong tệp DOCX.")
                xml_content = zf.read(doc_xml_name)


                paragraphs = self._extract_integrated_paragraphs(xml_content)


            # Bước 4: Chạy thuật toán Two-Pointer Paragraph bảo vệ câu cuối cùng
            questions = self.engine._parse_paragraphs_two_pointer(paragraphs)


            # Bước 5: Đóng gói Markdown thành phẩm và danh sách ảnh kèm Base64
            full_markdown = "\n\n".join([q.full_markdown for q in questions])


            image_payloads = []
            for img in self.engine.extracted_images:
                image_payloads.append({
                    "anchor_tag": img.anchor_tag,
                    "filename": img.target_filename,
                    "extension": img.extension,
                    "size_bytes": img.size_bytes,
                    "relationship_id": img.relationship_id,
                    "base64_data": base64.b64encode(img.data_bytes).decode('utf-8')
                })


            return {
                "success": True,
                "slug": self.slug,
                "total_questions": len(questions),
                "total_images": len(self.engine.extracted_images),
                "total_mathtype_formulas": len(self.ole_map),
                "questions": [asdict(q) for q in questions],
                "images": image_payloads,
                "markdown_output": full_markdown,
                "error": None
            }
        except Exception as err:
            return {
                "success": False,
                "slug": self.slug,
                "total_questions": 0,
                "total_images": 0,
                "total_mathtype_formulas": 0,
                "questions": [],
                "images": [],
                "markdown_output": "",
                "error": f"Lỗi xử lý bóc tách DOCX: {str(err)}"
            }


    def _process_ole_objects(self, zf: zipfile.ZipFile) -> None:
        """Quét và biên dịch tất cả tệp embeddings/oleObjectX.bin sang LaTeX."""
        self.ole_map.clear()
        rels_path = 'word/_rels/document.xml.rels'
        if rels_path not in zf.namelist():
            return


        try:
            root = ET.fromstring(zf.read(rels_path))
        except Exception:
            return


        for elem in root:
            r_id = elem.attrib.get('Id', '')
            target = elem.attrib.get('Target', '')
            r_type = elem.attrib.get('Type', '')


            # Nhận diện đối tượng OLE
            if 'oleobject' in r_type.lower() or 'embeddings/' in target.lower():
                clean_target = target.replace('\\', '/')
                zip_path = 'word/' + clean_target.lstrip('/') if not clean_target.startswith('word/') else clean_target


                if zip_path in zf.namelist():
                    ole_raw = zf.read(zip_path)
                    latex_str = convert_ole_to_latex(ole_raw)
                    if latex_str:
                        self.ole_map[r_id] = latex_str


    def _extract_integrated_paragraphs(self, xml_bytes: bytes) -> List[DocxParagraph]:
        """Bóc tách đoạn văn, tự động thay thế MathType OLE bằng LaTeX và ảnh bằng Mỏ neo."""
        paragraphs: List[DocxParagraph] = []
        try:
            root = ET.fromstring(xml_bytes)
        except Exception:
            return []


        for p_elem in root.iter(f"{{{XML_NAMESPACES['w']}}}p"):
            p_text_parts: List[str] = []
            p_anchors: List[str] = []


            # 1. Quét mỏ neo ảnh hợp lệ (đã lọc rác WMF)
            for blip in p_elem.iter(f"{{{XML_NAMESPACES['a']}}}blip"):
                embed_id = blip.attrib.get(f"{{{XML_NAMESPACES['r']}}}embed") or blip.attrib.get(f"{{{XML_NAMESPACES['r']}}}id")
                if embed_id and embed_id in self.engine.rels_map:
                    anchor = self.engine.rels_map[embed_id]
                    if anchor not in p_anchors:
                        p_anchors.append(anchor)


            # 2. Duyệt các phần tử con theo thứ tự xuất hiện
            for elem in p_elem.iter():
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag


                if tag == 'tab':
                    p_text_parts.append('    ')
                elif tag == 'br':
                    p_text_parts.append('\n')
                elif tag == 'OLEObject':
                    ole_id = elem.attrib.get(f"{{{XML_NAMESPACES['r']}}}id") or elem.attrib.get('id')
                    if ole_id and ole_id in self.ole_map:
                        p_text_parts.append(f" {self.ole_map[ole_id]} ")
                elif tag == 'blip':
                    embed_id = elem.attrib.get(f"{{{XML_NAMESPACES['r']}}}embed") or elem.attrib.get(f"{{{XML_NAMESPACES['r']}}}id")
                    if embed_id and embed_id in self.engine.rels_map:
                        anchor = self.engine.rels_map[embed_id]
                        p_text_parts.append(f"\n{anchor}\n")
                        if anchor not in p_anchors:
                            p_anchors.append(anchor)
                elif tag == 'imagedata':
                    r_id = elem.attrib.get(f"{{{XML_NAMESPACES['r']}}}id")
                    if r_id and r_id in self.engine.rels_map:
                        anchor = self.engine.rels_map[r_id]
                        p_text_parts.append(f"\n{anchor}\n")
                        if anchor not in p_anchors:
                            p_anchors.append(anchor)
                    ole_id = elem.attrib.get(f"{{{XML_NAMESPACES['r']}}}id") or elem.attrib.get('id')
                    if ole_id and ole_id in self.ole_map:
                        p_text_parts.append(f" {self.ole_map[ole_id]} ")
                elif tag == 'r':
                    rPr = elem.find(f"{{{XML_NAMESPACES['w']}}}rPr")
                    is_sub = (rPr is not None and rPr.find(f"{{{XML_NAMESPACES['w']}}}vertAlign") is not None and rPr.find(f"{{{XML_NAMESPACES['w']}}}vertAlign").attrib.get(f"{{{XML_NAMESPACES['w']}}}val") == 'subscript')
                    is_sup = (rPr is not None and rPr.find(f"{{{XML_NAMESPACES['w']}}}vertAlign") is not None and rPr.find(f"{{{XML_NAMESPACES['w']}}}vertAlign").attrib.get(f"{{{XML_NAMESPACES['w']}}}val") == 'superscript')
                    is_u = (rPr is not None and rPr.find(f"{{{XML_NAMESPACES['w']}}}u") is not None)


                    run_text = "".join([t.text for t in elem.findall(f"{{{XML_NAMESPACES['w']}}}t") if t.text])
                    if run_text:
                        if is_sub:
                            p_text_parts.append(f"_{{{run_text.strip()}}}")
                        elif is_sup:
                            p_text_parts.append(f"^{{{run_text.strip()}}}")
                        else:
                            if is_u and re.match(r'^\s*([A-Da-d][\.\)])', run_text):
                                run_text = re.sub(r'^(\s*)([A-Da-d][\.\)])', r'\1*\2', run_text)
                            p_text_parts.append(run_text)


            raw_p_text = "".join(p_text_parts)
            raw_p_text = re.sub(r"[ \t]+", " ", raw_p_text)
            raw_p_text = re.sub(r" *\n *", "\n", raw_p_text)
            full_text = re.sub(r"\n{3,}", "\n\n", raw_p_text).strip()
            for ra in re.findall(r'\[[a-zA-Z0-9_]+_IMG_\d+\]', full_text):
                if ra not in p_anchors:
                    p_anchors.append(ra)


            if full_text or p_anchors:
                paragraphs.append(DocxParagraph(text=full_text, image_anchors=p_anchors))


        return paragraphs




# ==============================================================================
# 2. KHỞI TẠO DỊCH VỤ FASTAPI REST API (SERVERLESS ENGINE)
# ==============================================================================
try:
    from fastapi import FastAPI, UploadFile, File, Form, HTTPException
    from fastapi.middleware.cors import CORSMiddleware
    from pydantic import BaseModel


    app = FastAPI(
        title="Chemistry Exam Parser & DOCX Exporter Serverless Engine",
        description="Động cơ bóc tách đề thi Hóa học GDPT 2018 kết hợp xuất ngược DOCX chuẩn Bộ GD&ĐT",
        version="2.0.0"
    )


    # Cho phép Google Apps Script và Web Client gọi không bị rào cản CORS
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )


    class Base64ParseRequest(BaseModel):
        base64_docx: str
        slug: Optional[str] = DEFAULT_SLUG


    class ExportDocxRequest(BaseModel):
        questions: List[Dict[str, Any]]
        images: Optional[Dict[str, Any]] = {}
        exam_title: Optional[str] = "ĐỀ KIỂM TRA ĐÁNH GIÁ NĂNG LỰC HÓA HỌC"
        school_name: Optional[str] = "SỞ GIÁO DỤC VÀ ĐÀO TẠO"
        academic_year: Optional[str] = "NĂM HỌC 2026 - 2027"
        duration_minutes: Optional[int] = 50
        include_solutions: Optional[bool] = False
        slug: Optional[str] = DEFAULT_SLUG
        slug: Optional[str] = DEFAULT_SLUG


    @app.get("/")
    def health_check():
        return {
            "status": "online",
            "service": "Chemistry Exam Parser & DOCX Exporter Engine",
            "version": "2.0.0",
            "modules": [
                "Module 1 (ex_text_engine)",
                "Module 2 (mtef_parser)",
                "Module 3 (FastAPI Service)",
                "Module 4 (docx_export_engine)"
            ]
        }


    @app.post("/api/parse-docx")
    async def parse_docx_endpoint(
        file: Optional[UploadFile] = File(None),
        slug: Optional[str] = Form(DEFAULT_SLUG)
    ):


        if not file:
            raise HTTPException(status_code=400, detail="Vui lòng đính kèm tệp DOCX.")
        content = await file.read()
        parser = IntegratedChemistryParser(slug=slug or DEFAULT_SLUG)
        result = parser.parse_docx_bytes(content)
        return result


    @app.post("/api/parse-docx-base64")
    def parse_docx_base64_endpoint(payload: Base64ParseRequest):


        if not payload.base64_docx:
            raise HTTPException(status_code=400, detail="Dữ liệu base64_docx không được để trống.")
        try:
            raw_bytes = base64.b64decode(payload.base64_docx)
        except Exception:
            raise HTTPException(status_code=400, detail="Chuỗi Base64 không hợp lệ.")


        parser = IntegratedChemistryParser(slug=payload.slug or DEFAULT_SLUG)
        result = parser.parse_docx_bytes(raw_bytes)
        return result


    @app.post("/api/export-docx-base64")
    def export_docx_base64_endpoint(payload: ExportDocxRequest):
        if not payload.questions:
            raise HTTPException(status_code=400, detail="Danh sách câu hỏi không được để trống.")


        clean_slug = re.sub(r'[^a-zA-Z0-9_]', '', payload.slug.strip()) if payload.slug else DEFAULT_SLUG


        try:
            b64_docx = export_exam_to_docx_base64(
                questions=payload.questions,
                images=payload.images or {},
                exam_title=payload.exam_title or "ĐỀ KIỂM TRA ĐÁNH GIÁ NĂNG LỰC HÓA HỌC",
                school_name=payload.school_name or "SỞ GIÁO DỤC VÀ ĐÀO TẠO",
                academic_year=payload.academic_year or "NĂM HỌC 2026 - 2027",
                duration_minutes=payload.duration_minutes or 50,
                include_solutions=payload.include_solutions or False
            )


            sol_suffix = "_KemLoiGiai" if payload.include_solutions else ""
            out_filename = f"{clean_slug}_DeThi_GDPT2018{sol_suffix}.docx"


            return {
                "success": True,
                "base64_docx": b64_docx,
                "filename": out_filename,
                "total_questions": len(payload.questions),
                "include_solutions": payload.include_solutions,
                "error": None
            }
        except Exception as err:
            return {
                "success": False,
                "base64_docx": "",
                "filename": "",
                "total_questions": 0,
                "include_solutions": False,
                "error": f"Lỗi tạo tệp DOCX: {str(err)}"
            }
except ImportError:


    app = None




# ==============================================================================
# 3. BỘ KIỂM THỬ THỰC NGHIỆM TÍCH HỢP (INTEGRATION TEST HARNESS)
# ==============================================================================
def run_integration_pipeline_test() -> Dict[str, Any]:
    """Kiểm tra toàn bộ luồng tích hợp: DOCX -> MathType OLE -> LaTeX -> Xuất ngược DOCX OMML."""
    import struct




    hdr = struct.pack("5B", 5, 0, 0, 5, 0) + b"MathType\x00"
    buf = io.BytesIO()
    buf.write(hdr)
    buf.write(b"\x01\x00")
    buf.write(b"\x02\x00" + struct.pack("<H", ord("H")) + b"\x00")
    buf.write(b"\x03\x00\x02")
    buf.write(b"\x01\x00")
    buf.write(b"\x02\x00" + struct.pack("<H", ord("2")) + b"\x00")
    buf.write(b"\x00")
    buf.write(b"\x02\x00" + struct.pack("<H", ord("S")) + b"\x00")
    buf.write(b"\x02\x00" + struct.pack("<H", ord("O")) + b"\x00")
    buf.write(b"\x03\x00\x02")
    buf.write(b"\x01\x00")
    buf.write(b"\x02\x00" + struct.pack("<H", ord("4")) + b"\x00")
    buf.write(b"\x00")
    buf.write(b"\x00")
    mtef_raw = buf.getvalue()


    mock_docx = io.BytesIO()
    with zipfile.ZipFile(mock_docx, 'w') as zf:
        rels = """<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/junk.wmf"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/diagram.png"/>
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>
        </Relationships>"""
        zf.writestr("word/_rels/document.xml.rels", rels)


        doc = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:o="urn:schemas-microsoft-com:office:office"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Câu 1: [NB] Chất </w:t></w:r>
                    <w:object><o:OLEObject r:id="rId3"/></w:object>
                    <w:r><w:t> tác dụng với BaCl2 sinh kết tủa. </w:t></w:r>
                    <w:r><w:drawing><a:blip r:embed="rId1"/></w:drawing></w:r>
                    <w:r><w:drawing><a:blip r:embed="rId2"/></w:drawing></w:r>
                    <w:r><w:t> A. Đúng. *B. Sai. C. Khác. D. Không rõ.</w:t></w:r>
                </w:p>
                <w:p><w:r><w:t>Lời giải - Tạo BaSO4 trắng.</w:t></w:r></w:p>
                <w:p><w:r><w:t>Câu 20: [VD] Câu cuối cùng được bảo toàn tuyệt đối.</w:t></w:r></w:p>
                <w:p><w:r><w:t>*a) Đúng. b) Sai. *c) Đúng. d) Sai.</w:t></w:r></w:p>
                <w:p><w:r><w:t>Lời giải a, c đúng; b, d sai. Dòng cuối không rớt.</w:t></w:r></w:p>
            </w:body>
        </w:document>"""
        zf.writestr("word/document.xml", doc)
        zf.writestr("word/media/junk.wmf", b"JUNK_WMF")
        zf.writestr("word/media/diagram.png", b"REAL_PNG_DATA")
        zf.writestr("word/embeddings/oleObject1.bin", mtef_raw)


    parser = IntegratedChemistryParser(slug="H12_TongOn_AminoAcid")
    res = parser.parse_docx_bytes(mock_docx.getvalue())


    report = {}
    report["PIPELINE_PARSE_SUCCESS"] = "PASS" if res["success"] else "FAIL"
    report["TOTAL_QUESTIONS_2"] = "PASS" if res["total_questions"] == 2 else "FAIL"
    report["MATHTYPE_TRANSLATED"] = "PASS" if r"$\mathrm{H_{2}SO_{4}}$" in res["markdown_output"] else "FAIL"
    report["WMF_FILTERED_OUT"] = "PASS" if res["total_images"] == 1 and res["images"][0]["extension"] == ".png" else "FAIL"
    report["LAST_QUESTION_PROTECTED"] = "PASS" if "Dòng cuối không rớt" in res["questions"][1]["solution"] else "FAIL"


    try:
        mock_img_map = {res["images"][0]["anchor_tag"]: res["images"][0]["base64_data"]} if res["images"] else {}
        exported_b64 = export_exam_to_docx_base64(
            questions=res["questions"],
            images=mock_img_map,
            include_solutions=True
        )
        report["EXPORT_REVERSE_DOCX"] = "PASS" if len(exported_b64) > 5000 else "FAIL"
    except Exception as e:
        report["EXPORT_REVERSE_DOCX"] = f"FAIL ({str(e)})"


    return report




if __name__ == "__main__":
    test_results = run_integration_pipeline_test()
    for test_key, test_status in test_results.items():
        print(f"[{test_status}] {test_key}")
    print("=========================================================")