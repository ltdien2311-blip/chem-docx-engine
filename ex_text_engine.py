#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
HỆ SINH THÁI KHÉP KÍN BÓC TÁCH ĐỀ THI HÓA HỌC CHUẨN GDPT 2018
MODULE 1: ex_text_engine.py (Pure Python Engine - 0 Token)
--------------------------------------------------------------------------------
Tính năng cốt lõi:
1. Thuật toán Two-Pointer Paragraph bảo vệ tuyệt đối câu cuối cùng (End[N] = Total).
2. Thuật toán Delimiter Slicing bóc tách phương án A-D và a-d miễn nhiễm dấu câu.
3. Bộ lọc phần cứng (Hardware Junk Filter) loại bỏ 100% tệp rác WMF, EMF, WDP.
4. Gắn mỏ neo hình ảnh chuẩn hóa: [{slug}_IMG_{index:02d}].
5. Phân tích ngữ nghĩa công thức hóa học, nhãn nhận thức [NB], [TH], [VD], [VDC].
================================================================================
"""


import os
import re
import io
import json
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional, Tuple




# ==============================================================================
# 1. CẤU HÌNH & HẰNG SỐ BẤT BIẾN (INVARIANTS & HARDWARE FILTER)
# ==============================================================================
XML_NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v': 'urn:schemas-microsoft-com:vml'
}


VALID_IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.webp'}
JUNK_IMAGE_EXTENSIONS = {'.wmf', '.emf', '.wdp', '.bin', '.xml', '.rels'}


DEFAULT_SLUG = "H12_TongOn_AminoAcid"




# ==============================================================================
# 2. MÔ HÌNH DỮ LIỆU ĐỘC LẬP (DATA MODELS)
# ==============================================================================
@dataclass
class ExtractedImage:
    """Đại diện cho hình ảnh trích xuất từ DOCX sau khi qua bộ lọc phần cứng."""
    anchor_tag: str
    target_filename: str
    extension: str
    data_bytes: bytes
    size_bytes: int
    relationship_id: str




@dataclass
class DocxParagraph:
    """Đoạn văn sau khi quét tuần tự từ document.xml."""
    text: str
    image_anchors: List[str] = field(default_factory=list)
    raw_xml: str = ""




@dataclass
class QuestionOption:
    """Phương án hoặc ý hỏi thành phần."""
    label: str           # 'A', 'B', 'C', 'D' hoặc 'a', 'b', 'c', 'd'
    content: str         # Nội dung phương án
    is_correct: bool     # True nếu có dấu * hoặc gạch chân
    level: Optional[str] = None # Nhãn nhận thức riêng của ý (nếu có)




@dataclass
class QuestionItem:
    """Cấu trúc dữ liệu chuẩn hóa của một câu hỏi."""
    question_number: int
    cognitive_level: Optional[str]       # 'NB', 'TH', 'VD', 'VDC'
    question_type: str                  # 'CHOICE', 'TRUE_FALSE', 'SHORT_ANSWER'
    stem: str                           # Lời dẫn câu hỏi
    options: List[QuestionOption]       # Danh sách phương án
    correct_answer: str                 # Đáp án tóm tắt (e.g. 'B', 'a, c đúng; b, d sai', '177')
    solution: str                       # Lời giải chi tiết
    image_anchors: List[str]            # Mỏ neo hình ảnh gắn với câu
    full_markdown: str                  # Markdown kết xuất chuẩn mực




# ==============================================================================
# 3. ĐỘNG CƠ XỬ LÝ CHÍNH (EX_TEXT_ENGINE)
# ==============================================================================
class ExTextEngine:
    def __init__(self, slug: str = DEFAULT_SLUG):
        self.slug = self._sanitize_slug(slug)
        self.extracted_images: List[ExtractedImage] = []
        self.rels_map: Dict[str, str] = {} # rId -> anchor_tag


    @staticmethod
    def _sanitize_slug(slug: str) -> str:
        clean = re.sub(r'[^a-zA-Z0-9_]', '', slug.strip())
        return clean if clean else DEFAULT_SLUG


    # --------------------------------------------------------------------------
    # 3.1. BỘ LỌC PHẦN CỨNG & QUẢN LÝ MỎ NEO ẢNH (MEDIA PROCESSOR)
    # --------------------------------------------------------------------------
    def _process_relationships(self, zf: zipfile.ZipFile) -> None:
        """
        Quét file word/_rels/document.xml.rels và áp dụng bộ lọc phần cứng:
        - Loại bỏ 100% rác WMF, EMF, WDP.
        - Đánh số mỏ neo tuần tự: [{slug}_IMG_01], [{slug}_IMG_02], ... theo thứ tự xuất hiện tuyến tính trong document.xml.
        """
        self.rels_map.clear()
        self.extracted_images.clear()


        rels_path = 'word/_rels/document.xml.rels'
        if rels_path not in zf.namelist():
            return


        try:
            rels_xml = zf.read(rels_path)
            root = ET.fromstring(rels_xml)
        except Exception:
            return


        img_counter = 1
        # Bước 1: Thu thập tất cả các tài nguyên ảnh hợp lệ (đã lọc rác phần cứng) theo r_id
        valid_media: Dict[str, Dict[str, Any]] = {}
        for elem in root:
            r_id = elem.attrib.get('Id', '')


            target = elem.attrib.get('Target', '')
            r_type = elem.attrib.get('Type', '')


            if 'image' not in r_type.lower() and not target.lower().startswith('media/'):
                continue
            clean_target = target.replace('\\', '/')
            if clean_target.startswith('/'):
                zip_target_path = clean_target.lstrip('/')
            elif clean_target.startswith('media/'):
                zip_target_path = 'word/' + clean_target


            else:


                zip_target_path = 'word/' + clean_target
            ext = os.path.splitext(zip_target_path)[1].lower()


            if ext in JUNK_IMAGE_EXTENSIONS or ext not in VALID_IMAGE_EXTENSIONS:
                continue
            if zip_target_path in zf.namelist():
                raw_bytes = zf.read(zip_target_path)
                valid_media[r_id] = {
                    "raw_bytes": raw_bytes,
                    "ext": ext,
                    "r_id": r_id,
                    "zip_target_path": zip_target_path
                }


        if not valid_media:
            return


        # Bước 2: Quét thứ tự xuất hiện thực tế của ảnh trong word/document.xml (Document Reading Flow)
        ordered_r_ids: List[str] = []
        doc_xml_path = 'word/document.xml'
        if doc_xml_path in zf.namelist():
            try:
                doc_xml_bytes = zf.read(doc_xml_path)
                pattern = re.compile(rb'(?:r:embed|r:id)\s*=\s*"([^"]+)"')
                for match in pattern.finditer(doc_xml_bytes):
                    found_rid = match.group(1).decode('utf-8', errors='ignore')
                    if found_rid in valid_media and found_rid not in ordered_r_ids:
                        ordered_r_ids.append(found_rid)
            except Exception:
                pass


        # Bổ sung các ảnh hợp lệ còn lại nếu không xuất hiện trực tiếp trong document.xml
        for r_id in valid_media:
            if r_id not in ordered_r_ids:
                ordered_r_ids.append(r_id)


        # Bước 3: Đánh số mỏ neo tuần tự 01, 02, 03... theo đúng mạch bài thi
        img_counter = 1
        for r_id in ordered_r_ids:
            media = valid_media[r_id]
            raw_bytes = media["raw_bytes"]
            ext = media["ext"]
            padded_idx = f"{img_counter:02d}"
            anchor_tag = f"[{self.slug}_IMG_{padded_idx}]"
            out_name = f"{self.slug}_IMG_{padded_idx}{ext}"


            img_obj = ExtractedImage(
                anchor_tag=anchor_tag,
                target_filename=out_name,
                extension=ext,
                data_bytes=raw_bytes,
                size_bytes=len(raw_bytes),
                relationship_id=r_id
            )
            self.extracted_images.append(img_obj)
            self.rels_map[r_id] = anchor_tag
            img_counter += 1


    # --------------------------------------------------------------------------
    # 3.2. BÓC TÁCH XML ĐOẠN VĂN (PARAGRAPH EXTRACTOR)
    # --------------------------------------------------------------------------
    def _extract_paragraphs_from_xml(self, xml_bytes: bytes) -> List[DocxParagraph]:
        paragraphs: List[DocxParagraph] = []
        try:
            root = ET.fromstring(xml_bytes)
        except Exception:
            return self._extract_paragraphs_regex(xml_bytes.decode('utf-8', errors='ignore'))


        for p_elem in root.iter(f"{{{XML_NAMESPACES['w']}}}p"):
            p_text_parts: List[str] = []
            p_anchors: List[str] = []


            for blip in p_elem.iter(f"{{{XML_NAMESPACES['a']}}}blip"):
                embed_id = blip.attrib.get(f"{{{XML_NAMESPACES['r']}}}embed") or blip.attrib.get(f"{{{XML_NAMESPACES['r']}}}id")
                if embed_id and embed_id in self.rels_map:
                    anchor = self.rels_map[embed_id]
                    if anchor not in p_anchors:
                        p_anchors.append(anchor)


            for elem in p_elem.iter():
                tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag


                if tag == 'tab':
                    p_text_parts.append('    ')
                elif tag == 'br':
                    p_text_parts.append('\n')
                elif tag == 'r':
                    rPr = elem.find(f"{{{XML_NAMESPACES['w']}}}rPr")
                    is_sub = False
                    is_sup = False
                    is_underline = False


                    if rPr is not None:
                        vert = rPr.find(f"{{{XML_NAMESPACES['w']}}}vertAlign")
                        if vert is not None:
                            val = vert.attrib.get(f"{{{XML_NAMESPACES['w']}}}val", "")
                            if val == 'subscript':
                                is_sub = True
                            elif val == 'superscript':
                                is_sup = True
                        u = rPr.find(f"{{{XML_NAMESPACES['w']}}}u")
                        if u is not None:
                            is_underline = True


                    run_text = ""
                    for t_elem in elem.findall(f"{{{XML_NAMESPACES['w']}}}t"):
                        if t_elem.text:
                            run_text += t_elem.text


                    if run_text:
                        if is_sub:
                            p_text_parts.append(f"_{{{run_text.strip()}}}")
                        elif is_sup:
                            p_text_parts.append(f"^{{{run_text.strip()}}}")
                        else:
                            if is_underline and re.match(r'^\s*([A-Da-d][\.\)])', run_text):
                                run_text = re.sub(r'^(\s*)([A-Da-d][\.\)])', r'\1*\2', run_text)
                            p_text_parts.append(run_text)


            full_p_text = "".join(p_text_parts).strip()
            full_p_text = re.sub(r'[ \t]+', ' ', full_p_text)


            raw_anchors = re.findall(r'\[[a-zA-Z0-9_]+_IMG_\d+\]', full_p_text)
            for ra in raw_anchors:
                if ra not in p_anchors:
                    p_anchors.append(ra)


            if full_p_text or p_anchors:
                paragraphs.append(DocxParagraph(
                    text=full_p_text,
                    image_anchors=p_anchors
                ))


        return paragraphs


    def _extract_paragraphs_regex(self, xml_str: str) -> List[DocxParagraph]:
        paragraphs: List[DocxParagraph] = []
        p_matches = re.findall(r'<w:p(?:\s+[^>]*)?>([\s\S]*?)</w:p>', xml_str)
        for p_content in p_matches:
            anchors = []
            for b_match in re.finditer(r'(?:r:embed|r:id)="([^"]+)"', p_content):
                r_id = b_match.group(1)
                if r_id in self.rels_map and self.rels_map[r_id] not in anchors:
                    anchors.append(self.rels_map[r_id])


            t_matches = re.findall(r'<w:t(?:\s+[^>]*)?>([\s\S]*?)</w:t>', p_content)
            clean_text = "".join(t_matches).strip()
            clean_text = re.sub(r'[ \t]+', ' ', clean_text)


            if clean_text or anchors:
                paragraphs.append(DocxParagraph(text=clean_text, image_anchors=anchors))
        return paragraphs


    # --------------------------------------------------------------------------
    # 3.3. THUẬT TOÁN TWO-POINTER PARAGRAPH (ZERO-REGRESSION QUESTION SLICER)
    # --------------------------------------------------------------------------
    @staticmethod
    def _is_question_header(text: str) -> bool:
        return bool(re.match(r'^\s*Câu\s+\d+[:.\s]', text, re.IGNORECASE))


    def _parse_paragraphs_two_pointer(self, paragraphs: List[DocxParagraph]) -> List[QuestionItem]:
        total_p = len(paragraphs)
        if total_p == 0:
            return []


        question_starts: List[int] = []
        for i, p in enumerate(paragraphs):
            if self._is_question_header(p.text):
                question_starts.append(i)


        if not question_starts:
            question_starts = [0]


        questions: List[QuestionItem] = []
        num_q = len(question_starts)


        for k in range(num_q):
            start_idx = question_starts[k]
            # VÙNG BẤT BIẾN: Điểm cuối của câu cuối cùng là điểm kết thúc tài liệu
            end_idx = question_starts[k + 1] if (k + 1 < num_q) else total_p


            q_paragraphs = paragraphs[start_idx:end_idx]
            q_item = self._build_question_item(q_paragraphs, default_number=k + 1)
            questions.append(q_item)


        return questions


    # --------------------------------------------------------------------------
    # 3.4. BÓC TÁCH NỘI DUNG CÂU HỎI, ĐÁP ÁN & LỜI GIẢI (DELIMITER SLICING)
    # --------------------------------------------------------------------------
    def _build_question_item(self, q_paragraphs: List[DocxParagraph], default_number: int) -> QuestionItem:
        all_anchors: List[str] = []
        for p in q_paragraphs:
            for anc in p.image_anchors:
                if anc not in all_anchors:
                    all_anchors.append(anc)


        header_text = q_paragraphs[0].text if q_paragraphs else ""
        num_match = re.match(r'^\s*Câu\s+(\d+)', header_text, re.IGNORECASE)
        q_num = int(num_match.group(1)) if num_match else default_number


        level_match = re.search(r'\[(NB|TH|VD|VDC)\]', header_text, re.IGNORECASE)
        q_level = level_match.group(1).upper() if level_match else None


        solution_start_idx = -1
        for idx, p in enumerate(q_paragraphs):
            if re.search(r'(?:^|\s)(?:Lời giải|Hướng dẫn giải)[:.\s]?', p.text, re.IGNORECASE):
                solution_start_idx = idx
                break


        body_paragraphs: List[DocxParagraph] = []
        solution_paragraphs: List[DocxParagraph] = []


        if solution_start_idx != -1:
            body_paragraphs = q_paragraphs[:solution_start_idx]
            sol_p = q_paragraphs[solution_start_idx]
            split_sol = re.split(r'(?:Lời giải|Hướng dẫn giải)[:.\s]?', sol_p.text, flags=re.IGNORECASE, maxsplit=1)
            pre_sol_text = split_sol[0].strip()
            post_sol_text = split_sol[1].strip() if len(split_sol) > 1 else ""


            if pre_sol_text:
                body_paragraphs.append(DocxParagraph(text=pre_sol_text, image_anchors=sol_p.image_anchors))


            sol_first_p = DocxParagraph(text=post_sol_text, image_anchors=[])
            solution_paragraphs = [sol_first_p] + q_paragraphs[solution_start_idx + 1:]
        else:
            body_paragraphs = q_paragraphs


        body_text_lines = [p.text for p in body_paragraphs if p.text]
        full_body_text = "\n".join(body_text_lines)


        sol_text_lines = [p.text for p in solution_paragraphs if p.text]
        full_solution_text = "\n".join(sol_text_lines).strip()


        q_type, stem, options, correct_ans = self._classify_and_extract_options(
            full_body_text, full_solution_text, all_anchors
        )


        markdown_str = self._render_question_markdown(
            q_num=q_num,
            level=q_level,
            stem=stem,
            options=options,
            q_type=q_type,
            solution=full_solution_text,
            anchors=all_anchors,
            include_solution=False
        )


        return QuestionItem(
            question_number=q_num,
            cognitive_level=q_level,
            question_type=q_type,
            stem=stem,
            options=options,
            correct_answer=correct_ans,
            solution=full_solution_text,
            image_anchors=all_anchors,
            full_markdown=markdown_str
        )


    def _classify_and_extract_options(
        self, body_text: str, solution_text: str, anchors: List[str]
    ) -> Tuple[str, str, List[QuestionOption], str]:
        """
        Thuật toán Delimiter Slicing (Cắt lát theo mốc phân cách):
        Miễn nhiễm 100% với dấu chấm câu (.), dấu ngoặc đơn lồng nhau ( ),
        và công thức hóa học phức tạp bên trong phương án.
        """
        clean_body = re.sub(r'^\s*Câu\s+\d+[:.\s]*', '', body_text, flags=re.IGNORECASE)
        clean_body = re.sub(r'^\s*\[(NB|TH|VD|VDC)\]\s*', '', clean_body, flags=re.IGNORECASE).strip()


        # 1. KIỂM TRA DẠNG ĐÚNG / SAI (a), b), c), d)) BẰNG DELIMITER SLICING
        tf_delims = list(re.finditer(r'(?:^|\s)(\*?[a-d]\))\s*(?:\[(NB|TH|VD|VDC)\])?\s+', clean_body))


        if len(tf_delims) >= 3:
            stem = clean_body[:tf_delims[0].start()].strip()
            options: List[QuestionOption] = []
            sol_tf_map = self._decode_true_false_solution(solution_text)


            for i in range(len(tf_delims)):
                raw_label = tf_delims[i].group(1).strip()
                is_correct = raw_label.startswith('*')
                letter = raw_label.lstrip('*').rstrip(')').strip()
                lvl = tf_delims[i].group(2).upper() if tf_delims[i].group(2) else None


                start_pos = tf_delims[i].end()
                end_pos = tf_delims[i + 1].start() if i + 1 < len(tf_delims) else len(clean_body)
                content = clean_body[start_pos:end_pos].strip()


                if letter in sol_tf_map:
                    is_correct = sol_tf_map[letter]


                options.append(QuestionOption(
                    label=letter,
                    content=content,
                    is_correct=is_correct,
                    level=lvl
                ))


            correct_letters = [opt.label for opt in options if opt.is_correct]
            wrong_letters = [opt.label for opt in options if not opt.is_correct]
            ans_summary = f"{', '.join(correct_letters)} đúng; {', '.join(wrong_letters)} sai."
            return "TRUE_FALSE", stem, options, ans_summary


        # 2. KIỂM TRA DẠNG TRẮC NGHIỆM 4 PHƯƠNG ÁN (A, B, C, D) BẰNG DELIMITER SLICING
        choice_delims = list(re.finditer(r'(?:^|\s)(\*?[A-D][\.\)])\s+', clean_body))


        if len(choice_delims) >= 3:
            stem = clean_body[:choice_delims[0].start()].strip()
            options: List[QuestionOption] = []
            correct_choice = ""


            for i in range(len(choice_delims)):
                raw_label = choice_delims[i].group(1).strip()
                is_correct = raw_label.startswith('*')
                letter = raw_label.lstrip('*').rstrip('.):').strip()


                start_pos = choice_delims[i].end()
                end_pos = choice_delims[i + 1].start() if i + 1 < len(choice_delims) else len(clean_body)
                # Cắt lát chính xác, loại bỏ dấu chấm cuối câu của phương án nếu có
                content = clean_body[start_pos:end_pos].strip().rstrip('.')


                if is_correct:
                    correct_choice = letter


                options.append(QuestionOption(
                    label=letter,
                    content=content,
                    is_correct=is_correct
                ))


            # Nếu chưa có dấu *, tìm phương án đúng từ Lời giải
            if not correct_choice:
                sol_choice = re.search(r'(?:Chọn|Đáp án|Phương án)\s+([A-D])\b', solution_text, re.IGNORECASE)
                if sol_choice:
                    correct_choice = sol_choice.group(1).upper()
                    for opt in options:
                        if opt.label == correct_choice:
                            opt.is_correct = True


            return "CHOICE", stem, options, correct_choice


        # 3. DẠNG TRẢ LỜI NGẮN / ĐIỀN SỐ (SHORT_ANSWER)
        stem = clean_body.strip()
        options: List[QuestionOption] = []
        short_ans = ""


        sa_match = re.search(r'(?:^|\s)\*?([A-D])\.\s*([\d,\.]+)', stem)
        if sa_match:
            short_ans = sa_match.group(2).strip()
            stem = stem[:sa_match.start()].strip()
            options.append(QuestionOption(label='A', content=short_ans, is_correct=True))
        else:
            num_in_sol = re.search(r'(?:kết quả|bằng|là)\s*[:=]?\s*([\d,\.]+)\b', solution_text, re.IGNORECASE)
            if num_in_sol:
                short_ans = num_in_sol.group(1).strip()


        return "SHORT_ANSWER", stem, options, short_ans


    @staticmethod
    def _decode_true_false_solution(solution_text: str) -> Dict[str, bool]:
        res: Dict[str, bool] = {}
        if not solution_text:
            return res


        true_match = re.search(r'([a-d\s,]+)\s+đúng', solution_text, re.IGNORECASE)
        if true_match:
            letters = re.findall(r'[a-d]', true_match.group(1).lower())
            for l in letters:
                res[l] = True


        false_match = re.search(r'([a-d\s,]+)\s+sai', solution_text, re.IGNORECASE)
        if false_match:
            letters = re.findall(r'[a-d]', false_match.group(1).lower())
            for l in letters:
                res[l] = False


        return res


    # --------------------------------------------------------------------------
    # 3.5. RENDER ĐỊNH DẠNG MARKDOWN KHỚP OUTPUT.TXT CHUẨN MỰC
    # --------------------------------------------------------------------------
    def _render_question_markdown(
        self, q_num: int, level: Optional[str], stem: str,
        options: List[QuestionOption], q_type: str, solution: str, anchors: List[str],
        include_solution: bool = False
    ) -> str:
        parts: List[str] = []
        lvl_str = f"[{level}] " if level else ""


        stem_with_anchors = stem
        for anc in anchors:
            if anc not in stem_with_anchors:
                stem_with_anchors += f" {anc}"


        header_line = f"Câu {q_num}: {lvl_str}{stem_with_anchors}"


        if q_type == "CHOICE":
            opt_strs = []
            for opt in options:
                prefix = "*" if opt.is_correct else ""
                opt_strs.append(f"{prefix}{opt.label}. {opt.content}.")
            parts.append(header_line)
            parts.extend(opt_strs)
        elif q_type == "TRUE_FALSE":
            parts.append(f"{header_line} Xét các phát biểu sau:")
            for opt in options:
                prefix = "*" if opt.is_correct else ""
                sub_lvl = f"[{opt.level}] " if opt.level else ""
                parts.append(f"{prefix}{opt.label}) {sub_lvl}{opt.content}")


            pass
        else: # SHORT_ANSWER
            ans_label = f"A. {options[0].content}" if options else ""
            parts.append(header_line)
            if ans_label:
                parts.append(ans_label)


        if include_solution and solution:
            parts.append("Lời giải")
            sol_lines = [line.strip() for line in solution.split('\n') if line.strip()]
            for s_line in sol_lines:
                if not s_line.startswith('-'):
                    s_line = f"  - {s_line}"
                else:
                    s_line = f"  {s_line}"
                parts.append(s_line)


        return "\n".join(parts)


    # --------------------------------------------------------------------------
    # 3.6. PUBLIC API: TRÍCH XUẤT TỪ FILE DOCX HOẶC TEXT
    # --------------------------------------------------------------------------
    def parse_docx(self, docx_source: Any) -> Dict[str, Any]:
        try:
            if isinstance(docx_source, (str, os.PathLike)):
                with open(docx_source, 'rb') as f:
                    file_bytes = f.read()
            elif isinstance(docx_source, bytes):
                file_bytes = docx_source
            elif hasattr(docx_source, 'read'):
                file_bytes = docx_source.read()
            else:
                raise ValueError("Nguồn DOCX không hợp lệ.")


            with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as zf:
                self._process_relationships(zf)
                doc_xml_name = 'word/document.xml'
                if doc_xml_name not in zf.namelist():
                    raise FileNotFoundError(f"Không tìm thấy {doc_xml_name} bên trong gói DOCX.")
                xml_content = zf.read(doc_xml_name)


            paragraphs = self._extract_paragraphs_from_xml(xml_content)
            questions = self._parse_paragraphs_two_pointer(paragraphs)
            full_markdown = "\n\n".join([q.full_markdown for q in questions])


            return {
                "success": True,
                "slug": self.slug,
                "total_questions": len(questions),
                "total_images": len(self.extracted_images),
                "questions": [asdict(q) for q in questions],
                "images": [
                    {
                        "anchor_tag": img.anchor_tag,
                        "filename": img.target_filename,
                        "extension": img.extension,
                        "size_bytes": img.size_bytes,
                        "relationship_id": img.relationship_id
                    }
                    for img in self.extracted_images
                ],
                "markdown_output": full_markdown,
                "error": None
            }
        except Exception as err:
            return {
                "success": False,
                "slug": self.slug,
                "total_questions": 0,
                "total_images": 0,
                "questions": [],
                "images": [],
                "markdown_output": "",
                "error": f"Lỗi thực thi Module 1 (ExTextEngine): {str(err)}"
            }


    def parse_plain_text(self, text: str) -> Dict[str, Any]:
        try:
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            paragraphs = [DocxParagraph(text=line) for line in lines]
            questions = self._parse_paragraphs_two_pointer(paragraphs)
            full_markdown = "\n\n".join([q.full_markdown for q in questions])


            return {
                "success": True,
                "slug": self.slug,
                "total_questions": len(questions),
                "total_images": 0,
                "questions": [asdict(q) for q in questions],
                "images": [],
                "markdown_output": full_markdown,
                "error": None
            }
        except Exception as err:
            return {
                "success": False,
                "slug": self.slug,
                "total_questions": 0,
                "total_images": 0,
                "questions": [],
                "images": [],
                "markdown_output": "",
                "error": str(err)
            }




# ==============================================================================
# 4. BỘ KIỂM THỬ TỰ ĐỘNG CÓ CÀI BẪY BIÊN (STRESS TEST HARNESS)
# ==============================================================================
def run_internal_stress_test() -> Dict[str, Any]:
    engine = ExTextEngine(slug="H12_TongOn_AminoAcid")
    results = {}


    # TEST CASE 1: Bẫy biên câu cuối và dấu chấm câu ở các phương án
    doc_text_case1 = (
        "Câu 1: [NB] Amino acid (X) có công thức cấu tạo: [H12_TongOn_AminoAcid_IMG_01] Tên gọi của X là\n"
        "A. Alanine. *B. Glycine. C. Valine. D. Lysine.\n"
        "Lời giải\n"
        "- Glycine là amino acid đơn giản nhất.\n"
        "Câu 20: [VD] Phân tích nguyên tố hợp chất X.\n"
        "*a) [NB] Đồng phân 1.\n"
        "*b) [TH] Đồng phân 2.\n"
        "c) [TH] Đồng phân 3.\n"
        "d) [VD] Đồng phân 4.\n"
        "Lời giải a, b đúng; c, d sai.\n"
        "- Đoạn phân tích số 1 cho câu 20.\n"
        "- Kết luận cuối cùng không bao giờ bị cắt ngắn."
    )
    res1 = engine.parse_plain_text(doc_text_case1)
    q1 = res1['questions'][0]
    q20 = res1['questions'][1]


    is_q1_ok = (
        q1['question_type'] == "CHOICE" and
        q1['correct_answer'] == "B" and
        len(q1['options']) == 4
    )
    is_q20_ok = (
        q20['question_number'] == 20 and
        q20['question_type'] == "TRUE_FALSE" and
        "Kết luận cuối cùng không bao giờ bị cắt ngắn" in q20['solution']
    )
    results["TEST_1_CHOICE_AND_LAST_QUESTION"] = "PASS" if (is_q1_ok and is_q20_ok) else "FAIL"


    # TEST CASE 2: Bẫy lọc phần cứng (Hardware Junk Filter WMF/WDP)
    docx_buf = io.BytesIO()
    with zipfile.ZipFile(docx_buf, 'w') as zf:
        rels_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/junk1.wmf"/>
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/photo1.png"/>
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/junk2.wdp"/>
        </Relationships>"""
        zf.writestr("word/_rels/document.xml.rels", rels_xml)
        doc_xml = """<?xml version="1.0" encoding="UTF-8"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <w:body>
                <w:p>
                    <w:r><w:t>Câu 1: [TH] Hình vẽ: </w:t></w:r>
                    <w:r><w:drawing><a:blip r:embed="rId1"/></w:drawing></w:r>
                    <w:r><w:drawing><a:blip r:embed="rId2"/></w:drawing></w:r>
                    <w:r><w:t> A. Đúng *B. Sai C. Khác D. Không rõ</w:t></w:r>
                </w:p>
                <w:p><w:r><w:t>Lời giải - Chỉ giữ PNG, bỏ WMF.</w:t></w:r></w:p>
            </w:body>
        </w:document>"""
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/media/junk1.wmf", b"WMF_JUNK")
        zf.writestr("word/media/photo1.png", b"VALID_PNG_DATA")
        zf.writestr("word/media/junk2.wdp", b"WDP_JUNK")


    res2 = engine.parse_docx(docx_buf.getvalue())
    wmf_filtered = (
        len(res2['images']) == 1 and
        res2['images'][0]['extension'] == '.png' and
        res2['questions'][0]['image_anchors'] == ['[H12_TongOn_AminoAcid_IMG_01]']
    )
    results["TEST_2_HARDWARE_JUNK_FILTER"] = "PASS" if wmf_filtered else "FAIL"


    return results




if __name__ == "__main__":
    print("=== KIỂM THỬ NỘI BỘ PHƯƠNG ÁN A (DELIMITER SLICING) ===")
    test_report = run_internal_stress_test()
    for test_name, status in test_report.items():
        print(f"[{status}] {test_name}")
    print("=========================================================")