#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
HỆ SINH THÁI KHÉP KÍN BÓC TÁCH ĐỀ THI HÓA HỌC CHUẨN GDPT 2018
MODULE 2: mtef_parser.py (Pure Python MathType MTEF to LaTeX Engine - 0 Token)
--------------------------------------------------------------------------------
Tính năng cốt lõi:
1. Giải mã cấu trúc nhị phân OLE2/CFBF từ word/embeddings/oleObjectX.bin.
2. Trích xuất luồng byte MTEF v3/v5 (Equation Native / Ole10Native).
3. Phân tích cây ngữ pháp nhị phân MTEF: CHAR, LINE, TMPL (sub/sup/frac/arrow).
4. Ánh xạ ký hiệu hóa học chuẩn: chữ đứng \\mathrm{...}, mũi tên, liên kết, số mũ.
5. 100% Python thuần, 0 Token AI, 0 phụ thuộc thư viện ngoài (Zero-dependency).
================================================================================
"""


import os
import io
import struct
import re
from typing import Optional, Dict, List, Tuple, Any




# ==============================================================================
# 1. CẤU HÌNH & HẰNG SỐ BẤT BIẾN (MTEF & CFBF INVARIANTS)
# ==============================================================================
CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


# Bảng mã ký tự Symbol / MT Extra sang LaTeX chuyên biệt cho Hóa Học GDPT 2018
SYMBOL_TO_LATEX: Dict[int, str] = {
    # Mũi tên và ký hiệu phản ứng hóa học
    0x2192: r"\rightarrow ",       # Mũi tên phản ứng một chiều
    0x21CC: r"\rightleftharpoons ",# Mũi tên phản ứng thuận nghịch
    0x2191: r"\uparrow ",          # Khí bay lên
    0x2193: r"\downarrow ",        # Kết tủa
    0x2261: r"\equiv ",            # Liên kết ba (alkyne)
    0x2212: "-",                   # Dấu trừ / Liên kết đơn hóa học
    0x00B0: r"^\circ ",            # Độ (nhiệt độ t°)
    0x00D7: r"\times ",            # Dấu nhân
    0x00F7: r"\div ",              # Dấu chia
    0x22C5: r"\cdot ",             # Dấu chấm nhân / tinh thể ngậm nước
    0x2248: r"\approx ",           # Xấp xỉ
    0x2260: r"\neq ",              # Khác
    0x2264: r"\le ",               # Nhỏ hơn hoặc bằng
    0x2265: r"\ge ",               # Lớn hơn hoặc bằng
    0x00B1: r"\pm ",               # Cộng trừ


    # Chữ cái Hy Lạp thường dùng trong Hóa học (đồng phân alpha, beta, gamma, epsilon...)
    0x03B1: r"\alpha",             # alpha (ví dụ: alpha-amino acid)
    0x03B2: r"\beta",              # beta
    0x03B3: r"\gamma",             # gamma
    0x03B4: r"\delta",             # delta
    0x03B5: r"\varepsilon",        # epsilon
    0x03C0: r"\pi",                # liên kết pi
    0x03C3: r"\sigma",             # liên kết sigma
    0x03C9: r"\omega",             # omega
    0x0394: r"\Delta",             # Biến thiên enthalpy Delta H, nhiệt độ


    # Legacy Symbol Font Mappings (khi MathType lưu mã 8-bit / Symbol bias 0xF000)
    0xF02B: "+",
    0xF02D: "-",
    0xF03D: "=",
    0xF0AE: r"\rightarrow ",
    0xF0DC: r"\rightleftharpoons ",
    0xF061: r"\alpha",
    0xF062: r"\beta",
    0xF067: r"\gamma",
    0xF064: r"\delta",
    0xF065: r"\varepsilon",
    0xF070: r"\pi",
    0xF0B0: r"^\circ ",
}




# ==============================================================================
# 2. BỘ TRÍCH XUẤT NHỊ PHÂN OLE2 / CFBF (COMPOUND FILE EXTRACTOR)
# ==============================================================================
class OLEExtractor:
    """Trích xuất luồng nhị phân Equation Native từ OLE Compound Document."""


    @staticmethod
    def extract_equation_stream(data: bytes) -> Optional[bytes]:
        """
        Tìm và trích xuất dữ liệu nhị phân của Equation Native:
        - Tự động duyệt FAT và Directory table trong CFBF.
        - Có cơ chế Fallback dò quét chữ ký MTEF trực tiếp nếu OLE bị bọc header.
        """
        if not data or len(data) < 32:
            return None


        # 1. Thử giải mã chuẩn CFBF
        if data[:8] == CFB_MAGIC:
            stream = OLEExtractor._read_cfbf_stream(data, ["Equation Native", "\x01Ole10Native"])
            if stream:
                mtef_payload = OLEExtractor._locate_mtef_signature(stream)
                if mtef_payload:
                    return mtef_payload


        # 2. Fallback: Quét tìm chữ ký MTEF trực tiếp trên toàn bộ khối byte
        return OLEExtractor._locate_mtef_signature(data)


    @staticmethod
    def _read_cfbf_stream(data: bytes, target_names: List[str]) -> Optional[bytes]:
        try:
            sector_shift = struct.unpack("<H", data[30:32])[0]
            sector_size = 1 << sector_shift
            num_fat = struct.unpack("<I", data[44:48])[0]
            mini_sector_shift = struct.unpack("<H", data[32:34])[0]
            mini_sector_size = 1 << mini_sector_shift


            first_dir = struct.unpack("<I", data[48:52])[0]
            mini_fat_start = struct.unpack("<I", data[60:64])[0]


            fat_sector_ids = [
                struct.unpack("<I", data[76 + i * 4 : 80 + i * 4])[0]
                for i in range(min(num_fat, 109))
            ]


            fat: List[int] = []
            for sid in fat_sector_ids:
                if sid >= 0xFFFFFFFC:
                    break
                offset = (sid + 1) * sector_size
                sec_bytes = data[offset : offset + sector_size]
                fat.extend(struct.unpack(f"<{len(sec_bytes)//4}I", sec_bytes))


            def read_chain(start_sid: int) -> bytes:
                chain = []
                curr = start_sid
                while curr < len(fat) and curr < 0xFFFFFFFC:
                    off = (curr + 1) * sector_size
                    chain.append(data[off : off + sector_size])
                    curr = fat[curr]
                return b"".join(chain)


            dir_data = read_chain(first_dir)
            root_entry = dir_data[:128]
            root_start_sector = struct.unpack("<I", root_entry[116:120])[0]
            mini_stream_data = read_chain(root_start_sector)


            mini_fat: List[int] = []
            if mini_fat_start < 0xFFFFFFFC:
                mini_fat_raw = read_chain(mini_fat_start)
                if mini_fat_raw:
                    mini_fat.extend(struct.unpack(f"<{len(mini_fat_raw)//4}I", mini_fat_raw))


            def read_mini_chain(start_sid: int) -> bytes:
                chain = []
                curr = start_sid
                while curr < len(mini_fat) and curr < 0xFFFFFFFC:
                    off = curr * mini_sector_size
                    chain.append(mini_stream_data[off : off + mini_sector_size])
                    curr = mini_fat[curr]
                return b"".join(chain)
            for i in range(0, len(dir_data), 128):
                entry = dir_data[i : i + 128]
                if len(entry) < 128:
                    break
                name_len = struct.unpack("<H", entry[64:66])[0]
                obj_type = entry[66]
                if obj_type not in (1, 2, 5) or name_len <= 2:
                    continue
                name = entry[: name_len - 2].decode("utf-16le", errors="ignore")
                start_sector = struct.unpack("<I", entry[116:120])[0]
                size = struct.unpack("<Q", entry[120:128])[0]


                if obj_type == 2 and any(t.lower() in name.lower() for t in target_names):
                    if size < 4096 and mini_stream_data and mini_fat:
                        stream_raw = read_mini_chain(start_sector)[:size]
                    else:
                        stream_raw = read_chain(start_sector)[:size]
                    # Nếu là Ole10Native, bỏ qua 4 byte độ dài ban đầu
                    if "\x01Ole10Native" in name and len(stream_raw) > 4:
                        return stream_raw[4:]
                    return stream_raw
        except Exception:
            pass
        return None


    @staticmethod
    def _locate_mtef_signature(data: bytes) -> Optional[bytes]:
        """
        Dò tìm chữ ký MTEF Header:
        - Byte 0: Phiên bản MTEF (3 hoặc 5).
        - Byte 1: Nền tảng (0 = Windows, 1 = Mac).
        - Byte 2: Sản phẩm (0 = MathType, 1 = Equation Editor).
        """
        limit = len(data) - 5
        for i in range(limit):
            if data[i] in (3, 5) and data[i + 1] in (0, 1) and data[i + 2] in (0, 1):
                # Kiểm tra tiếp tính hợp lệ của version
                prod_ver = data[i + 3]
                if prod_ver in (3, 4, 5, 6, 7):
                    return data[i:]
        return None




# ==============================================================================
# 3. ĐỘNG CƠ DỊCH NGỮ PHÁP MTEF SANG LATEX (MTEF PARSER)
# ==============================================================================
class MTEFParser:
    """Động cơ phân tích cú pháp MTEF v3/v5 và tạo chuỗi LaTeX Hóa Học."""


    def __init__(self):
        self.stream: Optional[io.BytesIO] = None
        self.version: int = 5


    def parse(self, mtef_data: bytes) -> str:
        """
        Chuyển đổi luồng byte MTEF thành chuỗi LaTeX hoàn chỉnh:
        - Chuẩn hóa chữ đứng \\mathrm{...} cho công thức hóa học.
        - Tự động xử lý phân số, chỉ số trên/dưới, căn thức, mũi tên.
        """
        if not mtef_data or len(mtef_data) < 5:
            return ""


        # Xác định vị trí header MTEF
        clean_mtef = OLEExtractor._locate_mtef_signature(mtef_data)
        if not clean_mtef:
            return ""


        self.version = clean_mtef[0]
        self.stream = io.BytesIO(clean_mtef)


        # 1. Đọc Header 5 bytes
        self.stream.read(5)


        # 2. Đọc chuỗi Application Key (kết thúc bằng byte 0x00)
        while True:
            b = self.stream.read(1)
            if not b or b == b"\x00":
                break


        # 3. Phân tích đệ quy biểu thức
        raw_content = self._parse_line()


        # 4. Hậu xử lý công thức Hóa học sang chuẩn LaTeX $\mathrm{...}$
        return self._format_chemistry_output(raw_content)


    def _parse_line(self) -> str:
        """Phân tích tuần tự một dòng (Line) chứa các ký tự và template lồng nhau."""
        parts: List[str] = []


        while True:
            tag_byte = self.stream.read(1)
            if not tag_byte:
                break
            tag = tag_byte[0] & 0x0F


            if tag == 0:  # END (Kết thúc dòng hoặc template)
                break
            elif tag == 1:  # LINE (Khối dòng con)
                self.stream.read(1)  # Bỏ qua byte options/nudge
                line_str = self._parse_line()
                parts.append(line_str)
            elif tag == 2:  # CHAR (Ký tự đơn lẻ)
                ch = self._parse_char()
                if ch:
                    parts.append(ch)
            elif tag == 3:  # TMPL (Template: phân số, số mũ, chỉ số dưới...)
                tmpl_str = self._parse_template()
                if tmpl_str:
                    parts.append(tmpl_str)
            elif tag in (4, 5):  # PILE, MATRIX (Cột dọc / Ma trận)
                self.stream.read(1)  # Options
            elif tag in (6, 7, 8, 9, 10, 11, 12, 13):
                # Các tag điều khiển, định nghĩa font, màu sắc: bỏ qua an toàn
                self._skip_control_record(tag)
            else:
                # Gặp tag không xác định: dừng dòng an toàn
                break


        return "".join(parts)


    def _parse_char(self) -> str:
        """Đọc và ánh xạ một ký tự MTEF CHAR."""
        opts = self.stream.read(1)
        if not opts:
            return ""


        # Đọc mã ký tự
        if self.version == 5:
            code_bytes = self.stream.read(2)
            if len(code_bytes) < 2:
                return ""
            char_code = struct.unpack("<H", code_bytes)[0]
        else:
            code_byte = self.stream.read(1)
            if not code_byte:
                return ""
            char_code = code_byte[0]


        # Đọc font index
        self.stream.read(1)


        # Ánh xạ qua từ điển chuyên dụng cho Hóa Học
        if char_code in SYMBOL_TO_LATEX:
            return SYMBOL_TO_LATEX[char_code]


        # Xử lý ký tự ASCII in được
        if 32 <= char_code <= 126:
            return chr(char_code)


        return ""


    def _parse_template(self) -> str:
        """Phân tích Template toán/hóa (Phân số, Căn số, Chỉ số trên/dưới)."""
        opts = self.stream.read(1)
        sel_byte = self.stream.read(1)
        if not sel_byte:
            return ""
        selector = sel_byte[0]


        # Selector 0 hoặc 1: Phân số (Fraction)
        if selector in (0, 1):
            num = self._read_enclosed_line()
            den = self._read_enclosed_line()
            return f"\\frac{{{num}}}{{{den}}}"


        # Selector 2: Chỉ số dưới (Subscript - ví dụ: H2O, C6H12O6)
        elif selector == 2:
            sub = self._read_enclosed_line()
            return f"_{{{sub}}}"


        # Selector 3: Chỉ số trên (Superscript - ví dụ: ion Fe3+, điện tích)
        elif selector == 3:
            sup = self._read_enclosed_line()
            return f"^{{{sup}}}"


        # Selector 4: Cả chỉ số dưới và trên (Sub-Superscript - ví dụ: đồng vị hạt nhân)
        elif selector == 4:
            sub = self._read_enclosed_line()
            sup = self._read_enclosed_line()
            return f"_{{{sub}}}^{{{sup}}}"


        # Selector 5: Căn thức (Radical / Square Root)
        elif selector == 5:
            rad = self._read_enclosed_line()
            return f"\\sqrt{{{rad}}}"


        # Selector 6: Dấu ngoặc mở rộng (Parentheses / Brackets)
        elif selector == 6:
            content = self._read_enclosed_line()
            return f"\\left({content}\\right)"


        # Mặc định: Đọc nội dung bên trong an toàn
        else:
            return self._read_enclosed_line()


    def _read_enclosed_line(self) -> str:
        """Đọc an toàn một khối dòng nằm trong một Template."""
        tag_byte = self.stream.read(1)
        if tag_byte and (tag_byte[0] & 0x0F) == 1:
            self.stream.read(1)  # Options
            return self._parse_line()
        return ""


    def _skip_control_record(self, tag: int) -> None:
        """Bỏ qua an toàn các bản ghi điều khiển/font định dạng."""
        try:
            if tag == 8:  # FONT_DEF
                self.stream.read(1)
                while True:
                    b = self.stream.read(1)
                    if not b or b == b"\x00":
                        break
        except Exception:
            pass


    def _format_chemistry_output(self, raw_latex: str) -> str:
        """
        Chuẩn hóa LaTeX theo phong cách hóa học:
        - Đóng gói trong $\\mathrm{...}$ nếu chứa công thức hóa học chữ cái đứng.
        - Bảo toàn các lệnh LaTeX toán học đặc thù (\\frac, \\sqrt, \\rightarrow).
        """
        cleaned = raw_latex.strip()
        if not cleaned:
            return ""


        # Nếu chuỗi đã có cặp dấu $ bao quanh thì giữ nguyên
        if cleaned.startswith("$") and cleaned.endswith("$"):
            return cleaned


        # Kiểm tra xem có cần bọc \mathrm{} không (áp dụng cho công thức hóa học)
        # Nếu có ký tự chữ cái latin (A-Z, a-z) và không phải toàn lệnh toán
        if re.search(r"[A-Za-z]", cleaned) and not cleaned.startswith(r"\mathrm{"):
            # Nếu chứa lệnh phức tạp như \frac, giữ nguyên cấu trúc lệnh nhưng bọc phần tử
            if r"\frac" in cleaned or r"\sqrt" in cleaned:
                return f"${cleaned}$"
            return f"$\\mathrm{{{cleaned}}}$"


        return f"${cleaned}$"




# ==============================================================================
# 4. GIAO DIỆN CÔNG CỘNG (PUBLIC HIGH-LEVEL API)
# ==============================================================================
def convert_ole_to_latex(ole_source: Any) -> str:
    """
    Điểm tiếp nhận chính: Chuyển đổi OLE Object (oleObjectX.bin) sang LaTeX:
    - Nhận đường dẫn tệp (str/PathLike) hoặc bytes thô.
    - Tự động trích xuất MTEF và dịch sang $\\mathrm{...}$.
    """
    try:
        if isinstance(ole_source, (str, os.PathLike)):
            with open(ole_source, "rb") as f:
                raw_bytes = f.read()
        elif isinstance(ole_source, bytes):
            raw_bytes = ole_source
        elif hasattr(ole_source, "read"):
            raw_bytes = ole_source.read()
        else:
            return ""


        m_tex = re.search(rb"TeX Input Language\x00([^\x00]+)\x00", raw_bytes)
        if m_tex:
            tex_str = m_tex.group(1).decode("utf-8", errors="ignore").strip()
            if tex_str:
                return tex_str if tex_str.startswith("$") else f"${tex_str}$"


        mtef_bytes = OLEExtractor.extract_equation_stream(raw_bytes)
        if not mtef_bytes:
            return ""


        parser = MTEFParser()
        return parser.parse(mtef_bytes)
    except Exception:
        return ""




def convert_mtef_bytes_to_latex(mtef_bytes: bytes) -> str:
    """Dịch trực tiếp từ chuỗi byte MTEF đã trích xuất sẵn."""
    parser = MTEFParser()
    return parser.parse(mtef_bytes)




# ==============================================================================
# 5. BỘ KIỂM THỬ THỰC NGHIỆM CÓ BẪY BIÊN (STRESS TEST SUITE)
# ==============================================================================
def run_mtef_stress_test() -> Dict[str, Any]:
    """
    Test Suite kiểm thử toàn diện 5 bẫy biên nhị phân của Module 2:
    1. Bẫy công thức hóa học chỉ số dưới: H2SO4 -> $\\mathrm{H_{2}SO_{4}}$.
    2. Bẫy phân số định lượng: 14/131 -> $\\frac{14}{131}$.
    3. Bẫy ion điện tích chỉ số trên: Fe3+ -> $\\mathrm{Fe^{3+}}$.
    4. Bẫy mũi tên phản ứng: A -> B -> $\\mathrm{A\\rightarrow B}$.
    5. Bẫy dữ liệu OLE rỗng/hỏng (Fail-safe resilience).
    """
    results: Dict[str, str] = {}


    # Helper tạo header MTEF 5
    def create_mtef_header():
        return struct.pack("5B", 5, 0, 0, 5, 0) + b"MathType\x00"


    # TEST CASE 1: H2SO4 (Chỉ số dưới lồng nhau)
    buf1 = io.BytesIO()
    buf1.write(create_mtef_header())
    buf1.write(b"\x01\x00")  # LINE
    buf1.write(b"\x02\x00" + struct.pack("<H", ord("H")) + b"\x00")
    buf1.write(b"\x03\x00\x02")  # TMPL SUB
    buf1.write(b"\x01\x00")
    buf1.write(b"\x02\x00" + struct.pack("<H", ord("2")) + b"\x00")
    buf1.write(b"\x00")  # END
    buf1.write(b"\x02\x00" + struct.pack("<H", ord("S")) + b"\x00")
    buf1.write(b"\x02\x00" + struct.pack("<H", ord("O")) + b"\x00")
    buf1.write(b"\x03\x00\x02")  # TMPL SUB
    buf1.write(b"\x01\x00")
    buf1.write(b"\x02\x00" + struct.pack("<H", ord("4")) + b"\x00")
    buf1.write(b"\x00")  # END
    buf1.write(b"\x00")  # END


    res1 = convert_mtef_bytes_to_latex(buf1.getvalue())
    results["TEST_1_H2SO4_SUBSCRIPT"] = "PASS" if res1 == r"$\mathrm{H_{2}SO_{4}}$" else f"FAIL ({res1})"


    # TEST CASE 2: Phân số \frac{14}{131}
    buf2 = io.BytesIO()
    buf2.write(create_mtef_header())
    buf2.write(b"\x01\x00")
    buf2.write(b"\x03\x00\x00")  # TMPL FRACTION (selector 0)
    buf2.write(b"\x01\x00")      # Num line
    buf2.write(b"\x02\x00" + struct.pack("<H", ord("1")) + b"\x00")
    buf2.write(b"\x02\x00" + struct.pack("<H", ord("4")) + b"\x00")
    buf2.write(b"\x00")
    buf2.write(b"\x01\x00")      # Den line
    buf2.write(b"\x02\x00" + struct.pack("<H", ord("1")) + b"\x00")
    buf2.write(b"\x02\x00" + struct.pack("<H", ord("3")) + b"\x00")
    buf2.write(b"\x02\x00" + struct.pack("<H", ord("1")) + b"\x00")
    buf2.write(b"\x00")
    buf2.write(b"\x00")


    res2 = convert_mtef_bytes_to_latex(buf2.getvalue())
    results["TEST_2_FRACTION_CALC"] = "PASS" if res2 == r"$\frac{14}{131}$" else f"FAIL ({res2})"


    # TEST CASE 3: Ion Fe3+ (Chỉ số trên)
    buf3 = io.BytesIO()
    buf3.write(create_mtef_header())
    buf3.write(b"\x01\x00")
    buf3.write(b"\x02\x00" + struct.pack("<H", ord("F")) + b"\x00")
    buf3.write(b"\x02\x00" + struct.pack("<H", ord("e")) + b"\x00")
    buf3.write(b"\x03\x00\x03")  # TMPL SUP (selector 3)
    buf3.write(b"\x01\x00")
    buf3.write(b"\x02\x00" + struct.pack("<H", ord("3")) + b"\x00")
    buf3.write(b"\x02\x00" + struct.pack("<H", ord("+")) + b"\x00")
    buf3.write(b"\x00")
    buf3.write(b"\x00")


    res3 = convert_mtef_bytes_to_latex(buf3.getvalue())
    results["TEST_3_ION_SUPERSCRIPT"] = "PASS" if res3 == r"$\mathrm{Fe^{3+}}$" else f"FAIL ({res3})"


    # TEST CASE 4: Mũi tên phản ứng hóa học (Symbol 0x2192)
    buf4 = io.BytesIO()
    buf4.write(create_mtef_header())
    buf4.write(b"\x01\x00")
    buf4.write(b"\x02\x00" + struct.pack("<H", ord("A")) + b"\x00")
    buf4.write(b"\x02\x00" + struct.pack("<H", 0x2192) + b"\x00")  # Arrow
    buf4.write(b"\x02\x00" + struct.pack("<H", ord("B")) + b"\x00")
    buf4.write(b"\x00")


    res4 = convert_mtef_bytes_to_latex(buf4.getvalue())
    results["TEST_4_REACTION_ARROW"] = "PASS" if res4 == r"$\mathrm{A\rightarrow B}$" else f"FAIL ({res4})"


    # TEST CASE 5: Fail-safe dữ liệu rỗng / rác
    res5 = convert_ole_to_latex(b"CORRUPTED_BYTES_WITHOUT_MTEF_SIGNATURE")
    results["TEST_5_FAIL_SAFE_EMPTY"] = "PASS" if res5 == "" else "FAIL"


    return results




if __name__ == "__main__":
    print("=== ĐANG CHẠY BỘ KIỂM THỬ NỘI BỘ MODULE 2 (MTEF_PARSER) ===")
    report = run_mtef_stress_test()
    for t_name, status in report.items():
        print(f"[{status}] {t_name}")
    print("===========================================================")