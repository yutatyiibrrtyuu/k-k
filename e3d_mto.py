"""
E3D 材料统计（MTO）工具
========================
支持读取 AVEVA E3D 导出的管道材料报表（CSV / Excel / 固定宽度文本），
汇总后输出带格式的 Excel 统计表。

支持的元件类型
--------------
PIPE   管子（按长度 m 统计）
ELBOW  弯头
TEE    三通
REDUC  大小头（异径管）
CAP    管帽
FLANGE 法兰
GASKET 垫片
BOLT   螺栓/螺母套组
VALVE  阀门
OLET   支管台（Weldolet / Sockolet …）
OTHER  其他

列名映射（支持 E3D 常见英文列名，以及中文列名）
-----------------------------------------------
LINE_NO / 管线号
ITEM_CODE / ITEM / 元件代码
DESCRIPTION / DESC / 描述
SPEC / 管道等级
SIZE / 管径 / NPS
SIZE2 / 第二管径（异径件）
MATERIAL / MAT / 材质
QUANTITY / QTY / QUAN / 数量
UNIT / 单位
TAG / 位号
AREA / 区域
ISOMETRIC / ISO / 单线图号
"""

from __future__ import annotations

import csv
import io
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

# 可选依赖：openpyxl（输出 Excel）、pandas（读取 Excel 输入）
try:
    import openpyxl
    from openpyxl.styles import (Alignment, Border, Font, PatternFill,
                                  Side)
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


# ---------------------------------------------------------------------------
# 常量 & 映射表
# ---------------------------------------------------------------------------

# 元件类型关键字 → 规范类型名
_ITEM_TYPE_MAP: Dict[str, str] = {
    "PIPE": "PIPE",
    "管子": "PIPE",
    "ELBOW": "ELBOW",
    "弯头": "ELBOW",
    "ELL": "ELBOW",
    "TEE": "TEE",
    "三通": "TEE",
    "REDUC": "REDUCER",
    "REDUCER": "REDUCER",
    "大小头": "REDUCER",
    "异径管": "REDUCER",
    "CAP": "CAP",
    "管帽": "CAP",
    "FLANGE": "FLANGE",
    "法兰": "FLANGE",
    "FLG": "FLANGE",
    "GASKET": "GASKET",
    "垫片": "GASKET",
    "BOLT": "BOLT",
    "螺栓": "BOLT",
    "STUD": "BOLT",
    "VALVE": "VALVE",
    "阀门": "VALVE",
    "GATE": "VALVE",
    "GLOBE": "VALVE",
    "CHECK": "VALVE",
    "BALL": "VALVE",
    "BUTTERFLY": "VALVE",
    "PLUG": "VALVE",
    "NEEDLE": "VALVE",
    "OLET": "OLET",
    "WELDOLET": "OLET",
    "SOCKOLET": "OLET",
    "THREDOLET": "OLET",
    "支管台": "OLET",
}

# 单位规范化
_UNIT_MAP: Dict[str, str] = {
    "M": "M", "METRE": "M", "METER": "M", "米": "M",
    "MM": "MM", "毫米": "MM",
    "NO": "NO", "NOS": "NO", "EA": "NO", "PCS": "NO",
    "PC": "NO", "个": "NO", "件": "NO", "套": "NO", "组": "NO",
    "SET": "SET",
    "KG": "KG", "公斤": "KG",
}

# 列名归一化映射
_COL_ALIASES: Dict[str, str] = {
    # 管线号
    "LINE_NO": "LINE_NO", "LINE": "LINE_NO", "管线号": "LINE_NO", "管线": "LINE_NO",
    # 元件代码
    "ITEM_CODE": "ITEM_CODE", "ITEM": "ITEM_CODE", "CODE": "ITEM_CODE",
    "元件代码": "ITEM_CODE", "物料代码": "ITEM_CODE",
    # 描述
    "DESCRIPTION": "DESC", "DESC": "DESC", "描述": "DESC",
    "COMPONENT": "DESC", "元件描述": "DESC",
    # 规格等级
    "SPEC": "SPEC", "PIPING_SPEC": "SPEC", "PIPE_SPEC": "SPEC",
    "管道等级": "SPEC", "管等级": "SPEC",
    # 管径
    "SIZE": "SIZE", "NPS": "SIZE", "管径": "SIZE",
    "DN": "SIZE", "SIZE1": "SIZE", "BORE": "SIZE",
    # 第二管径
    "SIZE2": "SIZE2", "SECOND_SIZE": "SIZE2", "第二管径": "SIZE2",
    "BORE2": "SIZE2",
    # 材质
    "MATERIAL": "MATERIAL", "MAT": "MATERIAL", "材质": "MATERIAL",
    "MATERIAL_GRADE": "MATERIAL", "MAT_GRADE": "MATERIAL",
    # 数量
    "QUANTITY": "QTY", "QTY": "QTY", "QUAN": "QTY", "数量": "QTY",
    "QTY.": "QTY", "AMOUNT": "QTY",
    # 单位
    "UNIT": "UNIT", "单位": "UNIT", "UOM": "UNIT",
    # 位号/TAG
    "TAG": "TAG", "TAG_NO": "TAG", "位号": "TAG",
    "VALVE_TAG": "TAG", "EQUIP_TAG": "TAG",
    # 区域
    "AREA": "AREA", "ZONE": "AREA", "区域": "AREA",
    # 单线图
    "ISOMETRIC": "ISO", "ISO": "ISO", "ISO_DWG": "ISO",
    "单线图": "ISO", "单线图号": "ISO",
    # 热处理/试压等附加字段
    "PWHT": "PWHT", "热处理": "PWHT",
    "TEST_PRESSURE": "TEST_P", "试验压力": "TEST_P",
}


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

@dataclass
class MTOItem:
    """一条 MTO 记录"""
    line_no: str = ""
    item_code: str = ""
    desc: str = ""
    spec: str = ""
    size: str = ""
    size2: str = ""
    material: str = ""
    qty: float = 0.0
    unit: str = "NO"
    tag: str = ""
    area: str = ""
    iso: str = ""
    item_type: str = "OTHER"   # 归一化类型

    def size_key(self) -> str:
        if self.size2 and self.size2 != self.size:
            return f"{self.size}x{self.size2}"
        return self.size

    def group_key(self) -> tuple:
        """聚合统计用的键"""
        return (
            self.item_type,
            self.item_code,
            self.desc,
            self.spec,
            self.size_key(),
            self.material,
            self.unit,
        )


# ---------------------------------------------------------------------------
# 解析器
# ---------------------------------------------------------------------------

class E3DMTOParser:
    """
    E3D MTO 文件解析器。

    支持三种输入格式：
      1. CSV（逗号或制表符分隔）
      2. Excel（.xlsx / .xls，需 pandas）
      3. 固定宽度文本（E3D 默认打印报表格式）
    """

    def __init__(self, encoding: str = "utf-8-sig"):
        self.encoding = encoding
        self.items: List[MTOItem] = []
        self._warnings: List[str] = []

    # ------------------------------------------------------------------
    # 公开接口
    # ------------------------------------------------------------------

    def parse_file(self, filepath: str | Path) -> List[MTOItem]:
        """自动判断文件格式并解析，返回 MTOItem 列表。"""
        fp = Path(filepath)
        if not fp.exists():
            raise FileNotFoundError(f"文件不存在: {fp}")

        suffix = fp.suffix.lower()
        if suffix in (".xlsx", ".xls"):
            rows, headers = self._read_excel(fp)
        elif suffix == ".csv":
            rows, headers = self._read_csv(fp)
        else:
            rows, headers = self._read_text(fp)

        self.items = self._build_items(rows, headers)
        return self.items

    def parse_text(self, text: str, fmt: str = "csv") -> List[MTOItem]:
        """从字符串直接解析（便于测试）。"""
        if fmt == "csv":
            rows, headers = self._read_csv_text(text)
        else:
            rows, headers = self._read_fixed_text(text)
        self.items = self._build_items(rows, headers)
        return self.items

    @property
    def warnings(self) -> List[str]:
        return self._warnings

    # ------------------------------------------------------------------
    # 内部读取方法
    # ------------------------------------------------------------------

    def _read_csv(self, fp: Path):
        with open(fp, encoding=self.encoding, errors="replace") as f:
            content = f.read()
        return self._read_csv_text(content)

    def _read_csv_text(self, text: str):
        # 自动探测分隔符
        sample = text[:2048]
        dialect = "excel"
        if sample.count("\t") > sample.count(","):
            delimiter = "\t"
        else:
            delimiter = ","

        reader = csv.reader(io.StringIO(text), delimiter=delimiter)
        rows_raw = list(reader)
        # 跳过完全空行和注释行
        rows_clean = [r for r in rows_raw if r and not str(r[0]).startswith("#")]
        if not rows_clean:
            return [], []

        headers = [c.strip().upper() for c in rows_clean[0]]
        data_rows = rows_clean[1:]
        return data_rows, headers

    def _read_excel(self, fp: Path):
        if not HAS_PANDAS:
            raise ImportError("读取 Excel 需要安装 pandas：pip install pandas openpyxl")
        df = pd.read_excel(fp, dtype=str).fillna("")
        headers = [str(c).strip().upper() for c in df.columns]
        rows = df.values.tolist()
        return rows, headers

    def _read_text(self, fp: Path):
        with open(fp, encoding=self.encoding, errors="replace") as f:
            text = f.read()
        return self._read_fixed_text(text)

    def _read_fixed_text(self, text: str):
        """
        解析 E3D 固定宽度文本报表，例如：

          LINE NO   ITEM     SPEC  SIZE   QTY  UNIT  DESCRIPTION
          -------- -------- ----- ----- ----- ----- -----------
          100-P-001 PIPE     CS3   2"    10.5  M     ERW PIPE A106B
        """
        lines = text.splitlines()
        # 找表头行（包含 ITEM 或 SPEC 或 SIZE 等关键词）
        header_idx = None
        for i, ln in enumerate(lines):
            up = ln.upper()
            if ("ITEM" in up or "DESC" in up or "SIZE" in up) and ("QTY" in up or "QUAN" in up or "数量" in up):
                header_idx = i
                break
        if header_idx is None:
            # 尝试自动识别：取第一行非空非分隔线
            for i, ln in enumerate(lines):
                if ln.strip() and not set(ln.strip()) <= set("-=_ \t"):
                    header_idx = i
                    break
        if header_idx is None:
            return [], []

        # 计算列宽（分隔线或空格分列）
        sep_line = ""
        if header_idx + 1 < len(lines):
            sep_line = lines[header_idx + 1]

        if set(sep_line.strip()) <= set("-= "):
            # 有分隔线：用分隔线确定列边界
            col_spans = _parse_col_spans_from_sep(sep_line)
            headers = [lines[header_idx][s:e].strip().upper()
                       for s, e in col_spans]
            data_lines = [ln for ln in lines[header_idx + 2:]
                          if ln.strip() and not set(ln.strip()) <= set("-= ")]
            rows = [[ln[s:e].strip() for s, e in col_spans]
                    for ln in data_lines]
        else:
            # 无分隔线：用空白分割
            headers = lines[header_idx].split()
            headers = [h.upper() for h in headers]
            rows = [ln.split() for ln in lines[header_idx + 1:]
                    if ln.strip()]

        return rows, headers

    # ------------------------------------------------------------------
    # 构建 MTOItem
    # ------------------------------------------------------------------

    def _build_items(self, rows, raw_headers) -> List[MTOItem]:
        # 归一化列名
        col_map: Dict[int, str] = {}
        for idx, h in enumerate(raw_headers):
            norm = _normalize_col(h)
            if norm:
                col_map[idx] = norm

        items = []
        for row_no, row in enumerate(rows, start=2):
            row = [str(c).strip() for c in row]
            if not any(row):
                continue
            kv: Dict[str, str] = {}
            for idx, val in enumerate(row):
                if idx in col_map:
                    kv[col_map[idx]] = val

            item = MTOItem()
            item.line_no = kv.get("LINE_NO", "")
            item.item_code = kv.get("ITEM_CODE", "")
            item.desc = kv.get("DESC", "")
            item.spec = kv.get("SPEC", "")
            item.size = _normalize_size(kv.get("SIZE", ""))
            item.size2 = _normalize_size(kv.get("SIZE2", ""))
            item.material = kv.get("MATERIAL", "")
            item.tag = kv.get("TAG", "")
            item.area = kv.get("AREA", "")
            item.iso = kv.get("ISO", "")

            # 数量
            qty_raw = kv.get("QTY", "0")
            try:
                item.qty = float(re.sub(r"[^\d.\-]", "", qty_raw) or "0")
            except ValueError:
                item.qty = 0.0
                self._warnings.append(f"第 {row_no} 行：数量解析失败 '{qty_raw}'")

            # 单位
            unit_raw = kv.get("UNIT", "NO").upper().strip()
            item.unit = _UNIT_MAP.get(unit_raw, "NO")

            # 元件类型识别
            item.item_type = _detect_type(item.item_code, item.desc)

            if item.qty == 0:
                continue
            items.append(item)

        return items


# ---------------------------------------------------------------------------
# 统计汇总
# ---------------------------------------------------------------------------

@dataclass
class SummaryRow:
    item_type: str
    item_code: str
    desc: str
    spec: str
    size: str
    material: str
    unit: str
    total_qty: float = 0.0
    line_count: int = 0   # 涉及管线数


class MTOStatistics:
    """对 MTOItem 列表进行多维度汇总。"""

    def __init__(self, items: List[MTOItem]):
        self.items = items

    # ------------------------------------------------------------------
    # 总量汇总（按元件类型+规格聚合）
    # ------------------------------------------------------------------

    def summary(self) -> List[SummaryRow]:
        """全量汇总（按 group_key 合并）。"""
        agg: Dict[tuple, SummaryRow] = {}
        line_sets: Dict[tuple, set] = defaultdict(set)

        for it in self.items:
            key = it.group_key()
            if key not in agg:
                agg[key] = SummaryRow(
                    item_type=it.item_type,
                    item_code=it.item_code,
                    desc=it.desc,
                    spec=it.spec,
                    size=it.size_key(),
                    material=it.material,
                    unit=it.unit,
                )
            agg[key].total_qty += it.qty
            if it.line_no:
                line_sets[key].add(it.line_no)

        for key, row in agg.items():
            row.line_count = len(line_sets[key])
            # 保留 2 位小数
            row.total_qty = round(row.total_qty, 3)

        # 按类型 → 规格排序
        return sorted(agg.values(),
                      key=lambda r: (r.item_type, r.spec, r.size, r.item_code))

    # ------------------------------------------------------------------
    # 按管线汇总
    # ------------------------------------------------------------------

    def by_line(self) -> Dict[str, List[SummaryRow]]:
        """每条管线的材料列表。"""
        by_line: Dict[str, List[MTOItem]] = defaultdict(list)
        for it in self.items:
            by_line[it.line_no or "UNKNOWN"].append(it)
        return {
            ln: MTOStatistics(items).summary()
            for ln, items in sorted(by_line.items())
        }

    # ------------------------------------------------------------------
    # 按区域汇总
    # ------------------------------------------------------------------

    def by_area(self) -> Dict[str, List[SummaryRow]]:
        by_area: Dict[str, List[MTOItem]] = defaultdict(list)
        for it in self.items:
            by_area[it.area or "UNKNOWN"].append(it)
        return {
            area: MTOStatistics(items).summary()
            for area, items in sorted(by_area.items())
        }

    # ------------------------------------------------------------------
    # 类型数量概览
    # ------------------------------------------------------------------

    def type_overview(self) -> Dict[str, Dict[str, float]]:
        """
        返回 {类型: {单位: 总数量}} 的概览字典，
        例如 {"PIPE": {"M": 235.5}, "VALVE": {"NO": 12}}
        """
        result: Dict[str, Dict[str, float]] = defaultdict(lambda: defaultdict(float))
        for it in self.items:
            result[it.item_type][it.unit] += it.qty
        return {k: dict(v) for k, v in sorted(result.items())}


# ---------------------------------------------------------------------------
# Excel 输出
# ---------------------------------------------------------------------------

# 颜色配置（浅色系，对打印友好）
_HEADER_FILL = "1F497D"   # 深蓝
_HEADER_FONT = "FFFFFF"   # 白色字
_SUBHEADER_FILL = "4472C4"
_TYPE_FILLS = {
    "PIPE":    "DDEEFF",
    "ELBOW":   "E2EFDA",
    "TEE":     "FFF2CC",
    "REDUCER": "FCE4D6",
    "FLANGE":  "E2EFDA",
    "VALVE":   "F4CCFF",
    "BOLT":    "F5F5F5",
    "GASKET":  "F5F5F5",
    "CAP":     "FFF2CC",
    "OLET":    "DDEEFF",
    "OTHER":   "FFFFFF",
}


def _make_border():
    thin = Side(style="thin", color="AAAAAA")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def export_excel(stats: MTOStatistics, output_path: str | Path,
                 title: str = "E3D 材料统计报表") -> None:
    """
    将统计结果输出到 Excel，包含三个工作表：
      1. 总量汇总
      2. 按管线汇总
      3. 按区域汇总
    """
    if not HAS_OPENPYXL:
        raise ImportError("导出 Excel 需要安装 openpyxl：pip install openpyxl")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # 删除默认空 Sheet

    _write_summary_sheet(wb, stats, title)
    _write_by_line_sheet(wb, stats)
    _write_by_area_sheet(wb, stats)
    _write_overview_sheet(wb, stats)

    wb.save(output_path)
    print(f"[OK] Excel 已保存：{output_path}")


def _col_headers():
    return ["序号", "元件类型", "元件代码", "描述", "管道等级", "管径", "材质", "数量", "单位", "涉及管线数"]


def _write_summary_sheet(wb, stats: MTOStatistics, title: str):
    ws = wb.create_sheet("总量汇总")
    rows = stats.summary()

    # 标题行
    ws.merge_cells("A1:J1")
    cell = ws["A1"]
    cell.value = title
    cell.font = Font(bold=True, size=14, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=_HEADER_FILL)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 列头
    headers = _col_headers()
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True, color=_HEADER_FONT)
        cell.fill = PatternFill("solid", fgColor=_HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _make_border()
    ws.row_dimensions[2].height = 20

    # 数据行
    for seq, row in enumerate(rows, 1):
        r = seq + 2
        fill_color = _TYPE_FILLS.get(row.item_type, "FFFFFF")
        vals = [seq, row.item_type, row.item_code, row.desc,
                row.spec, row.size, row.material,
                row.total_qty, row.unit, row.line_count]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.fill = PatternFill("solid", fgColor=fill_color)
            cell.border = _make_border()
            cell.alignment = Alignment(
                horizontal="center" if c in (1, 2, 8, 9, 10) else "left",
                vertical="center"
            )

    # 列宽
    col_widths = [6, 10, 16, 30, 10, 12, 14, 10, 8, 12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 冻结首两行
    ws.freeze_panes = "A3"


def _write_by_line_sheet(wb, stats: MTOStatistics):
    ws = wb.create_sheet("按管线汇总")
    by_line = stats.by_line()

    ws.merge_cells("A1:J1")
    cell = ws["A1"]
    cell.value = "各管线材料明细"
    cell.font = Font(bold=True, size=13, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=_HEADER_FILL)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    cur_row = 2
    for line_no, rows in by_line.items():
        # 管线标题
        ws.merge_cells(
            start_row=cur_row, start_column=1,
            end_row=cur_row, end_column=10
        )
        cell = ws.cell(row=cur_row, column=1, value=f"管线号：{line_no}  (共 {len(rows)} 类)")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=_SUBHEADER_FILL)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[cur_row].height = 18
        cur_row += 1

        # 列头
        for c, h in enumerate(_col_headers(), 1):
            cell = ws.cell(row=cur_row, column=c, value=h)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="BDD7EE")
            cell.border = _make_border()
            cell.alignment = Alignment(horizontal="center")
        cur_row += 1

        for seq, row in enumerate(rows, 1):
            vals = [seq, row.item_type, row.item_code, row.desc,
                    row.spec, row.size, row.material,
                    row.total_qty, row.unit, row.line_count]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=cur_row, column=c, value=v)
                cell.border = _make_border()
                cell.alignment = Alignment(
                    horizontal="center" if c in (1, 2, 8, 9, 10) else "left"
                )
            cur_row += 1
        cur_row += 1  # 空行分隔

    col_widths = [6, 10, 16, 30, 10, 12, 14, 10, 8, 12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def _write_by_area_sheet(wb, stats: MTOStatistics):
    ws = wb.create_sheet("按区域汇总")
    by_area = stats.by_area()

    ws.merge_cells("A1:J1")
    cell = ws["A1"]
    cell.value = "各区域材料明细"
    cell.font = Font(bold=True, size=13, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=_HEADER_FILL)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    cur_row = 2
    for area, rows in by_area.items():
        ws.merge_cells(
            start_row=cur_row, start_column=1,
            end_row=cur_row, end_column=10
        )
        cell = ws.cell(row=cur_row, column=1,
                       value=f"区域：{area}  (共 {len(rows)} 类)")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=_SUBHEADER_FILL)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[cur_row].height = 18
        cur_row += 1

        for c, h in enumerate(_col_headers(), 1):
            cell = ws.cell(row=cur_row, column=c, value=h)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="BDD7EE")
            cell.border = _make_border()
            cell.alignment = Alignment(horizontal="center")
        cur_row += 1

        for seq, row in enumerate(rows, 1):
            vals = [seq, row.item_type, row.item_code, row.desc,
                    row.spec, row.size, row.material,
                    row.total_qty, row.unit, row.line_count]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=cur_row, column=c, value=v)
                cell.border = _make_border()
                cell.alignment = Alignment(
                    horizontal="center" if c in (1, 2, 8, 9, 10) else "left"
                )
            cur_row += 1
        cur_row += 1

    col_widths = [6, 10, 16, 30, 10, 12, 14, 10, 8, 12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def _write_overview_sheet(wb, stats: MTOStatistics):
    ws = wb.create_sheet("类型概览")
    overview = stats.type_overview()

    ws.merge_cells("A1:D1")
    cell = ws["A1"]
    cell.value = "材料类型数量概览"
    cell.font = Font(bold=True, size=13, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=_HEADER_FILL)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    headers = ["元件类型", "单位", "总数量", "说明"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True, color=_HEADER_FONT)
        cell.fill = PatternFill("solid", fgColor=_HEADER_FILL)
        cell.border = _make_border()
        cell.alignment = Alignment(horizontal="center")

    _type_notes = {
        "PIPE": "管子长度合计",
        "ELBOW": "弯头数量",
        "TEE": "三通数量",
        "REDUCER": "大小头数量",
        "CAP": "管帽数量",
        "FLANGE": "法兰数量",
        "GASKET": "垫片数量",
        "BOLT": "螺栓套数",
        "VALVE": "阀门数量",
        "OLET": "支管台数量",
        "OTHER": "其他",
    }

    row = 3
    for itype, units in overview.items():
        fill_color = _TYPE_FILLS.get(itype, "FFFFFF")
        for unit, qty in units.items():
            vals = [itype, unit, round(qty, 3), _type_notes.get(itype, "")]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=c, value=v)
                cell.fill = PatternFill("solid", fgColor=fill_color)
                cell.border = _make_border()
                cell.alignment = Alignment(
                    horizontal="center" if c != 4 else "left"
                )
            row += 1

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 20


# ---------------------------------------------------------------------------
# 文本报表输出（无 Excel 时的备选）
# ---------------------------------------------------------------------------

def print_summary(stats: MTOStatistics) -> None:
    """在终端打印汇总报表。"""
    rows = stats.summary()
    overview = stats.type_overview()

    print("\n" + "=" * 80)
    print("  E3D 材料统计汇总")
    print("=" * 80)

    # 类型概览
    print("\n【材料类型概览】")
    print(f"  {'类型':<12}{'单位':<6}{'总数量':>12}")
    print("  " + "-" * 32)
    for itype, units in overview.items():
        for unit, qty in units.items():
            print(f"  {itype:<12}{unit:<6}{qty:>12.3f}")

    # 详细明细
    print("\n【详细明细】")
    header = (f"  {'序':<4}{'类型':<10}{'元件代码':<16}{'描述':<26}"
              f"{'等级':<8}{'管径':<10}{'数量':>8}{'单位':<6}")
    print(header)
    print("  " + "-" * 90)

    for seq, row in enumerate(rows, 1):
        line = (f"  {seq:<4}{row.item_type:<10}{row.item_code:<16}"
                f"{row.desc[:24]:<26}{row.spec:<8}{row.size:<10}"
                f"{row.total_qty:>8.3f}{row.unit:<6}")
        print(line)

    print("=" * 80)
    print(f"  合计 {len(rows)} 类材料，{len(stats.items)} 条原始记录")
    print("=" * 80 + "\n")


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------

def _normalize_col(raw: str) -> Optional[str]:
    """将原始列名归一化为内部字段名。"""
    key = raw.strip().upper().replace(" ", "_").replace("-", "_")
    if key in _COL_ALIASES:
        return _COL_ALIASES[key]
    # 模糊匹配
    for alias, norm in _COL_ALIASES.items():
        if alias in key or key in alias:
            return norm
    return None


def _normalize_size(raw: str) -> str:
    """规范化管径表示，例如 '2"' → '2"', 'DN50' → 'DN50', '50MM' → 'DN50'。"""
    s = raw.strip().upper()
    if not s:
        return ""
    # 已有 DN 前缀
    if s.startswith("DN"):
        return s
    # 纯数字 mm（非英寸）
    if re.fullmatch(r"\d+", s):
        val = int(s)
        if val >= 6:   # 认为是 mm
            return f"DN{val}"
        return s
    # NPS 英寸格式：1/2", 2", 3/4" 等
    if '"' in s or "INCH" in s:
        return s.replace("INCH", '"').strip()
    return s


def _detect_type(code: str, desc: str) -> str:
    """根据元件代码和描述判断元件类型。"""
    text = (code + " " + desc).upper()
    for keyword, itype in _ITEM_TYPE_MAP.items():
        if keyword.upper() in text:
            return itype
    return "OTHER"


def _parse_col_spans_from_sep(sep_line: str) -> List[tuple]:
    """从分隔线（--- --- ---）解析列位置区间。"""
    spans = []
    i = 0
    n = len(sep_line)
    while i < n:
        if sep_line[i] in "-=":
            start = i
            while i < n and sep_line[i] in "-=":
                i += 1
            spans.append((start, i))
        else:
            i += 1
    return spans


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------

def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="E3D MTO 材料统计工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  python e3d_mto.py input.csv
  python e3d_mto.py input.xlsx -o output.xlsx
  python e3d_mto.py input.txt --encoding gbk -o report.xlsx
        """
    )
    parser.add_argument("input", help="输入文件（CSV / Excel / 文本）")
    parser.add_argument("-o", "--output", default="", help="输出 Excel 文件路径")
    parser.add_argument("--encoding", default="utf-8-sig",
                        help="文件编码（默认 utf-8-sig，中文文件可用 gbk）")
    parser.add_argument("--title", default="E3D 材料统计报表",
                        help="报表标题")
    args = parser.parse_args()

    parser_obj = E3DMTOParser(encoding=args.encoding)
    items = parser_obj.parse_file(args.input)

    if parser_obj.warnings:
        for w in parser_obj.warnings:
            print(f"[WARN] {w}", file=sys.stderr)

    print(f"[INFO] 解析完成，共 {len(items)} 条有效记录")

    stats = MTOStatistics(items)
    print_summary(stats)

    if args.output:
        export_excel(stats, args.output, title=args.title)
    elif HAS_OPENPYXL:
        out = Path(args.input).with_suffix(".mto_report.xlsx")
        export_excel(stats, out, title=args.title)


if __name__ == "__main__":
    main()
