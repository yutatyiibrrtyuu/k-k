"""
E3D 材料统计示例脚本
====================
生成示例 CSV 数据并运行统计，展示全部功能。
"""

from e3d_mto import (
    E3DMTOParser, MTOStatistics,
    print_summary, export_excel,
    HAS_OPENPYXL,
)

# ---------------------------------------------------------------------------
# 构造示例 CSV（模拟 E3D 导出格式）
# ---------------------------------------------------------------------------

SAMPLE_CSV = """LINE_NO,ITEM_CODE,DESCRIPTION,SPEC,SIZE,SIZE2,MATERIAL,QUANTITY,UNIT,AREA,ISOMETRIC
100-P-001,PIPE-001,ERW PIPE ASTM A106 GR.B,CS3,2",, A106B,12.500,M,AREA-A,ISO-001
100-P-001,ELBOW-001,90DEG ELBOW LR BW,CS3,2",,A234 WPB,4,NO,AREA-A,ISO-001
100-P-001,ELBOW-001,45DEG ELBOW LR BW,CS3,2",,A234 WPB,2,NO,AREA-A,ISO-001
100-P-001,TEE-001,EQUAL TEE BW,CS3,2",,A234 WPB,1,NO,AREA-A,ISO-001
100-P-001,FLANGE-001,WELD NECK FLANGE CL150,CS3,2",,A105,4,NO,AREA-A,ISO-001
100-P-001,GASKET-001,SPIRAL WOUND GASKET,CS3,2",,SS316+GRAPHITE,4,NO,AREA-A,ISO-001
100-P-001,BOLT-001,STUD BOLT WITH 2 HEX NUTS,CS3,2",,B7/2H,4,SET,AREA-A,ISO-001
100-P-001,VALVE-001,GATE VALVE CL150,CS3,2",,A216WCB,2,NO,AREA-A,ISO-001
200-P-002,PIPE-002,SEAMLESS PIPE ASTM A106 GR.B,CS3,4",,A106B,25.000,M,AREA-B,ISO-002
200-P-002,ELBOW-002,90DEG ELBOW LR BW,CS3,4",,A234 WPB,6,NO,AREA-B,ISO-002
200-P-002,REDUC-001,CONCENTRIC REDUCER BW,CS3,4",2",A234 WPB,2,NO,AREA-B,ISO-002
200-P-002,FLANGE-002,WELD NECK FLANGE CL300,CS3,4",,A105,6,NO,AREA-B,ISO-002
200-P-002,GASKET-002,RING JOINT GASKET,CS3,4",,SS316,6,NO,AREA-B,ISO-002
200-P-002,BOLT-002,STUD BOLT WITH 2 HEX NUTS,CS3,4",,B7/2H,6,SET,AREA-B,ISO-002
200-P-002,VALVE-002,BALL VALVE CL300,CS3,4",,A216WCB,3,NO,AREA-B,ISO-002
200-P-002,OLET-001,WELDOLET,CS3,4",2",A105,2,NO,AREA-B,ISO-002
300-P-003,PIPE-003,ERW PIPE ASTM A106 GR.B,CS3,2",,A106B,8.200,M,AREA-A,ISO-003
300-P-003,ELBOW-003,90DEG ELBOW LR BW,CS3,2",,A234 WPB,3,NO,AREA-A,ISO-003
300-P-003,CAP-001,CAP BW,CS3,2",,A234 WPB,1,NO,AREA-A,ISO-003
300-P-003,FLANGE-003,BLIND FLANGE CL150,CS3,2",,A105,2,NO,AREA-A,ISO-003
300-P-003,GASKET-003,SPIRAL WOUND GASKET,CS3,2",,SS316+GRAPHITE,2,NO,AREA-A,ISO-003
300-P-003,BOLT-003,STUD BOLT WITH 2 HEX NUTS,CS3,2",,B7/2H,2,SET,AREA-A,ISO-003
300-P-003,VALVE-003,CHECK VALVE CL150,CS3,2",,A216WCB,1,NO,AREA-A,ISO-003
400-P-004,PIPE-004,SEAMLESS PIPE A312 TP316L,SS3,3",,SS316L,15.600,M,AREA-C,ISO-004
400-P-004,ELBOW-004,90DEG ELBOW LR BW,SS3,3",,A403 WP316L,4,NO,AREA-C,ISO-004
400-P-004,TEE-004,EQUAL TEE BW,SS3,3",,A403 WP316L,2,NO,AREA-C,ISO-004
400-P-004,FLANGE-004,WELD NECK FLANGE CL150,SS3,3",,A182 F316L,4,NO,AREA-C,ISO-004
400-P-004,GASKET-004,SPIRAL WOUND GASKET,SS3,3",,SS316+GRAPHITE,4,NO,AREA-C,ISO-004
400-P-004,BOLT-004,STUD BOLT WITH 2 HEX NUTS,SS3,3",,B8M/8M,4,SET,AREA-C,ISO-004
400-P-004,VALVE-004,BUTTERFLY VALVE CL150,SS3,3",,A351 CF8M,2,NO,AREA-C,ISO-004
"""

# 固定宽度文本必须严格按分隔线位置对齐
# 列宽(sep段长): LINE_NO=9, ITEM_CODE=10, SPEC=5, SIZE=6, QTY=7, UNIT=5, DESCRIPTION=rest
SAMPLE_FIXED_TEXT = (
    "MATERIAL TAKE OFF REPORT\n"
    "PROJECT: DEMO PLANT            DATE: 2026-03-28\n"
    "\n"
    "LINE_NO   ITEM_CODE  SPEC  SIZE   QTY     UNIT  DESCRIPTION\n"
    "--------- ---------- ----- ------ ------- ----- ---------------------------\n"
    "500-P-005 PIPE-005   CS3   6\"     32.000  M     ERW PIPE ASTM A106 GR.B\n"
    "500-P-005 ELBOW-005  CS3   6\"     8       NO    90DEG ELBOW LR BW\n"
    "500-P-005 FLANGE-005 CS3   6\"     8       NO    WELD NECK FLANGE CL150\n"
    "500-P-005 VALVE-005  CS3   6\"     2       NO    GATE VALVE CL150\n"
    "500-P-005 BOLT-005   CS3   6\"     8       SET   STUD BOLT WITH 2 HEX NUTS\n"
)


def demo_csv():
    print("=" * 60)
    print("  示例 1：解析 CSV 格式 MTO")
    print("=" * 60)
    parser = E3DMTOParser()
    items = parser.parse_text(SAMPLE_CSV, fmt="csv")
    print(f"解析到 {len(items)} 条记录")

    stats = MTOStatistics(items)
    print_summary(stats)

    if HAS_OPENPYXL:
        export_excel(stats, "e3d_mto_output.xlsx", title="DEMO PLANT 材料统计报表")
    else:
        print("[提示] 安装 openpyxl 后可导出 Excel：pip install openpyxl")

    return stats


def demo_fixed_text():
    print("=" * 60)
    print("  示例 2：解析固定宽度文本格式 MTO")
    print("=" * 60)
    parser = E3DMTOParser()
    items = parser.parse_text(SAMPLE_FIXED_TEXT, fmt="text")
    print(f"解析到 {len(items)} 条记录")

    stats = MTOStatistics(items)
    print_summary(stats)
    return stats


def demo_api():
    """展示如何在代码中使用 API。"""
    print("=" * 60)
    print("  示例 3：API 用法演示")
    print("=" * 60)

    from e3d_mto import MTOItem
    items = [
        MTOItem(line_no="600-P-006", item_code="PIPE-006",
                desc="ERW PIPE A106B", spec="CS3", size='8"',
                material="A106B", qty=45.0, unit="M",
                area="AREA-D", item_type="PIPE"),
        MTOItem(line_no="600-P-006", item_code="ELBOW-006",
                desc="90DEG ELBOW LR BW", spec="CS3", size='8"',
                material="A234 WPB", qty=10.0, unit="NO",
                area="AREA-D", item_type="ELBOW"),
        MTOItem(line_no="600-P-006", item_code="VALVE-006",
                desc="GATE VALVE CL150", spec="CS3", size='8"',
                material="A216WCB", qty=4.0, unit="NO",
                area="AREA-D", item_type="VALVE"),
    ]

    stats = MTOStatistics(items)

    print("类型概览：")
    for itype, units in stats.type_overview().items():
        for unit, qty in units.items():
            print(f"  {itype:<12} {qty:.3f} {unit}")

    print("\n汇总明细（总量）：")
    for row in stats.summary():
        print(f"  {row.item_type:<10} {row.size:<8} {row.total_qty:>8.3f} {row.unit}")


if __name__ == "__main__":
    demo_csv()
    demo_fixed_text()
    demo_api()
