# CLAUDE.md

This file provides guidance for AI assistants working with this repository.

## Project Overview

This is a Python-based engineering tools project containing two independent utilities:

1. **E3D Material Takeoff (MTO) Statistics Tool** (`e3d_mto.py`) — Parses and aggregates piping material reports exported from AVEVA E3D design software.
2. **Heat Exchanger Design Calculator** (`heat_exchanger.py`) — Implements thermal design calculations using both LMTD and NTU-ε methods.

Both modules work as standalone CLI tools and as importable libraries. The project targets Python 3.10+ and has no mandatory external dependencies (optional: `openpyxl`, `pandas`).

## Repository Structure

```
k-k/
├── e3d_mto.py          # Main MTO tool: parser, aggregation, Excel/console output
├── e3d_mto_demo.py     # Usage examples for e3d_mto.py
├── heat_exchanger.py   # Heat exchanger design calculator
└── .gitignore          # Excludes Python cache, build artifacts, .xlsx files
```

## Running the Tools

```bash
# E3D MTO tool (CLI)
python e3d_mto.py <input_file> [-o output.xlsx] [--encoding utf-8-sig] [--title "Report Title"]

# E3D MTO demos
python e3d_mto_demo.py

# Heat exchanger examples (runs 4 worked examples)
python heat_exchanger.py
```

No install step is required. Optional dependencies:
- `pip install openpyxl` — required for Excel output in e3d_mto.py
- `pip install pandas` — required for reading `.xlsx`/`.xls` input files

## e3d_mto.py Architecture

### Classes
| Class | Role |
|---|---|
| `MTOItem` | Dataclass for a single MTO record |
| `E3DMTOParser` | Parses CSV, Excel, and fixed-width text formats |
| `SummaryRow` | Dataclass for one aggregated summary entry |
| `MTOStatistics` | Groups `MTOItem` objects and generates stats |

### Data Flow
```
Input file (CSV / Excel / fixed-width text)
    → E3DMTOParser.parse()
    → List[MTOItem]
    → MTOStatistics.add_item()
    → export_excel() / print_summary()
```

### Supported Component Types
Normalized via `_ITEM_TYPE_MAP`: `PIPE`, `ELBOW`, `TEE`, `REDUCER`, `CAP`, `FLANGE`, `GASKET`, `BOLT`, `VALVE`, `OLET`, `OTHER`.

### Column Name Mapping
The parser accepts both English and Chinese column headers. Key mappings:
- `LINE_NO` / `管线号`
- `ITEM_CODE` / `ITEM` / `元件代码`
- `SPECIFICATION` / `SPEC` / `管道等级`
- `SIZE` / `口径`
- `MATERIAL` / `材质`
- `QUANTITY` / `QTY` / `数量`
- `UNIT` / `单位`
- `DESCRIPTION` / `DESC` / `描述`

## heat_exchanger.py Architecture

### Key Types
| Symbol | Meaning |
|---|---|
| `HXType` | Enum: `PARALLEL`, `COUNTER`, `CROSS`, `SHELL_TUBE` |
| `FluidStream` | Dataclass: `mass_flow`, `cp`, `T_in`, `T_out` (K) |

### Core Functions
| Function | Purpose |
|---|---|
| `lmtd(T_h_in, T_h_out, T_c_in, T_c_out, hx_type)` | Returns corrected LMTD [K] |
| `effectiveness(NTU, C_ratio, hx_type)` | Returns ε from NTU and capacity ratio |
| `design_lmtd(hot, cold, U, hx_type)` | Design: find area from temperatures |
| `design_ntu(hot, cold, U, A, hx_type)` | Design: find outlet temps from area |
| `rate_ntu(hot, cold, U, A, hx_type)` | Rating: verify existing design |
| `overall_U(h_hot, h_cold, R_hot, R_cold, t_wall, k_wall)` | Calculate overall U [W/m²K] |

All temperatures are in **Kelvin** internally; results display in both K and °C.

## Code Conventions

### Style
- **Python 3.10+** type hints throughout (use `tuple[...]`, `list[...]`, not `Tuple`/`List` from `typing`)
- Dataclasses for plain data containers
- 4-space indentation, PEP 8 naming
- Module-level constants in `UPPER_CASE` (private ones prefixed `_`)
- Private helpers prefixed with `_` (e.g., `_normalize_col`, `_detect_type`)

### Error Handling
- Collect non-fatal warnings in `self.warnings: list[str]` (E3DMTOParser pattern)
- Raise `ValueError` with descriptive messages for invalid inputs
- Use try-except only at module boundaries (optional imports)
- Clamp values to safe ranges rather than crashing (e.g., F-factor in LMTD)

### Documentation
- Module-level docstring describes purpose, usage, supported formats
- Function docstrings include `Parameters` and `Returns` sections
- Section separators: `# --------` lines between logical blocks

### Bilingual Support
The codebase intentionally supports both Chinese and English column names for the E3D tool. When adding new field mappings, add aliases for both languages in `_COL_ALIASES`.

## Development Workflow

### Making Changes
1. Edit files directly — no build step needed.
2. Test locally: `python e3d_mto_demo.py` and `python heat_exchanger.py`.
3. For Excel output testing, `openpyxl` must be installed.

### Adding a New Component Type to e3d_mto.py
1. Add an entry to `_ITEM_TYPE_MAP` dict (maps normalized name fragments → type string).
2. Update `SummaryRow` if the new type needs its own column in Excel output.
3. Add the type to the component-type order list in `MTOStatistics._build_summary()` if ordering matters.

### Adding a New Heat Exchanger Type
1. Add a variant to the `HXType` enum.
2. Implement the ε-NTU relationship in the `effectiveness()` function.
3. Implement the LMTD correction factor F in the `lmtd()` function.
4. Add a worked example to the `if __name__ == "__main__"` block.

### Commit Style
Follow the existing commit message pattern: imperative mood, short subject line describing what was added or changed.

```
Add support for butterfly valves in MTO parser
Fix F-factor clamping for cross-flow heat exchangers
```

## Key Constraints & Gotchas

- `.xlsx` files are excluded from git (see `.gitignore`). Do not commit test data files.
- The MTO parser's fixed-width format relies on AVEVA E3D's specific column widths — do not generalize this parser without testing against real E3D output.
- All heat exchanger temperatures must be in **Kelvin**; the `FluidStream` dataclass does not enforce this, so incorrect units will produce silently wrong results.
- `E3DMTOParser.warnings` should be checked after parsing — it surfaces non-fatal issues like unrecognized column names or skipped rows.
- `pandas` is only imported inside `_parse_excel()`, so the rest of the module works without it installed.
