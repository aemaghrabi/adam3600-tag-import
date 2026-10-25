---
name: adam3600-tag-import
description: Build an Advantech ADAM-3600 RTU tag import .xls from a tag list or any Modbus register map, using the VA_Temp template convention. Asks base-0/base-1 on every run.
license: MIT
---

# ADAM-3600 tag import builder

Turns a tag list into the `.xls` the ADAM-3600 configuration software imports. The layout follows `assets/adam-3600_tags_template.xls`: sheet `BoardIO`, then one sheet per I/O device, then `User Tag` and `Calc Tag`. Every rule below was reverse-engineered from that template's `VA_Temp` sheet and confirmed by the template's author. The template and this file are the only authority. Do not add codes, defaults or meanings from general knowledge.

## Files in this skill
- `scripts/adam3600_tags.py`: the generator. Subcommands: `selftest`, `template`, `build`, `decode`. In Claude Code, the path is `${CLAUDE_SKILL_DIR}/scripts/adam3600_tags.py`; elsewhere, use the path of this skill's folder.
- `assets/adam-3600_tags_template.xls`: the original template, which is the source of truth. `decode` it to inspect.
- `assets/input_template.xlsx`: the blank canonical input (a `Tags` sheet with dropdowns, plus a `Reference` sheet). Send it as is when someone asks for the input form.

## Non-negotiable rules

1. **Ask the address base on every run** with AskUserQuestion (or plainly, if that tool is missing), before computing any address. Never infer it from the data, the file name or an earlier run. Explain both options:
   - **base-0**: input address 0 = Modbus offset 0, so the address is written +1 (HR 0 → `40001`).
   - **base-1**: input address 1 = Modbus offset 0, so the address is written as is (HR 1 → `40001`).
2. **No assumptions.** A missing required value means stop and ask. Never fix data yourself: no truncating names, no guessing formats or byte orders, no defaults beyond the ones listed under *Blank optional fields*.
3. **ERROR lines** (exit 1): show them to the user and ask for corrections. Apply only the values the user gives, then rerun.
4. **CONFIRM lines** (exit 2): show them verbatim and get explicit approval. Rerun with `--confirm` only after the user approves. If the user does not approve, ask what to change.
5. **NOTE lines** (registers shared by several tags) never block the build. Report them briefly, because the template itself reuses 40001 on purpose.
6. **Scaling is not supported yet.** `scalingType` is always `0`, and every `Scale.*` column gets the template value. If the source has scaling or engineering-range columns, say they will be ignored and ask whether to continue.
7. **User Tag and Calc Tag** sheets get their header row only. **BoardIO** is the fixed built-in list (AI.0–7, DI.0–7, DO.0–3). It changes only if the user supplies an export to copy it from (`--boardio`).
8. When the source is a device datasheet or register map, bind to what it says. Where it is silent (device name, scan rate, access, word order), ask.

## Workflow

### 0. Set up (every run)
```bash
pip install xlrd xlwt openpyxl --break-system-packages -q   # or: pip install -r requirements.txt
python3 -I <skill dir>/scripts/adam3600_tags.py selftest
```
The self-test must print `selftest OK`. It rebuilds the template's 16 VA_Temp tags cell by cell. If it fails, stop and tell the user. Do not generate anything.

### 1. Identify the request
- **Blank input template wanted**: send `assets/input_template.xlsx`, or regenerate it with `template OUT.xlsx`.
- **Mode A, canonical input**: the sheet has the headers `Device, Tag, RegisterType, Address, Format, Length, Access, StartBit, ActiveLow, ScanRate, Description`. Header case and spaces don't matter, and `Length`, `StartBit`, `ActiveLow` and `Description` may be absent. Go to step 3.
- **Mode B, any other tag list or register map** (vendor datasheet table, CSV, a table copied from a PDF, an old point list): go to step 2.
- **Editing or extending an existing ADAM-3600 export**: run `decode EXPORT.xls -o canonical.xlsx`. Addresses come out base-1. Relay every NOTE line it prints, for example scaling values that cannot be carried over. Then treat the result as Mode A. Recommend base-1 when asking the base, but still ask.

### 2. Mode B: map the source to the canonical input
a. Read the whole source. Find the header row, the tag rows, and the rows that are not tags (section titles, "reserved", blanks).
b. Present two tables for the user to confirm:
   - **Column mapping**: source column → canonical column.
   - **Value mapping**: every distinct source value in the register-type, data-type, access and byte-order columns → canonical value.

   Also state the row count, the rows to be skipped (and why), and the device name(s) to use.
c. Ask for anything the source lacks: the device name (one per device), the scan rate (one value, or per group), access, and word or byte order. If source names would make `Device:Tag` longer than 26 characters, or contain spaces or symbols, list them and ask how to rename. You may propose names, but use them only when approved.
d. Wait for confirmation. Then write the canonical input: copy `assets/input_template.xlsx` and fill its `Tags` sheet with openpyxl, or write a CSV with the same headers. Keep it, because it is delivered with the result.

Mapping guidance. Every mapping still needs the user's confirmation.
- **Register type**: Coil / 0x / FC01 → `CS`; Discrete input / Input status / 1x / FC02 → `IS`; Input register / 3x / FC04 → `IR`; Holding register / 4x / FC03 → `HR`.
- **Addresses**: split a prefixed address into area + number (`40001` → HR 1, `400001` → HR 1, `30010` → IR 10). Convert hex to decimal. You may tell the user which base the source appears to use, but the base question in step 3 is still asked.
- **Data types**: 16-bit unsigned → `UINT` 16; 16-bit signed → `INT` 16; 32-bit unsigned or signed → `UINT` / `INT` with Length 32 (the same conversion code, only Length changes); 32-bit float → `FLOAT` 32 or `RFLOAT` 32 depending on word order; 64-bit double → `DOUBLE` / `DOUBLE_BE` / `DOUBLE_LE`; BCD → `BCD` with its bit length; bit or boolean → `DISCRETE`, plus `StartBit` for a bit inside an HR/IR register.
- **Byte and word order**: the template defines only these. `FLOAT` = IEEE 754; `RFLOAT` = reversed word order; `INT_SWAP` / `UINT_SWAP` = bytes swapped; `HIWORD_INT` / `HIWORD_UINT` = high-word integer; `DOUBLE` = plain double; `DOUBLE_BE` = big-endian (IEEE) double; `DOUBLE_LE` = little-endian double. Vendor terms (ABCD, CDAB, BADC, DCBA, "word swap", "Modicon order") do not map one-to-one onto these. Always show the proposed match and ask. If nothing fits (for example a byte-swapped float), stop and ask which format to use.
- **Access**: R / RO / Read → `R`; R/W / RW → `RW`; write-only → ask, because it is not supported.
- **Bit fields and status words**: ask whether to create one `DISCRETE` tag per bit.

### 3. Ask the base
Use the two options from rule 1.

### 4. Build
```bash
python3 -I <skill dir>/scripts/adam3600_tags.py build INPUT --base N -o OUT.xls [--sheet NAME] [--boardio EXPORT.xls] [--confirm]
```
- **exit 1**: errors, nothing written. Show the ERROR lines (summarise by kind when there are many, and list the row numbers) and ask for fixes.
- **exit 2**: confirmations needed, nothing written. Show the CONFIRM lines and ask for approval. If approved, rerun with `--confirm`.
- **exit 0**: written. The script has already re-read the file and checked every cell's value and type.

### 5. Deliver
Name the output `<input stem>_ADAM3600_tags.xls` and hand it to the user. Add one line of context: the devices with their tag counts, the base used, and any NOTE lines. In Mode B, also hand over the canonical input `.xlsx` so the user can reuse or edit it.

## The convention (reverse-engineered from VA_Temp)

### Generated columns
| Column | Cell type | Value |
|---|---|---|
| name | text | `Device:Tag`. The left side of `:` is the device (= sheet name), the right side is the tag. At most **26** characters in total |
| type | number | 0 |
| initialValue | text | `0` discrete, `0.0` analog |
| dataType | number | 1 discrete, 0 analog |
| address | text | register-type digit + 4-digit base-1 number, from `0001` to `9999`: `0`=CS (FC01), `1`=IS (FC02), `3`=IR (FC04), `4`=HR (FC03) |
| Scale.spanLow / Scale.spanHigh | number | 0 / 0 discrete, 0 / 100 analog |
| signalReverse | text | `true` = active low, otherwise `false` |
| description | text | copied as given (empty stays an empty text cell) |
| scalingType | text | `0` |
| Scale.scale / bias / outputLow / outputHigh | number | 0 |
| Scale.clampToLow / clampToHigh / clampToZero | text | `false` |
| readWrite | text | `1` read-only, `3` read/write. IS and IR are always `1` |
| template | number | register area: 0 CS, 1 IS, 2 IR, 3 HR |
| length | text | bits ADAM-3600 reads from Modbus: 1 discrete; 8 / 16 / 32 / 64 analog |
| conversionType | text | from Format (table below). It does not depend on length |
| startBit | text | bit index 0–15 for a discrete bit on HR/IR (required); 0 on CS/IS; analog 0 unless given |
| scanRate | text | whole seconds, at least 1 |

The output is `.xls` (BIFF8) in Arial 10 with the template's column widths. Sheet order is BoardIO, then the devices in first-seen order, then User Tag and Calc Tag. One file can carry several devices.

### Formats
| Format | conversionType | Meaning | Usual length |
|---|---|---|---|
| DISCRETE | 21 | discrete bit (coil, input status, or one bit of a register) | 1 |
| UINT | 21 | unsigned integer | 8/16/32/64 |
| INT | 22 | signed integer | 8/16/32/64 |
| UINT_SWAP | 1 | unsigned integer, bytes swapped | 8/16/32/64 |
| INT_SWAP | 2 | signed integer, bytes swapped | 8/16/32/64 |
| HIWORD_INT | 16 | high-word signed integer | 8/16/32/64 |
| HIWORD_UINT | 17 | high-word unsigned integer | 8/16/32/64 |
| FLOAT | 25 | IEEE 754 floating point | 32 |
| RFLOAT | 26 | IEEE 754 floating point, reversed word order | 32 |
| BCD | 35 | BCD | 8/16/32 |
| DOUBLE | 61 | plain double | 64 |
| DOUBLE_BE | 150 | big-endian (IEEE) double | 64 |
| DOUBLE_LE | 151 | little-endian double | 64 |

A width suffix can stand in for the Length column: `INT32` = INT + 32 bits, `UINT16`, `BCD16`, `FLOAT32`, `INT16_SWAP`. If the suffix and the Length column disagree, that is an error.

### Blank optional fields (agreed defaults)
- StartBit on an analog tag or on CS/IS → `0`.
- ActiveLow → `FALSE`.
- Description → empty.
- Length on a DISCRETE tag → `1`.

Everything else is required: Device, Tag, RegisterType, Address, Format, Length (analog), Access, ScanRate, and StartBit for a discrete bit on HR/IR.

### Checks the script runs
- **ERROR**:
  - The name is longer than 26 characters, or contains `:` in the device or tag part.
  - Duplicate name. Device names that differ only by case.
  - Reserved device name (BoardIO / User Tag / Calc Tag), or a character that is illegal in sheet names.
  - The address is not a whole number, or the result is outside 1–9999 (with a hint when a prefix such as 40001 was left in).
  - Unknown register type, format or access. Write-only access. RW on IS/IR.
  - A non-DISCRETE format on CS/IS.
  - Missing analog length, a length other than 8/16/32/64, or a suffix/length conflict.
  - A discrete bit on HR/IR without a StartBit, a StartBit outside 0–15, or a non-zero StartBit on CS/IS.
  - Missing, fractional or zero scan rate.
- **CONFIRM** (needs user approval):
  - FLOAT/RFLOAT with a length other than 32, DOUBLE* other than 64, BCD with 64, DISCRETE other than 1.
  - Characters outside `A–Z a–z 0–9 _ .` in a name. Names that differ only by case.
  - An analog tag with a non-zero StartBit. ActiveLow on an analog tag.
  - A multi-register tag that runs past register 9999.
- **NOTE**: tags sharing registers, grouped per cluster. Bits at different positions in the same register do not count as overlap. Each note says when it includes identical definitions.

### BoardIO (fixed, built in)
20 rows, name `BoardIO:<address>`, address `AI.0`–`AI.7` / `DI.0`–`DI.7` / `DO.0`–`DO.3`. All have type 0, initialValue `0`, an empty description, template 0, conversionType `0`, startBit `0` and scanRate `1`.
- AI: dataType 0, span 0/0, readWrite `1`, length `16`.
- DI: dataType 1, spanHigh 3.4028234663852886e+38, readWrite `1`, length `1`.
- DO: as DI, but readWrite `3`.

If the user's BoardIO ever changes, pass `--boardio <their export.xls>` to copy it exactly, and propose updating `boardio_builtin()` in the script.

### Not covered yet: ask the user for a sample export before supporting
- Scaling (`scalingType` other than 0) and non-template span values.
- User Tag and Calc Tag contents.
- Write-only access.
- Any conversionType outside the table above.
