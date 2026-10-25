# adam3600-tag-import

A Claude skill, and a standalone Python tool, that turns a tag list or any Modbus register map into a tag import file for the **Advantech ADAM-3600 RTU**. The output is an `.xls` workbook with the same layout as the configuration software's own export: `BoardIO`, one sheet per I/O device, `User Tag` and `Calc Tag`.

The conventions were reverse-engineered from a real template exported by the configuration software (`adam3600-tag-import/assets/adam-3600_tags_template.xls`). The generator reproduces that template's 16 sample tags cell for cell, including whether each cell is stored as text or as a number.

## What it does
- **Two kinds of input**:
  - The canonical input sheet (`assets/input_template.xlsx`), with one row per tag.
  - Any register map. Claude proposes a column and value mapping, and you confirm it before anything is generated.
- **Address base** (base-0 or base-1) is asked on every run, then converted to the ADAM-3600 address form: register-type digit (`0` CS / `1` IS / `3` IR / `4` HR) + 4-digit base-1 number.
- **Format and bit length are separate fields.** For example, INT32 is conversion type `22` with length `32`, and BCD16 is `35` with length `16`.
- **Validation before writing**:
  - Errors stop the build: names over 26 characters, RW on IS/IR, addresses beyond 9999, missing scan rate / access / format, fractional scan rates, and more.
  - Unusual but possible cases need confirmation (FLOAT with 16 bits, spaces in names, and so on).
  - Overlapping registers are reported but don't block the build.
- **Several devices** fit in one file.
- **`decode`** turns an existing ADAM-3600 export back into the canonical input, so you can edit it and regenerate.

## Repository layout
```
adam3600-tag-import/          <- the skill folder (install this)
  SKILL.md                    instructions Claude follows, plus the full convention reference
  scripts/adam3600_tags.py    generator: selftest | template | build | decode
  assets/adam-3600_tags_template.xls   original template (source of truth)
  assets/input_template.xlsx           blank canonical input with dropdowns
tests/test_adam3600_tags.py   regression tests (round trip against the template, rules)
requirements.txt
```

## Install as a Claude skill
- **Claude Code**: copy the `adam3600-tag-import` folder to `~/.claude/skills/` (all your projects) or to `.claude/skills/` inside a project.
- **Claude apps**: zip the `adam3600-tag-import` folder and upload it as a custom skill (Customize > Skills). Code execution must be enabled. See [Use Skills in Claude](https://support.claude.com/en/articles/12512180-use-skills-in-claude).

Then ask something like *"Generate the ADAM-3600 import file from this register map"* and attach the sheet.

## Use without Claude
```bash
pip install -r requirements.txt
S=adam3600-tag-import/scripts/adam3600_tags.py

python3 $S selftest                                   # must print "selftest OK"
python3 $S template my_tags.xlsx                      # blank input sheet
python3 $S build my_tags.xlsx --base 0 -o RTU1_ADAM3600_tags.xls
python3 $S decode existing_export.xls -o my_tags.xlsx # export -> editable input (base-1)
```
`build` exit codes:
- `0`: written and verified.
- `1`: errors, nothing written.
- `2`: confirmations needed, nothing written. Review the items, then rerun with `--confirm`.

## Input columns
| Column | Required | Notes |
|---|---|---|
| Device | yes | sheet name and the left part of `Device:Tag`; full name at most 26 characters |
| Tag | yes | right part of `Device:Tag` |
| RegisterType | yes | `CS`, `IS`, `IR`, `HR` (FC01/02/04/03 also accepted) |
| Address | yes | number inside the area, without the prefix; the base is chosen at build time |
| Format | yes | `DISCRETE UINT INT UINT_SWAP INT_SWAP HIWORD_INT HIWORD_UINT FLOAT RFLOAT BCD DOUBLE DOUBLE_BE DOUBLE_LE`, optionally with a width suffix (`INT32`) |
| Length | analog | 8 / 16 / 32 / 64 bits; discrete is 1 |
| Access | yes | `R` or `RW` |
| StartBit | discrete on HR/IR | bit 0–15; blank = 0 elsewhere |
| ActiveLow | no | `TRUE` = signal reversed |
| ScanRate | yes | whole seconds |
| Description | no | free text |

The complete column-by-column convention and conversion-type table are in [`adam3600-tag-import/SKILL.md`](adam3600-tag-import/SKILL.md).

## Tests
```bash
pip install -r requirements.txt
python -m unittest discover tests -v
```

## Limitations
- Scaling (`scalingType` other than 0), User Tag and Calc Tag contents, write-only access, and conversion types beyond the 13 formats listed are not supported yet.
- Files are written with `xlwt`, not by the ADAM-3600 software. Import one generated file to confirm your software version accepts it.

## Disclaimer
Advantech and ADAM-3600 are trademarks of Advantech Co., Ltd. This project is independent and not affiliated with or endorsed by Advantech.

## License
[MIT](LICENSE)
