---
name: lsv-txt-processor
description: Process TXT LSV/IPV files from a user-selected data folder into an Excel LSV workbook using the bundled template, naming, directory mirroring, color, and summary rules.
---

# TXT LSV Processor

Use this skill when the user asks to process LSV/IPV TXT files from a selected data folder into an `*-LSV.xlsx` workbook.

## Scope

- Input: one selected folder containing exported TXT files.
- Recognize LSV data by filenames containing `LSV` or `IPV` case-insensitively.
- Do not process EIS, DRT, ECM, STA, or impedance files as LSV.
- Use the bundled project template in `support/templates`.
- Output to the configured processed-data folder, or `./processed` by default.
- Preserve the raw-data directory context when the input folder is under the exported raw-data root.
- Avoid redundant final folders: if the workbook name carries the final folder name, place it in the parent context directory.
- Keep source TXT filenames as worksheet names except for Excel's 31-character limit and uniqueness suffixes.

## Standard Processing Rules

- Raw current in `mA` is converted to `A` before applying the template formula pattern.
- Active area is `0.196 cm^2`.
- Summary sheet uses one 3-column block per dataset:
  - current density
  - voltage
  - power density
- Each group uses the same temperature color family, with different contrast levels for repeated datasets.
- Summary block fill color and worksheet tab color must match.
- Treat names like `1A800` and `A800` as different groups, even though they share the same temperature family.

## Run

Use the bundled script:

```powershell
py -3.11 ".\skills\lsv-txt-processor\scripts\process_lsv_folder.py" --input-dir "C:\path\to\selected\folder"
```

Optional:

```powershell
--output-root "C:\path\to\processed"
```

## Verification

After running:

- Confirm the output workbook exists.
- Confirm all LSV/IPV TXT files in the selected folder became worksheets.
- Confirm the summary block count equals the data worksheet count.
- Confirm summary fills match worksheet tab colors.
- Confirm Excel formulas were recalculated if Excel is available.

Never invent missing data. If no LSV/IPV TXT files are found, stop and report that clearly.
