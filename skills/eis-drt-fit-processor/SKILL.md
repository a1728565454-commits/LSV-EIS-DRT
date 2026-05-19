---
name: eis-drt-fit-processor
description: Process TXT EIS or impedance files from a user-selected data folder into an Excel EIS+DRT+fit workbook using the bundled template, EISART DRT fitting workflow, ECM/DRT column extraction, formatting, and validation rules.
---

# TXT EIS + DRT + Fit Processor

Use this skill when the user asks to process EIS TXT files or impedance TXT files into an `*EIS+DRT+fit.xlsx` workbook.

## Scope

- Input: one selected folder containing exported TXT impedance data.
- Accept TXT files when:
  - filename contains `EIS`, or
  - the selected folder path indicates a generic impedance context.
- Exclude `STA` unless the user explicitly asks otherwise.
- Use the bundled project template in `support/templates`.
- Use the bundled EISART-based DRT fitting workflow.
- Keep only the final workbook. Intermediate DRT/ECM folders are temporary and should be removed.

## Output Rules

- Output to the configured processed-data folder, or `./processed` by default.
- Preserve raw-data directory context when the selected folder is under the exported raw-data root.
- Avoid redundant final folders: if the workbook name carries the final folder name, place it in the parent context directory.
- Workbook name: replace the template's folder-name placeholder with the selected folder name.
- Keep source TXT filenames as worksheet names except for Excel's 31-character limit and uniqueness suffixes.

## Workbook Structure

The first three sheets must be:

1. `R summary`
2. `Rp summary`
3. `DRT summary`

Then append one worksheet per raw EIS TXT file.

## Data Rules

- Active area is `0.196 cm^2`.
- EIS raw data sheets keep three columns: frequency, Zre, Zim.
- `R summary` uses 4 columns per dataset:
  - measured `Zre * 0.196`
  - measured `Zim * -0.196`
  - ECM exported column 2
  - ECM exported column 3 multiplied by `-1`
- `Rp summary` compares measured Rp and fitted Rp; do not keep literal fitted-Zim text.
- `DRT summary` uses 2 columns per dataset:
  - DRT exported column 1
  - DRT exported column 2
- Do not add a text-description row under `DRT summary`; data starts directly below the dataset header row.

## Formatting Standard

- Same temperature point uses the same color family, with different contrast for repeated datasets.
- Summary sheet block color matches the worksheet tab color.
- Fill the entire used block area, including empty cells inside the block.
- Do not fill outside the used block area.
- Add a vertical divider after every two columns.
- Use black font throughout the summary sheets.

## Run

Use the bundled script:

```powershell
py -3.11 ".\skills\eis-drt-fit-processor\scripts\process_eis_drt_fit_folder.py" --input-dir "C:\path\to\selected\folder"
```

Optional:

```powershell
--output-root "C:\path\to\processed"
```

## Verification

After running:

- Confirm the workbook exists.
- Confirm the first three sheets are the three summary sheets.
- Confirm the number of data worksheets equals the accepted TXT file count.
- Confirm formula errors are zero after Excel recalculation.
- Confirm no summary cell contains `Zim` text.
- Confirm block-internal fill is complete and block-external fill is absent.

Never invent DRT, ECM, or fit results. If the DRT fitting tool fails for a file, report the failed file instead of fabricating output.
