# Codex LSV/EIS Processing Skills

This repository contains two local Codex skills for generic LSV and EIS data processing:

- `lsv-txt-processor`: process TXT LSV/IPV files into an Excel LSV workbook.
- `eis-drt-fit-processor`: process TXT EIS/impedance files into an Excel EIS + DRT + fit workbook.

## Included Support Files

The repository also includes the support files needed by the skills:

- `support/analysis_scripts`: workbook-building scripts used by the two skill wrappers.
- `support/templates`: Excel templates for LSV and EIS + DRT + fit workbooks.
- `support/eisart_drt_tool`: EISART/DRT fitting code, settings, and lookup arrays.

The original bundled `virtualenv` is not committed. Recreate the Python environment with:

```powershell
py -3.11 -m pip install -r .\requirements.txt
```

## Install

Copy the skill folders into your Codex skills directory:

```powershell
Copy-Item -Recurse .\skills\lsv-txt-processor "<codex-skills-dir>\"
Copy-Item -Recurse .\skills\eis-drt-fit-processor "<codex-skills-dir>\"
```

Restart Codex or reload skills after copying.

If you run the skills from this cloned repository, the wrappers prefer the included files under `support/`. If you install only the skill folders into another Codex profile, also keep this repository available or set these optional environment variables:

```powershell
$env:LSV_TEMPLATE = "C:\path\to\template-LSV.xlsx"
$env:EIS_DRT_TEMPLATE = "C:\path\to\template-EIS+DRT+fit.xlsx"
$env:DRT_RUNNER = "C:\path\to\run_eisart_batch_summary_final.py"
$env:LSV_EIS_DATA_ROOT = "C:\path\to\exported_txt"
$env:LSV_EIS_OUTPUT = "C:\path\to\processed"
```

## Notes

Set the environment variables above when using a different raw-data or output layout.
