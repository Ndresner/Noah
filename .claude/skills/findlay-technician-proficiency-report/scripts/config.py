"""
Shared run configuration for the Technician Proficiency workbook rebuild.

Each new period, create a config.json (see references/config.example.json in
the skill folder) with these keys and place it in the working directory
alongside this scripts/ folder:

{
  "roster_path": "path/to/this-period's ADP Tech Employee List export (.xlsx)",
  "qlik_data_path": "tech_data.json",           // written by the Qlik-pull step, see SKILL.md
  "template_path": "path/to/PRIOR period's finished workbook (.xlsx)",
  "output_path": "Findlay_Technician_Proficiency_<NewPeriod>.xlsx",
  "new_period_label": "Aug 25 to Sep 9, 2026", // human-readable, used in titles
  "qlik_date_start": "8/25/2026",               // M/D/YYYY, no leading zeros -- matches RO_Header.closedate format
  "qlik_date_end": "9/9/2026"
}
"""
import json
import os

_CONFIG = None


def load_config(path="config.json"):
    global _CONFIG
    if _CONFIG is None:
        if not os.path.exists(path):
            raise FileNotFoundError(
                f"{path} not found. Create it before running the build -- see the docstring "
                f"at the top of this file, or references/config.example.json in the skill folder."
            )
        with open(path) as fh:
            _CONFIG = json.load(fh)
        required = ["roster_path", "qlik_data_path", "template_path", "output_path",
                    "new_period_label", "qlik_date_start", "qlik_date_end"]
        missing = [k for k in required if k not in _CONFIG]
        if missing:
            raise ValueError(f"config.json is missing required keys: {missing}")
    return _CONFIG
