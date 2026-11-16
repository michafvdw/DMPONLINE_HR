# DMPonline (HR Template) Exporter

This Repository contains pythin code meant for datastewards or other professionals who aim to analyze DMP's. This code is based on Kees den Heijers code from TU delft. See [https://github.com/TU-Delft-DCC/dmponline](https://github.com/TU-Delft-DCC/dmponline) for more information. For this script we are using the **Hogeschool Rotterdam (HR) Template** from DMPonline that was created by the Rotterdam university of applied sciences.

## Features

- Export your DMPS into Excel from DMPonline
- Combine all existing DMP's with the same template in one Excel file
- Include comments in answers 
- Filter on different answers so you can analyze your DMP's 

## Requirements

- Python
- DMPonline API token
- DMPonline Python client
- openpyxl

Install openpyxl:

```bash
pip install openpyxl
```

## Usage

Run the script using:

```bash
python HR_generate_plan_overview_HR_template_all_sections_clean_answers.py -t YOUR_API_TOKEN -u YOUR_EMAIL -o DMP_overview.xlsx
```

### Exclude test plans

Add `--exclude-tests` to exclude test plans:

```bash
python HR_generate_plan_overview_HR_template_all_sections_clean_answers.py -t YOUR_API_TOKEN -u YOUR_EMAIL -o DMP_overview.xlsx --exclude-tests 
```

## Excel Output

The script generates one Excel file containing:

- One row per DMP.
- One column per question.
- DMP name and ID.
- Answers and additional comments.

## Command-line options

| Option | Description |
|---|---|
| `-t` | DMPonline API token (required) |
| `-u` | User email address (required) |
| `-o` | Output Excel filename (required?) |
| `--exclude-tests` | Exclude test plans |

## Notes

text still needed


