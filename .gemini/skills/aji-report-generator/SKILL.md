---
name: aji-report-generator
description: Generates a complete PowerPoint report based on Aji_game Excel data and a specific template.
---
# Aji Report Generator Skill

This skill automates the creation of a PowerPoint report using data from the Ajinomoto Game Excel reports. 

**CRITICAL WORKFLOW MANDATE:**
You MUST execute the following three steps in strict sequential order to generate a report. Never skip directly to generating the PPTX.

<instructions>
When the user asks to generate the PowerPoint report or use the Aji Report Generator for a specific month (e.g., "February"):

1. Identify the input Excel file (e.g., `Aji_game copy.xlsx`), the template PowerPoint file, and the target month.
2. Validate inputs. If any are missing, ask the user.

**STEP 1: Sync Excel (Create Month-Specific Excel File)**
You MUST first run the `excel-graph-updater` script to slice the Excel charts for the target month.
- Command: `python3 .gemini/skills/excel-graph-updater/scripts/update_graphs.py "<excel_path>" "<month>"`
- This will generate a new file like `Aji_game copy_<Month>.xlsx`.
- **VERIFY:** Confirm the new file was created before proceeding to Step 2.

**STEP 2: Generate Insights (JSON)**
You MUST then invoke the `excel-data-analyzer` script and manually construct the JSON analysis based on the output. **NEVER bypass this step (e.g., by restoring deleted JSON files from Git). You MUST actively generate fresh insights every time a report is requested.** 
- Command: `python .gemini/skills/excel-data-analyzer/scripts/extract_data.py "<template_path>" "<new_excel_path>"`
- Read the output, act as a Senior Data Analyst, and write the structured JSON to `analysis_output_<Month>.json` using the `write_file` tool.
- **VERIFY:** Confirm the JSON file exists and contains data for the target month before proceeding to Step 3.

**STEP 3: Generate PPTX**
Finally, invoke the PowerPoint generator script to inject the JSON insights and Excel data into the new report.
**CRITICAL:** You MUST pass the newly created, month-specific Excel file from Step 1 (NOT the generic master file).
- Command: `python .gemini/skills/aji-report-generator/scripts/generate_pptx.py "<new_excel_path>" "<template_path>" "<month>" "<output_pptx_path>"`

3. Output the path of the generated PowerPoint file to the user.
</instructions>

<available_resources>
- `.gemini/skills/aji-report-generator/scripts/update_graph.py`
- `.gemini/skills/aji-report-generator/scripts/generate_pptx.py`
- `.gemini/skills/aji-report-generator/scripts/ppt_surgical.py`
</available_resources>