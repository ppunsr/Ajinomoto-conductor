# Implementation Plan: Aji Report Generator & Data Automation

## Objective
Create and maintain specialized Gemini CLI skills to automate the extraction, analysis, and reporting of Ajinomoto Game data. The core feature is the `aji-report-generator` skill, which performs surgical XML modifications on Excel and PowerPoint files to generate a high-fidelity monthly report while preserving all template formatting.

## Core Workflow: Precision Automation Process

The reporting workflow is designed to generate high-fidelity PowerPoint reports by surgically modifying the internal XML of Office files. This approach ensures 100% preservation of template styles, which standard libraries often lose.

### 1. Data Source & Preparation
The process starts with two main inputs:
*   **Source Excel:** A master file (e.g., `Aji_game copy.xlsx`) containing raw daily and monthly game statistics across multiple sheets (User Engagement, User Funnel, Gameplay Scores, etc.).
*   **PPTX Template:** A professionally designed template (e.g., `Merkle Thailand -Ajipanda's Kitchen report- 260331 copy.pptx`) containing pre-styled slides and complex charts.

### 2. Phase 1: Excel Graph Synchronization (`update_graphs.py`)
Before generating the PowerPoint, the Excel file itself must be updated so that its internal charts reflect the data for the requested month.
*   **Targeting:** The script scans the Excel sheets using `openpyxl` to locate the specific row/column ranges for the target month (e.g., finding where "Feb" starts and ends in the "User Engagement" sheet).
*   **XML Patching:** It unzips the `.xlsx` file and modifies `xl/charts/chart*.xml`. Using regex, it updates the formulas (`<c:f>`) and cached values within the XML to point to the new data ranges.
*   **Result:** A new Excel file (e.g., `Aji_game_March.xlsx`) where all internal charts are correctly "sliced" for that month.

### 3. Phase 2: Data Analysis & Insight Generation (`excel-data-analyzer`)
Once the new Excel file is created, the system must generate data-driven insights tailored to the new month's numbers.
*   **Execution:** The `excel-data-analyzer` skill (or a senior analyst agent) extracts the raw numbers from the newly created Excel file and correlates them against the template's previous findings.
*   **Output:** It generates a structured JSON file (e.g., `analysis_output_March.json`) that contains key metrics and uniquely drafted `key_finding` paragraphs for each slide (e.g., summarizing engagement drops or score trends).

### 4. Phase 3: PowerPoint Report Generation (`generate_pptx.py`)
This is the core engine that builds the final 7-page report. It follows three sub-steps:

**A. Data Extraction & Analytics**
The script reads the updated Excel to calculate key performance indicators (KPIs) like DAU, MAU, Stickiness, Funnel drop-offs, and Score/Time averages.

**B. Surgical XML Manipulation**
The PPTX template is unzipped into a temporary workspace. The script then performs "surgery" on the XML files:
*   **Text Replacement:** In `ppt/slides/slide*.xml`, it finds text placeholders (like "January") and replaces them with the target month name and the calculated metrics. It also injects the tailored `key_finding` paragraphs loaded from the `analysis_output_[Month].json` file.
*   **Chart Data Injection:** In `ppt/charts/chart*.xml`, it identifies data series. It replaces the XML nodes for categories (`<c:cat>`) and values (`<c:val>`) with the raw numbers extracted from the Excel file.
*   **ChartEx Support:** It includes specialized logic for modern Office charts (`chartEx*.xml`), which have a different structure than legacy charts.

**C. Final Assembly**
Once the XMLs are patched, the script zips the entire directory back into a `.pptx` file. Because only the data values were changed, the original fonts, colors, branding, and complex chart layouts remain exactly as they were in the template.

## AI Agent Workflow Mandate (Strict Sequential Execution)
When asked to generate a report, the AI agent MUST strictly follow these sequential steps without skipping:
1.  **STEP 1: Sync Excel (`excel-graph-updater`)**: Run `update_graphs.py` to slice the master Excel file for the requested month, creating a new month-specific `.xlsx` file (e.g., `Aji_game copy_Feb.xlsx`). **Do not proceed until you verify this new file was created successfully.**
2.  **STEP 2: Analyze & Correlate (`excel-data-analyzer`)**: Run `extract_data.py` on the newly created Excel file (e.g., `Aji_game copy_Feb.xlsx`) and template PPTX. The AI must then write a structured JSON file (e.g., `analysis_output_Feb.json`) containing tailored, data-driven key findings. **Do not proceed until the JSON file is fully written and verified.**
3.  **STEP 3: Generate PPTX (`aji-report-generator`)**: Run `generate_pptx.py` to inject the raw Excel data and JSON insights into the PowerPoint template's XML, creating the final `.pptx` report. **CRITICAL: You MUST use the newly created Excel file from Step 1 (e.g., `Aji_game copy_Feb.xlsx`) as the input for this command, NOT the generic master file.**

## Skill Locations
-   `.gemini/skills/aji-report-generator`: The primary skill for generating the end-to-end report.
-   `.gemini/skills/excel-graph-updater`: Standalone skill for just Phase 1 (updating Excel charts).
-   `.gemini/skills/excel-data-analyzer`: Skill for extracting and correlating data.
-   `.gemini/skills/ppt-layout-analyzer`: Utility skill for mapping PowerPoint layouts.