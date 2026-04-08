---
name: excel-data-analyzer
description: Analyzes Excel game data and outputs JSON based on key findings from a corresponding PowerPoint report. Use when the user asks to analyze an Excel file against a PPTX file and write the result as JSON.
---

# Excel Data Analyzer

This skill helps extract data from an Excel file and a PowerPoint presentation, correlates the raw Excel data with the Key Findings in the PPTX, and outputs a clean JSON file that is easy to understand.

## Workflow

1. **Extract Data**: Use the provided Python script to extract text from the PPTX and data from the XLSX.
   Run the following command:
   ```bash
   python {{skill_dir}}/scripts/extract_data.py "<path_to_pptx>" "<path_to_xlsx>"
   ```

2. **Analyze Output**: Read the output of the script, which contains the text from the PPTX slides and a summary of the data from the Excel sheets.
   
3. **Correlate and Format**:
   - Identify the "Key Findings" in the PPTX text.
   - Look at the Excel data summary and map the corresponding metrics or values that match the Key Findings.
   - Structure this information into a clear JSON format. For example:
     ```json
     {
       "report_summary": {
         "key_findings": [
           {
             "finding": "Player retention increased by 20% in February",
             "supporting_data": {
               "sheet_name": "Retention",
               "metric": "Retention Rate",
               "value": "20%"
             }
           }
         ]
       }
     }
     ```

4. **Write JSON**: Write the final formatted JSON to a file named `analysis_output.json` (or whatever the user specifies) using the `write_file` tool.
