---
name: excel-data-analyzer
description: Analyzes Excel game data and outputs JSON based on key findings from a corresponding PowerPoint report. Use when the user asks to analyze an Excel file against a PPTX file and write the result as JSON.
---

# Excel Data Analyzer

This skill helps extract data from an Excel file and a PowerPoint presentation, correlates the raw Excel data with the findings on each slide in the PPTX, and outputs a clean, page-by-page JSON file that is easy to understand.

## Workflow

1. **Extract Data**: Use the provided Python script to extract text from the PPTX and data from the XLSX.
   Run the following command:
   ```bash
   python {{skill_dir}}/scripts/extract_data.py "<path_to_pptx>" "<path_to_xlsx>"
   ```

2. **Analyze Output**: Read the output of the script, which contains the text grouped by slide (e.g., "Slide 1", "Slide 2") from the PPTX and a summary of the data from the Excel sheets.
   
3. **Correlate and Format**:
   - Analyze each slide from the PPTX output.
   - For each slide (page), identify the title, sections, and key findings/metrics mentioned.
   - Look at the Excel data summary and map the corresponding metrics or values that support the findings on that specific slide.
   - Structure this information into a comprehensive JSON format grouped by page, similar to the following example:
     ```json
     {
       "report_title": "Monthly Report",
       "pages": [
         {
           "page_number": 3,
           "sections": {
             "User Funnel": {
               "metrics": {
                 "Total click": 572,
                 "Register": 510,
                 "Player": 230,
                 "Conversion rate": "89.1%",
                 "Drop off": "54.9%"
               }
             },
             "User Engagement": {
               "comparison": {
                 "previous_month": "January",
                 "current_month": "February"
               },
               "metrics": {
                 "Daily Active Users (Avg.)": {
                   "January": 5.6,
                   "February": 4.0
                 },
                 "Monthly Active Users": {
                   "January": 82,
                   "February": 37
                 },
                 "User Stickiness": {
                   "January": "6.83%",
                   "February": "10.81%"
                 }
               }
             }
           }
         },
         {
           "page_number": 4,
           "title": "Game Performance (Score)",
           "metrics": {
             "AVG Score per Day": {
               "January": 87248,
               "February": 28986,
               "difference": "-66.8%"
             }
           }
         }
       ]
     }
     ```

4. **Write JSON**: Write the final formatted JSON to a file (like `analysis_output.json`, or whatever the user specifies) using the `write_file` tool.
