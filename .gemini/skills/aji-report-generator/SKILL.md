---
name: aji-report-generator
description: Generates a complete PowerPoint report based on Aji_game Excel data and a specific template.
---
# Aji Report Generator Skill

This skill automates the creation of a PowerPoint report using data from the Ajinomoto Game Excel reports, relying on the scripts provided in the skill directory.

<instructions>
When the user asks to generate the PowerPoint report or use the Aji Report Generator:

1. Identify the input Excel file containing the data.
2. Identify the template PowerPoint file to be used.
3. Identify the target month for the report.
4. If any of these are not provided, ask the user for them.
5. Identify the destination path for the final generated PowerPoint.

Use the provided Python scripts located in this skill's `scripts/` directory to perform the work:

- `update_graph.py`: Can be used to update the Excel file graphs if needed before generating.
- `generate_pptx.py` or `ppt_surgical.py`: Use these to modify the PowerPoint file based on the surgical replacement methods. You may need to review these scripts to see their expected arguments.
    - `ppt_surgical.py` expects arguments: `<excel_path> <template_path> <month> <output_path>`
    - Check the exact script arguments using `run_shell_command` with `-h` or by reading the script headers if necessary.

Proceed step-by-step:
1. Validate inputs.
2. Run the necessary script via `run_shell_command` with the `python` executable.
3. Output the path of the generated PowerPoint file to the user.
</instructions>

<available_resources>
- `.gemini/skills/aji-report-generator/scripts/update_graph.py`
- `.gemini/skills/aji-report-generator/scripts/generate_pptx.py`
- `.gemini/skills/aji-report-generator/scripts/ppt_surgical.py`
</available_resources>