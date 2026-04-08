---
name: ppt-layout-analyzer
description: Extracts a complete JSON map of all shapes, texts, charts, and their coordinates from a PowerPoint presentation. Use when you need to understand the position of everything in a PowerPoint file.
---

# PPT Layout Analyzer

This skill provides a Python script to "perform surgery" on a PowerPoint file. It parses the entire presentation and extracts the type, content, and exact positioning (Left, Top, Width, Height) of every single shape, text box, image, and chart across all slides. The extracted data is saved to a structured JSON file.

## Usage

When you need to extract the structural map of a PowerPoint file, run the bundled script using `run_shell_command`.

```bash
python .gemini/skills/ppt-layout-analyzer/scripts/analyze_pptx.py <input.pptx> <output.json>
```

### Arguments:
- `<input.pptx>`: The path to the source PowerPoint file you want to analyze.
- `<output.json>`: The path where the resulting JSON layout map should be saved.

### Example Output

The output JSON will look like this:

```json
{
    "file": "presentation.pptx",
    "total_slides": 7,
    "slides": [
        {
            "slide_number": 1,
            "shapes": [
                {
                    "name": "TextBox 1",
                    "position": {
                        "left": 5016283,
                        "top": 995616,
                        "width": 6885895,
                        "height": 4031171
                    },
                    "data": {
                        "type": "TEXT",
                        "content": "Hello World"
                    }
                }
            ]
        }
    ]
}
```

Once generated, you can use the `read_file` or `grep_search` tools on the resulting JSON to locate specific placeholders, charts, or images by their content or type, and find their exact slide numbers and coordinates.
