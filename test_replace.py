import re

with open("temp_slides/ppt/slides/slide3.xml") as f: content = f.read()

def replace_paragraph(content, marker, new_text):
    idx = content.find(marker)
    if idx != -1:
        start_p = content.rfind('<a:p>', 0, idx)
        end_p = content.find('</a:p>', idx) + 6
        if start_p != -1 and end_p != -1:
            new_p = f'<a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>{new_text}</a:t></a:r></a:p>'
            return content[:start_p] + new_p + content[end_p:]
    return content

content = replace_paragraph(content, "Although user stickiness improved", "NEW SLIDE 3 FINDING")
print("NEW SLIDE 3 FINDING" in content)
print("Although user stickiness" in content)
