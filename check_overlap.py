import xml.etree.ElementTree as ET

tree = ET.parse("temp_check/original/ppt/slides/slide3.xml")
root = tree.getroot()

print("Shapes in template slide 3:")
for elem in root.iter():
    if elem.tag.endswith("sp"):
        nvSpPr = elem.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}nvSpPr")
        name = "Unknown"
        if nvSpPr is not None:
            cNvPr = nvSpPr.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")
            if cNvPr is not None:
                name = cNvPr.attrib.get("name", "Unknown")
        
        xfrm = elem.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}off")
        ext = elem.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}ext")
        if xfrm is not None and ext is not None:
            x = int(xfrm.attrib.get("x", 0))
            y = int(xfrm.attrib.get("y", 0))
            w = int(ext.attrib.get("cx", 0))
            h = int(ext.attrib.get("cy", 0))
            
            # Print text if any
            texts = []
            for t in elem.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"):
                if t.text:
                    texts.append(t.text)
            
            if texts:
                print(f"Name: {name}, x: {x}, y: {y}, w: {w}, h: {h}")
                print(f"  Text: {''.join(texts)}")
