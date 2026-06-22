import re
import os
import win32com.client

ppt_file = os.path.abspath('Time_Management_Presentation_v4.pptx')
pdf_file = os.path.abspath('Time_Management_Presentation_WhatsApp.pdf')

with open('create_time_management_ppt.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Restore original V2 circles for Slide 1
content = re.sub(r'# Decorative circles \(top-right\) - kept inside page(.|\n)*?# Decorative dots', 
r'''# Decorative circles (top-right)
    add_shape(slide, MSO_SHAPE.OVAL, Inches(10.5), Inches(-0.5), Inches(2.5), Inches(2.5), Colors.VERY_LIGHT_BLUE)
    add_shape(slide, MSO_SHAPE.OVAL, Inches(11.2), Inches(0.2), Inches(1.5), Inches(1.5), Colors.VERY_LIGHT_PURPLE)
    
    # Decorative circles (bottom-left)
    add_shape(slide, MSO_SHAPE.OVAL, Inches(-0.8), Inches(5.5), Inches(2.5), Inches(2.5), Colors.VERY_LIGHT_PINK)
    add_shape(slide, MSO_SHAPE.OVAL, Inches(0.5), Inches(6.0), Inches(1.5), Inches(1.5), Colors.VERY_LIGHT_MINT)
    
    # Decorative dots''', content)

# Restore original for Slide 2
content = re.sub(r'# Decorative circle - kept inside page\n.*?VERY_LIGHT_BLUE\)',
r'''# Decorative circle
    add_shape(slide, MSO_SHAPE.OVAL, Inches(11.5), Inches(5.5), Inches(2.5), Inches(2.5), Colors.VERY_LIGHT_BLUE)''', content)

# Restore original for Slide 8
content = re.sub(r'# Decorative element - kept inside page\n.*?VERY_LIGHT_PURPLE\)',
r'''# Decorative element
    add_shape(slide, MSO_SHAPE.OVAL, Inches(11.5), Inches(-0.5), Inches(2.5), Inches(2.5), Colors.VERY_LIGHT_PURPLE)''', content)

content = re.sub(r'# Right icon - kept inside page\n.*?18\)',
r'''# Right icon
        add_circle_icon(slide, Inches(11.5), y + Inches(0.45), Inches(0.3), color, icon, Colors.WHITE, 18)''', content)

# Restore original for Slide 11
content = re.sub(r'# Decorative circles - kept fully inside page, smaller.*?\n.*?VERY_LIGHT_PINK\)',
r'''# Decorative circles
    add_shape(slide, MSO_SHAPE.OVAL, Inches(-1), Inches(-1), Inches(3), Inches(3), Colors.VERY_LIGHT_BLUE)
    add_shape(slide, MSO_SHAPE.OVAL, Inches(11), Inches(-1), Inches(3), Inches(3), Colors.VERY_LIGHT_PURPLE)
    add_shape(slide, MSO_SHAPE.OVAL, Inches(-1), Inches(5.5), Inches(3), Inches(3), Colors.VERY_LIGHT_MINT)
    add_shape(slide, MSO_SHAPE.OVAL, Inches(11), Inches(5.5), Inches(3), Inches(3), Colors.VERY_LIGHT_PINK)''', content, flags=re.DOTALL)

# Let's keep the Slide 10 Fixes since the user explicitly wanted the overlap fixed over the text!

# Finally update the output filename
content = re.sub(r'Time_Management_Presentation_v[0-9]+\.pptx', 'Time_Management_Presentation_v4.pptx', content)

with open('create_time_management_ppt.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("File patched.")
