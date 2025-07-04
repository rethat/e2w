# font_family.py

from enum import Enum
from docx.shared import Pt, RGBColor


class FontStyle(Enum):
    """Enumeration for font styles."""
    NORMAL = 'Normal'
    ITALIC = 'Italic'
    BOLD = 'Bold'
    BOLD_ITALIC = 'Bold Italic'
    
    
class FontFamily:
    '''Enumeration for font families.'''

    def __init__(self, name: str="Segoe UI", size: int=10, 
                 style: FontStyle=FontStyle.NORMAL, 
                 color: RGBColor=RGBColor(0, 0, 0)):
        self.name = name
        self.size = Pt(size)
        self.style = style
        self.color = color
    