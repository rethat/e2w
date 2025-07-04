# page_orient.py

from enum import Enum

class Orientation(Enum):
    """Enumeration for page orientation."""
    PORTRAIT = 'portrait'
    LANDSCAPE = 'landscape'


class Size(Enum):
    """Enumeration for page sizes."""
    A3      = (11.69, 16.54)
    A4      = (8.27, 11.69)
    A5      = (5.83, 8.27)
    LETTER  = (8.5, 11)
    LEGAL   = (8.5, 14)
    TABLOID = (11 , 17)
    
    
class PageLayout:
    """Class to represent page layout with orientation and size."""
    
    def __init__(self, orientation: Orientation = Orientation.PORTRAIT, size: Size = Size.A4):
        self._orientation = orientation
        self._size = size
    
    @property
    def orientation(self):
        """Get the page orientation."""
        return self._orientation.value
    
    @property
    def size(self):
        """Get the page size."""
        return self._size.value
    