# e2w.py

import os
import pandas as pd
from bs4 import BeautifulSoup, Tag, NavigableString
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
import requests
from styles.page_layout import PageLayout, Orientation, Size
from styles.font_family import FontFamily, FontStyle
from styles.table_style import TableFormat


class ExportToWord:
    '''Support converting the template with html tag to Word file'''

    def __init__(self, context:dict, template: str="sample-word.template", output_path: str="output.docx",
                 page_layout: PageLayout = PageLayout(Orientation.LANDSCAPE, Size.A4),
                 font_family: FontFamily = FontFamily(),
                 table_style: TableFormat = TableFormat(),
                 heading_levels: int = 6,
                 error_font: FontFamily = FontFamily(name="Arial", size=8, style=FontStyle.ITALIC, color=RGBColor(255,0,0))
                 ):
        '''Initialize the E2W class with context, template path, and output path.'''
        if not isinstance(context, dict):
            raise ValueError("Context must be a dictionary.")
        if not os.path.exists(template):
            raise FileNotFoundError(f"Template file {template} does not exist.")
        self.template = template
        self.output_path = output_path
        self.context = context
        self.page_layout = page_layout
        self.font_family = font_family
        self.error_font = error_font
        self.page_dimensions = page_layout.size
        self.table_style = table_style
        self.heading_levels = heading_levels
        self.document = self._document_setup()


    def render(self):
        '''Render word document from template'''
        cleaned_template = self._clean_template()
        _content = self._replace_variables(cleaned_template)
        soup = BeautifulSoup(_content, 'html.parser')

        # Fetch all tags in template content
        for tag in soup.find_all(True):
            handler = self._tag_handlers().get(tag.name)
            if handler:
                handler(tag)
            else:
                self._add_paragraph(tag)

        self.document.save(self.output_path)

    def _document_setup(self):
        """Document setup such as page size, orientation, font family, font style, etc."""
        document = Document()
        section = document.sections[0]
        _width, _height = self.page_dimensions
        if self.page_layout.orientation == Orientation.LANDSCAPE.value:
            section.page_width = Inches(_height)
            section.page_height = Inches(_width)
            self.page_dimensions = (_height, _width)  # Update dimensions for landscape
        else:
            section.page_width = Inches(_width)
            section.page_height = Inches(_height)

        # Set up the default font family for the document.
        style = document.styles[self.font_family.style.value]
        style.font.name = self.font_family.name
        style.font.size = self.font_family.size
        style.font.color.rgb = self.font_family.color
        return document
    

    def _set_error_font_style(self, run):
        '''Change the font color to red and italicize it.'''
        run.font.name = self.error_font.name
        run.font.color.rgb = self.error_font.color
        run.font.size = self.error_font.size
        if self.error_font.style == FontStyle.ITALIC:
            run.font.italic = True
        elif self.error_font.style == FontStyle.BOLD:
            run.font.bold = True    


    def _tag_handlers(self):
        """Mapping of tag names to handler methods."""
        return {
            'header': lambda tag : self._handle_header_footer(self.document.sections[0], tag),
            'footer': lambda tag: self._handle_header_footer(self.document.sections[0], tag),
            'title': self._handle_title,
            **{f'h{i}': self._handle_heading for i in range(1, self.heading_levels+1)},  # h1 to h6
            'image': lambda tag: self._handle_image(tag.get('src')) if tag.get('src') else None,
            'dataframe': self._handle_dataframe,
            'base64-image': self._handle_base64_image
        }
    
    def _clean_template(self):
        """Remove all comments in the template content. 
        That is, lines starting with #."""
        with open(self.template, 'r', encoding='utf-8') as f:
            _content = f.read()
        cleaned_lines = []
        for line in _content.split('\n'):
            stripped = line.strip()
            if not stripped.startswith('#'):
                cleaned_lines.append(line)
        return '\n'.join(cleaned_lines)
    

    def _handle_header_footer(self, section, tag): 
        """Process header and footer in the template."""
        _rows = 1
        _columns = 1
        _height = 0.5
        section_part = section.footer
        if tag.name == "header":
            section_part = section.header
            _columns = 3

        _table_width = Inches(self.page_dimensions[0] / _columns)
        table = section_part.add_table(rows=_rows, cols=_columns, width=Inches(self.page_dimensions[0]))
        table.autofit = True
        table.alignment = WD_ALIGN_VERTICAL.CENTER
        
        for column in table.columns:
            column.width = _table_width 
        _first_cell = table.rows[0].cells[0]
        _first_cell.text = tag.text.strip()
        _last_cell = table.rows[0].cells[-1]
        _image = tag.find('image')
        if _image and 'src' in _image.attrs:        
            self._handle_image(_image['src'], aligment=WD_ALIGN_PARAGRAPH.RIGHT, 
                               paragraph=_last_cell.paragraphs[0], height=_height) 
            
    def _handle_image(self, image_path: str, 
                      aligment: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER,
                      paragraph=None, 
                      height: float = 0.0):
        """Insert an image into the document."""
        paragraph = self.document.add_paragraph() if paragraph is None else paragraph
        paragraph.alignment = aligment
        if os.path.exists(image_path):            
            _width, _height = self._get_image_size(image_path, height) if height != 0.0 else self._get_image_size(image_path)
            run = paragraph.add_run()
            run.add_picture(image_path, width=Inches(_width), height=Inches(_height))            
        else:
            run = paragraph.add_run(f"[Missing image: {image_path}]")
            self._set_error_font_style(run)


    def _get_image_size(self, image_path: str, target_height: float = 0.0) -> tuple:
        """Calculates the target size for the image while preserving the aspect ratio.

        Args:
            image_path (str): Path to the image file.
            target_height (float): Desired height of the image in inches.

        Returns:
            tuple: Width and height of the image in inches.
        """
        from PIL import Image
        with Image.open(image_path) as img:
            img_width, img_height = img.size
            aspect_ratio = img_width / img_height
            if target_height == 0.0:
                target_width = self.page_dimensions[0] * 0.6
                target_height = target_width / aspect_ratio
            else:
                target_width = target_height * aspect_ratio
        return (target_width, target_height)
    
    def _handle_title(self, tag: Tag):
        '''Handle the title tag in the template.'''
        paragraph = self.document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(tag.get_text().strip().upper())
        run.bold = True
        run.font.size = Pt(16)  
        
    def _handle_heading(self, tag: Tag):
        """Handle the heading tags in the template."""
        level = int(tag.name[1:-1]) if tag.name[1:-1].isdigit() else 1  # Extract number from h1, h2, etc.
        heading = self.document.add_heading(level=level)
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        heading.add_run(tag.get_text().strip())

    def _handle_dataframe(self, tag: Tag):
        """Handle the dataframe tag in the template."""
        df = pd.DataFrame()
        if 'src' in tag.attrs:
            if os.path.exists(tag['src']):
                df = pd.read_csv(tag['src']) 
        elif 'api' in tag.attrs:
            api_url = tag['api']
            api_config = self.context.get('apis',{}).get(api_url, {})
            headers = self.context.get('api_headers',{}) 
            params = api_config.get('params', {})
            try:
                resp = requests.post(api_url, json=params, headers=headers)
                if resp.status_code == 200:
                    data = resp.json()
                    if data:
                        df = pd.DataFrame(data.get('data',{}))
            except Exception as e:
                self._add_paragraph(str(e))
        if df.empty:
            paragraph = self.document.add_paragraph()
            run = paragraph.add_run(f"{tag.get_text()} No data available.")
            self._set_error_font_style(run)
            return
        
        table = self.document.add_table(rows=1, cols=len(df.columns))
        table.style = self.table_style.style.value
        table.autofit = True
        # header processing
        hdr_cells = table.rows[0].cells
        for i, column in enumerate(df.columns):
            hdr_cells[i].text = column
            hdr_cells[i].paragraphs[0].runs[0].bold=True
        # content processing
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                row_cells[i].text = str(value) if value else ''
                row_cells[i].paragraphs[0].runs[0].bold=False

    def _add_paragraph(self, tag: Tag):
        text = tag.strip() if isinstance(tag, NavigableString) else  str(tag).strip()
        if text:
            self.document.add_paragraph(text)

    def _replace_variables(self, text: str) -> str:
        """Replace placeholders in the content with values from the context."""
        for key, value in self.context.items():
            placeholder = f"<{key}/>"
            if placeholder in text:
                text = text.replace(placeholder, str(value))
        return text 
    
    def _handle_base64_image(self, tag: Tag):
        """Handle base64 encoded images in the template."""
        import base64
        from io import BytesIO
        image_data = tag.get_text().strip()
        if not image_data.startswith('data:image/'):
            raise ValueError("Invalid base64 image data.")
        image_data = image_data.split(',')[1]
        # Decode the base64 image data
        if image_data:
            try:
                image_bytes = base64.b64decode(image_data)
                image_stream = BytesIO(image_bytes)
                self.document.add_picture(image_stream, width=Inches(4))
            except Exception as e:
                paragraph = self.document.add_paragraph()
                run = paragraph.add_run(f"[Error loading image {tag.name}: {e}]")
                self._set_run_font_style(run, font_style=FontStyle.ITALIC)