# Create:20241226
#modify codes using class and create testing code
from docx import Document, document
from document_converter import DocumentConverter
import os



from document_format_setter import DocumentFormatSetter
from docx import Document

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION, WD_ORIENT, WD_SECTION
from docx.enum.style import WD_STYLE_TYPE, WD_STYLE
from docx.shared import Cm, Pt, Inches, RGBColor
from docx.oxml.ns import qn

from document_page_number_adder import DocumentPageNumberAdder
from docx import Document


