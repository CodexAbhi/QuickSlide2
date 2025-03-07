from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import os

class PPTGenerator:
    def __init__(self):
        self.ppt = Presentation()
        self.title_slide_layout = self.ppt.slide_layouts[0]
        self.title_content_layout = self.ppt.slide_layouts[1]
        
    def add_title_slide(self, title, subtitle=None):
        """Add title slide to the presentation"""
        slide = self.ppt.slides.add_slide(self.title_slide_layout)
        
        # Set title
        title_shape = slide.shapes.title
        title_shape.text = title
        
        # Set subtitle if provided
        if subtitle:
            subtitle_shape = slide.placeholders[1]
            subtitle_shape.text = subtitle
            
        return slide
    
    def add_section_slide(self, title, content):
        """Add a slide for a section with bullet points"""
        slide = self.ppt.slides.add_slide(self.title_content_layout)
        
        # Set slide title
        title_shape = slide.shapes.title
        title_shape.text = title
        
        # Add content as bullet points
        content_shape = slide.placeholders[1]
        text_frame = content_shape.text_frame
        
        for idx, point in enumerate(content):
            p = text_frame.add_paragraph() if idx > 0 else text_frame.paragraphs[0]
            p.text = point
            p.level = 0
            
        return slide
    
    def apply_theme(self, theme_name="default"):
        """Apply a visual theme to the presentation"""
        # This is a simple implementation that could be expanded
        if theme_name == "default":
            # Apply a blue-based color scheme
            for slide in self.ppt.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text_frame"):
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if shape == slide.shapes.title:
                                    run.font.color.rgb = RGBColor(0, 51, 102)  # Dark blue for titles
                                else:
                                    run.font.color.rgb = RGBColor(0, 0, 0)  # Black for content
    
    def generate_from_content(self, content):
        """Generate a complete PowerPoint from structured content"""
        # Add title slide
        self.add_title_slide(content.get("title", "Presentation"), content.get("subtitle", ""))
        
        # Add content slides
        for section in content.get("sections", []):
            self.add_section_slide(section.get("title", "Section"), section.get("content", []))
        
        # Apply a theme
        self.apply_theme()
        
        return self.ppt
    
    def save(self, filename="presentation.pptx"):
        """Save the presentation to a file"""
        # Ensure the filename has the correct extension
        if not filename.endswith('.pptx'):
            filename += '.pptx'
            
        self.ppt.save(filename)
        return filename