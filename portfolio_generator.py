#!/usr/bin/env python3
"""
Portfolio PDF Generator
Generates a professional PDF portfolio from daily project folders containing .docx notes and .pck files.
"""

import os
import re
from pathlib import Path
from datetime import datetime
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib import colors


class PortfolioGenerator:
    """Generate a professional PDF portfolio from daily project folders."""

    def __init__(self, repo_root="."):
        """
        Initialize the portfolio generator.
        
        Args:
            repo_root: Root directory of the repository (default: current directory)
        """
        self.repo_root = Path(repo_root)
        self.pdf_filename = "Networking_Portfolio.pdf"
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()

    def _setup_custom_styles(self):
        """Setup custom paragraph styles for the PDF."""
        # Title style
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=36,
            textColor=HexColor('#1f4788'),
            spaceAfter=30,
            alignment=1,  # Center alignment
            fontName='Helvetica-Bold'
        ))

        # Subtitle style
        self.styles.add(ParagraphStyle(
            name='CustomSubtitle',
            parent=self.styles['Normal'],
            fontSize=14,
            textColor=HexColor('#555555'),
            spaceAfter=12,
            alignment=1  # Center alignment
        ))

        # Day heading style
        self.styles.add(ParagraphStyle(
            name='DayHeading',
            parent=self.styles['Heading2'],
            fontSize=16,
            textColor=HexColor('#2e5c8a'),
            spaceAfter=12,
            spaceBefore=12,
            fontName='Helvetica-Bold'
        ))

        # Content style
        self.styles.add(ParagraphStyle(
            name='BodyContent',
            parent=self.styles['Normal'],
            fontSize=11,
            leading=14,
            spaceAfter=10,
            textColor=HexColor('#333333')
        ))

        # File info style
        self.styles.add(ParagraphStyle(
            name='FileInfo',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=HexColor('#666666'),
            spaceAfter=10,
            fontName='Helvetica-Oblique'
        ))

    def _get_day_folders(self):
        """
        Get all day X folders in numerical order (case-insensitive).
        
        Returns:
            List of day folder paths sorted numerically
        """
        day_folders = []
        for item in self.repo_root.iterdir():
            if item.is_dir():
                match = re.match(r'day\s+(\d+)', item.name, re.IGNORECASE)
                if match:
                    day_number = int(match.group(1))
                    day_folders.append((day_number, item))
        
        # Sort by day number
        day_folders.sort(key=lambda x: x[0])
        return day_folders

    def _extract_docx_text(self, docx_path):
        """
        Extract text from a Word document.
        
        Args:
            docx_path: Path to the .docx file
            
        Returns:
            Extracted text as a string
        """
        try:
            doc = Document(docx_path)
            text = []
            for para in doc.paragraphs:
                if para.text.strip():
                    text.append(para.text)
            return "\n".join(text)
        except Exception as e:
            return f"Error reading document: {str(e)}"

    def _find_pck_file(self, day_folder):
        """
        Find the .pck (Packet Tracer) file in a day folder.
        
        Args:
            day_folder: Path to the day folder
            
        Returns:
            Filename of the .pck file or "Not found"
        """
        for file in day_folder.iterdir():
            if file.suffix.lower() == '.pck':
                return file.name
        return "Not found"

    def _find_docx_file(self, day_folder):
        """
        Find the .docx file in a day folder.
        
        Args:
            day_folder: Path to the day folder
            
        Returns:
            Path to the .docx file or None
        """
        for file in day_folder.iterdir():
            if file.suffix.lower() == '.docx':
                return file
        return None

    def generate_pdf(self):
        """Generate the PDF portfolio."""
        # Create PDF document
        pdf_path = self.repo_root / self.pdf_filename
        doc = SimpleDocTemplate(
            str(pdf_path),
            pagesize=letter,
            rightMargin=0.75*inch,
            leftMargin=0.75*inch,
            topMargin=0.75*inch,
            bottomMargin=0.75*inch
        )

        # Story to hold all elements
        story = []

        # Add title page
        story.extend(self._create_title_page())
        story.append(PageBreak())

        # Get and process day folders
        day_folders = self._get_day_folders()
        
        if not day_folders:
            print("No 'day X' folders found in the repository.")
            return

        total_days = len(day_folders)
        print(f"Found {total_days} day folders. Generating PDF...")

        # Process each day
        for day_number, day_folder in day_folders:
            # Find files
            pck_filename = self._find_pck_file(day_folder)
            docx_file = self._find_docx_file(day_folder)

            # Extract content
            if docx_file:
                content = self._extract_docx_text(docx_file)
            else:
                content = "No Word document found in this folder."

            # Add day section to PDF
            story.extend(self._create_day_section(day_number, content, pck_filename))
            
            # Add page break between days (except for the last day)
            if day_number != day_folders[-1][0]:
                story.append(PageBreak())

            print(f"  Processing day {day_number}... ✓")

        # Build PDF
        try:
            doc.build(story)
            print(f"\n✓ PDF successfully generated: {pdf_path}")
            print(f"  Total days processed: {total_days}")
        except Exception as e:
            print(f"Error generating PDF: {str(e)}")

    def _create_title_page(self):
        """Create the title page elements."""
        elements = []
        
        # Spacer for vertical centering
        elements.append(Spacer(1, 1.5*inch))
        
        # Title
        title = Paragraph("Networking Portfolio", self.styles['CustomTitle'])
        elements.append(title)
        
        # Subtitle
        subtitle = Paragraph("A 60-Day Learning Journey", self.styles['CustomSubtitle'])
        elements.append(subtitle)
        
        # Spacer
        elements.append(Spacer(1, 0.3*inch))
        
        # Generated date
        date_text = f"Generated on {datetime.now().strftime('%B %d, %Y')}"
        date_para = Paragraph(date_text, self.styles['CustomSubtitle'])
        elements.append(date_para)
        
        # Spacer
        elements.append(Spacer(1, 0.5*inch))
        
        # Description
        description = Paragraph(
            "This portfolio contains daily project documentation and notes from a comprehensive "
            "networking learning program. Each section represents one day of study, including "
            "detailed notes and the Packet Tracer files used for hands-on practice.",
            self.styles['BodyContent']
        )
        elements.append(description)
        
        return elements

    def _create_day_section(self, day_number, content, pck_filename):
        """
        Create a day section for the PDF.
        
        Args:
            day_number: Day number
            content: Extracted text content from the Word document
            pck_filename: Name of the .pck file
            
        Returns:
            List of Platypus elements
        """
        elements = []
        
        # Day heading
        day_heading = Paragraph(f"Day {day_number}", self.styles['DayHeading'])
        elements.append(day_heading)
        
        # Packet Tracer file info
        pck_info = f"<b>Packet Tracer File:</b> {pck_filename}"
        pck_para = Paragraph(pck_info, self.styles['FileInfo'])
        elements.append(pck_para)
        
        # Spacer
        elements.append(Spacer(1, 0.15*inch))
        
        # Content from Word document
        if content and content != "No Word document found in this folder.":
            # Truncate content if too long (optional)
            if len(content) > 3000:
                content = content[:3000] + "\n[Content truncated...]"
            content_para = Paragraph(content, self.styles['BodyContent'])
        else:
            content_para = Paragraph(
                "<i>No notes available for this day.</i>",
                self.styles['BodyContent']
            )
        
        elements.append(content_para)
        elements.append(Spacer(1, 0.2*inch))
        
        return elements


def main():
    """Main entry point."""
    print("=" * 60)
    print("Networking Portfolio PDF Generator")
    print("=" * 60)
    
    # Get repository root (current directory)
    repo_root = Path(".")
    
    # Create generator
    generator = PortfolioGenerator(repo_root)
    
    # Generate PDF
    generator.generate_pdf()
    
    print("=" * 60)


if __name__ == "__main__":
    main()
