"""
PDF Assembler

Assembles final quarterly report PDF from multiple sources:
1. Main report (Word -> PDF)
2. Rent roll (Excel sheet -> PDF)
3. Management Accounts (Excel sheet -> PDF)
4. Sales tracker (Excel sheet -> PDF)
5. Compliance certificate (Excel sheets -> PDF)
"""

from dataclasses import dataclass
from typing import List, Optional
from pathlib import Path
import subprocess
import tempfile
import os
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from loguru import logger


@dataclass
class PDFSource:
    """Definition of a PDF source for merging."""
    name: str
    source_path: Path
    source_type: str  # 'pdf', 'docx', 'xlsx'
    sheet_name: Optional[str] = None  # For xlsx sources
    page_range: Optional[tuple] = None  # (start, end) 1-indexed, None = all


class PDFAssembler:
    """
    Assembles final merged PDF for ASX submission.
    
    Merge order (per client specification):
    1. Main report (from Word) - Pages 1-7
    2. Rent roll (from Excel) - Pages 8-19
    3. Management Accounts (from Excel) - Pages 20-21
    4. Sales tracker (from Excel) - Pages 22-23
    5. Compliance certificate (from Excel) - Pages 24-28
    """
    
    def __init__(self, output_path: str, working_dir: str = None):
        self.output_path = Path(output_path)
        if working_dir is None:
            self.working_dir = Path(tempfile.gettempdir()) / "pdf_assembly"
        else:
            self.working_dir = Path(working_dir)
        self.working_dir.mkdir(parents=True, exist_ok=True)
    
    def _find_soffice(self) -> str:
        """
        Find LibreOffice soffice command.
        
        Checks common installation paths on different platforms.
        """
        import shutil
        import platform
        
        # Common paths to check
        paths_to_check = [
            'soffice',  # In PATH
            '/usr/bin/soffice',  # Linux
            '/usr/local/bin/soffice',  # Linux homebrew
        ]
        
        # macOS specific paths
        if platform.system() == 'Darwin':
            paths_to_check.extend([
                '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                '/Applications/LibreOffice.app/Contents/MacOS/libreoffice',
                os.path.expanduser('~/Applications/LibreOffice.app/Contents/MacOS/soffice'),
            ])
        
        # Windows paths
        if platform.system() == 'Windows':
            paths_to_check.extend([
                r'C:\Program Files\LibreOffice\program\soffice.exe',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
            ])
        
        # Check each path
        for path in paths_to_check:
            if path == 'soffice':
                # Check if in PATH
                if shutil.which('soffice'):
                    return 'soffice'
            elif os.path.isfile(path):
                return path
        
        raise FileNotFoundError(
            "LibreOffice not found. Please install LibreOffice:\n"
            "  - macOS: brew install --cask libreoffice\n"
            "  - Linux: sudo apt install libreoffice\n"
            "  - Windows: Download from https://www.libreoffice.org/download/"
        )
        
    def assemble(self, sources: List[PDFSource]) -> Path:
        """
        Assemble final PDF from multiple sources.
        
        Args:
            sources: Ordered list of PDF sources
            
        Returns:
            Path to merged PDF
        """
        logger.info(f"Assembling PDF from {len(sources)} sources")
        
        # Step 1: Convert all sources to PDF
        pdf_parts = []
        for idx, source in enumerate(sources):
            logger.info(f"Processing source {idx + 1}/{len(sources)}: {source.name}")
            
            if source.source_type == 'pdf':
                pdf_path = source.source_path
            elif source.source_type == 'docx':
                pdf_path = self._convert_docx_to_pdf(source.source_path)
            elif source.source_type == 'xlsx':
                pdf_path = self._convert_xlsx_to_pdf(
                    source.source_path, 
                    source.sheet_name
                )
            else:
                raise ValueError(f"Unknown source type: {source.source_type}")
            
            # Extract page range if specified
            if source.page_range:
                pdf_path = self._extract_pages(pdf_path, source.page_range)
            
            pdf_parts.append(pdf_path)
        
        # Step 2: Merge all PDFs
        merged_path = self._merge_pdfs(pdf_parts)
        
        # Step 3: Move to final output location
        if merged_path.exists():
            if self.output_path.exists():
                self.output_path.unlink()
            merged_path.rename(self.output_path)
        
        logger.info(f"Final PDF saved to: {self.output_path}")
        return self.output_path
    
    def _convert_docx_to_pdf(self, docx_path: Path) -> Path:
        """Convert Word document to PDF using LibreOffice."""
        output_dir = self.working_dir / "docx_converted"
        output_dir.mkdir(exist_ok=True)
        
        soffice_cmd = self._find_soffice()
        
        cmd = [
            soffice_cmd,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', str(output_dir),
            str(docx_path)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
        
        # Find the output file
        pdf_name = docx_path.stem + '.pdf'
        pdf_path = output_dir / pdf_name
        
        if not pdf_path.exists():
            raise FileNotFoundError(f"Expected PDF not found: {pdf_path}")
        
        return pdf_path
    
    def _convert_xlsx_to_pdf(self, xlsx_path: Path, sheet_name: str = None) -> Path:
        """
        Convert Excel sheet to PDF.
        
        Uses LibreOffice for conversion. If sheet_name specified,
        exports only that sheet.
        """
        output_dir = self.working_dir / "xlsx_converted"
        output_dir.mkdir(exist_ok=True)
        
        if sheet_name:
            # For specific sheet, use Python to extract and convert
            pdf_path = self._export_excel_sheet_to_pdf(xlsx_path, sheet_name, output_dir)
        else:
            # Convert entire workbook
            soffice_cmd = self._find_soffice()
            cmd = [
                soffice_cmd,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(output_dir),
                str(xlsx_path)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0:
                raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
            
            pdf_name = xlsx_path.stem + '.pdf'
            pdf_path = output_dir / pdf_name
        
        return pdf_path
    
    def _export_excel_sheet_to_pdf(self, xlsx_path: Path, sheet_name: str, 
                                    output_dir: Path) -> Path:
        """Export a specific Excel sheet to PDF."""
        import openpyxl
        from copy import copy
        
        # Load workbook
        wb = openpyxl.load_workbook(xlsx_path)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in {xlsx_path}")
        
        # Create a new workbook with just this sheet
        temp_xlsx = output_dir / f"temp_{sheet_name.replace(' ', '_')}.xlsx"
        
        # Copy the sheet to new workbook
        source_sheet = wb[sheet_name]
        new_wb = openpyxl.Workbook()
        new_sheet = new_wb.active
        new_sheet.title = sheet_name
        
        # Copy cell values and formatting
        for row in source_sheet.iter_rows():
            for cell in row:
                new_cell = new_sheet.cell(row=cell.row, column=cell.column)
                new_cell.value = cell.value
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.alignment = copy(cell.alignment)
                    new_cell.number_format = cell.number_format
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
        
        # Copy column widths
        for col_letter, dim in source_sheet.column_dimensions.items():
            new_sheet.column_dimensions[col_letter].width = dim.width
        
        # Copy row heights
        for row_idx, dim in source_sheet.row_dimensions.items():
            new_sheet.row_dimensions[row_idx].height = dim.height
        
        # Set print settings
        new_sheet.page_setup.orientation = 'landscape'
        new_sheet.page_setup.fitToPage = True
        new_sheet.page_setup.fitToWidth = 1
        new_sheet.page_setup.fitToHeight = 0
        
        new_wb.save(temp_xlsx)
        
        # Convert to PDF with LibreOffice
        soffice_cmd = self._find_soffice()
        cmd = [
            soffice_cmd,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', str(output_dir),
            str(temp_xlsx)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")
        
        pdf_path = output_dir / f"temp_{sheet_name.replace(' ', '_')}.pdf"
        
        # Cleanup temp xlsx
        if temp_xlsx.exists():
            temp_xlsx.unlink()
        
        return pdf_path
    
    def _extract_pages(self, pdf_path: Path, page_range: tuple) -> Path:
        """Extract specific pages from PDF."""
        start_page, end_page = page_range
        
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        
        # Convert to 0-indexed
        for page_num in range(start_page - 1, min(end_page, len(reader.pages))):
            writer.add_page(reader.pages[page_num])
        
        output_path = self.working_dir / f"extracted_{pdf_path.stem}.pdf"
        with open(output_path, 'wb') as f:
            writer.write(f)
        
        return output_path
    
    def _merge_pdfs(self, pdf_paths: List[Path]) -> Path:
        """Merge multiple PDFs into one."""
        merger = PdfMerger()
        
        for pdf_path in pdf_paths:
            logger.debug(f"Adding to merge: {pdf_path}")
            merger.append(str(pdf_path))
        
        output_path = self.working_dir / "merged_output.pdf"
        merger.write(str(output_path))
        merger.close()
        
        logger.info(f"Merged {len(pdf_paths)} PDFs into {output_path}")
        return output_path
    
    def add_page_numbers(self, pdf_path: Path) -> Path:
        """Add page numbers to merged PDF."""
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        from io import BytesIO
        
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        
        total_pages = len(reader.pages)
        
        for page_num, page in enumerate(reader.pages, 1):
            # Create page number overlay
            packet = BytesIO()
            c = canvas.Canvas(packet, pagesize=A4)
            
            # Add page number at bottom center
            c.drawString(A4[0] / 2 - 20, 30, f"Page {page_num} of {total_pages}")
            c.save()
            packet.seek(0)
            
            # Merge overlay with page
            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
        
        output_path = self.working_dir / "numbered_output.pdf"
        with open(output_path, 'wb') as f:
            writer.write(f)
        
        return output_path

