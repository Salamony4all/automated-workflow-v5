import os
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, KeepInFrame, LongTable
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
import json
from datetime import datetime
import re
import logging

import logging
import html

logger = logging.getLogger(__name__)

class OfferGenerator:
    """Generate offer documents with costing factors applied"""
    
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self.setup_custom_styles()
    
    def setup_custom_styles(self):
        """Setup custom paragraph styles"""
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#1a365d'),
            spaceAfter=30,
            alignment=TA_CENTER
        )
        
        self.header_style = ParagraphStyle(
            'CustomHeader',
            parent=self.styles['Heading2'],
            fontSize=16,
            textColor=colors.HexColor('#1a365d'),
            spaceAfter=12
        )
        
        # Compact style for table cells
        self.table_cell_style = ParagraphStyle(
            'TableCell',
            parent=self.styles['Normal'],
            fontSize=7,  # Reduced from 8 to prevent wrapping
            leading=8,
            spaceAfter=0,
            spaceBefore=0,
            leftIndent=0,
            rightIndent=0,
            wordWrap='CJK'  # Better word wrapping
        )
        
        # Extra small style for numeric columns to prevent wrapping (single-line display)
        self.table_numeric_style = ParagraphStyle(
            'TableNumeric',
            parent=self.styles['Normal'],
            fontSize=6,  # Very small for numbers to fit in one line
            leading=7,
            spaceAfter=0,
            spaceBefore=0,
            leftIndent=0,
            rightIndent=0,
            wordWrap='LTR'
        )
        
        # Smaller style for headers to fit in 1-2 lines MAX
        self.table_header_style = ParagraphStyle(
            'TableHeader',
            parent=self.styles['Normal'],
            fontSize=6,  # Small font for compact headers
            leading=7,
            spaceAfter=0,
            spaceBefore=0,
            leftIndent=0,
            rightIndent=0,
            alignment=TA_CENTER,
            wordWrap='CJK'
        )
        
        # Extra small header for tables with many columns (10+)
        self.table_header_tiny_style = ParagraphStyle(
            'TableHeaderTiny',
            parent=self.styles['Normal'],
            fontSize=5,  # Minimum readable size for many columns
            leading=6,
            spaceAfter=0,
            spaceBefore=0,
            leftIndent=0,
            rightIndent=0,
            alignment=TA_CENTER,
            wordWrap='CJK'
        )
        
        # Smaller style for heavy text content (descriptions)
        self.table_description_style = ParagraphStyle(
            'TableDescription',
            parent=self.styles['Normal'],
            fontSize=6,
            leading=7,
            spaceAfter=0,
            spaceBefore=0,
            leftIndent=0,
            rightIndent=0,
            wordWrap='CJK'
        )
    
    def _sanitize_text(self, text):
        """Aggressively sanitize text to remove any Python object representations."""
        if text is None:
            return ''
        
        text = str(text)
        
        # Remove ALL Python object representations using regex
        # Pattern: <ClassName at 0xHEXADDRESS>anything or <ClassName at 0xHEXADDRESS>
        text = re.sub(r'<[A-Za-z_][A-Za-z0-9_]* at 0x[0-9a-fA-F]+>', '', text)
        
        # Also remove patterns like <Paragraph at 0x...>text
        text = re.sub(r'<\w+ at 0x[0-9a-fA-F]+>', '', text)
        
        # Remove any remaining angle brackets that could break XML
        text = text.replace('<', '').replace('>', '')
        
        # Remove 'object at 0x' patterns
        text = re.sub(r'object at 0x[0-9a-fA-F]+', '', text)
        
        # Clean up any resulting double spaces
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text
    
    def _safe_cell(self, text, max_length=200):
        """Create a safe cell value - just returns sanitized plain string."""
        text = self._sanitize_text(text)
        
        # Limit length
        if len(text) > max_length:
            text = text[:max_length-3] + '...'
        
        return text
    
    def _safe_paragraph(self, text, style, max_length=200, bold=False):
        """Create a Paragraph with fully sanitized text for proper text wrapping."""
        # Sanitize first
        text = self._sanitize_text(text)
        
        # Limit length
        if len(text) > max_length:
            text = text[:max_length-3] + '...'
        
        # If empty after sanitization, use placeholder
        if not text:
            text = '-'
        
        # Escape for XML/HTML safety (important for ReportLab)
        text = html.escape(text)
        
        # Wrap in bold if requested
        if bold:
            text = f"<b>{text}</b>"
        
        try:
            return Paragraph(text, style)
        except Exception as e:
            logger.warning(f"Failed to create Paragraph: {e}")
            return text[:50] if text else '-'

    def _get_logo_path(self):
        """Return the best available logo path."""
        candidates = [
            os.path.join('static', 'images', 'AlShaya-Logo-color@2x.png'),
            os.path.join('static', 'images', 'LOGO.png'),
            os.path.join('static', 'images', 'al-shaya-logo-white@2x.png')
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
        return None

    def _draw_header_footer(self, canv: canvas.Canvas, doc):
        """Draw properly placed header logo and footer website."""
        page_width, page_height = doc.pagesize
        gold = colors.HexColor('#d4af37')
        dark = colors.HexColor('#1a365d')

        # Logo centered top header with proper spacing
        logo_path = self._get_logo_path()
        if logo_path and os.path.exists(logo_path):
            try:
                logo_w = 150  # Increased width
                logo_h = 54   # Increased height for full logo visibility
                # Center horizontally
                x = (page_width - logo_w) / 2
                y = page_height - 65  # More space from top for complete logo
                canv.drawImage(logo_path, x, y, width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
            except Exception:
                pass

        # Top separator line positioned below the logo with proper spacing
        canv.setStrokeColor(gold)
        canv.setLineWidth(2)
        canv.line(doc.leftMargin, page_height - 75, page_width - doc.rightMargin, page_height - 75)

        # Footer with gold line and website centered
        canv.setStrokeColor(gold)
        canv.setLineWidth(2)
        canv.line(doc.leftMargin, doc.bottomMargin + 15, page_width - doc.rightMargin, doc.bottomMargin + 15)
        
        canv.setFillColor(dark)
        canv.setFont('Helvetica', 10)
        footer_text = 'https://alshayaenterprises.com'
        # Center the website in footer
        canv.drawCentredString(page_width / 2, doc.bottomMargin + 5, footer_text)
    
    def generate(self, file_id, session):
        """
        Generate offer document
        Returns: path to generated PDF
        """
        # Get file info and costed data
        uploaded_files = session.get('uploaded_files', [])
        file_info = None
        
        for f in uploaded_files:
            if f['id'] == file_id:
                file_info = f
                break
        
        if not file_info:
            raise Exception('File info not found')
        
        # Priority: costed_data -> stitched_table -> extraction_result
        # Check if costed_data exists and has valid tables with data
        costed_data = None
        if 'costed_data' in file_info and file_info['costed_data']:
            costed_data_check = file_info['costed_data']
            # Verify it has the expected structure with tables
            if isinstance(costed_data_check, dict) and 'tables' in costed_data_check:
                tables = costed_data_check.get('tables', [])
                # Check if tables exist and have data
                if tables and len(tables) > 0:
                    # Check if first table has rows
                    if tables[0].get('rows') and len(tables[0]['rows']) > 0:
                        costed_data = costed_data_check
                        logger.info(f"Using existing costed data for offer generation (tables: {len(tables)}, rows in first table: {len(tables[0]['rows'])})")
        
        # Fallback: Parse from stitched table HTML only if costed_data is not available
        if not costed_data and ('stitched_table' in file_info and file_info['stitched_table']):
            logger.info("Parsing stitched table HTML for offer generation (no costed data found)")
            from bs4 import BeautifulSoup
            html_content = file_info['stitched_table']['html']
            soup = BeautifulSoup(html_content, 'html.parser')
            table = soup.find('table')
            
            if not table:
                raise Exception('No table found in stitched data')
            
            # Parse table to costed_data format (excluding Product Selection and Actions columns)
            headers = []
            header_row = table.find('tr')
            if header_row:
                for th in header_row.find_all(['th', 'td']):
                    header_text = th.get_text(strip=True)
                    # Sanitize the header text to remove any object representations
                    header_text = self._sanitize_text(header_text)
                    # Exclude Product Selection and Actions columns
                    if header_text.lower() not in ['action', 'actions', 'product selection', 'productselection']:
                        headers.append(header_text)
            
            logger.info(f"Parsed headers from HTML: {headers[:3] if len(headers) > 3 else headers}")
            
            rows = []
            for row in table.find_all('tr')[1:]:
                cells = row.find_all('td')
                if len(cells) == 0:
                    continue
                
                row_data = {}
                col_idx = 0
                for i, cell in enumerate(cells):
                    # Skip Product Selection and Actions cells
                    if cell.find(class_='product-selection-dropdowns') or cell.find('button'):
                        continue
                    text = cell.get_text(strip=True).lower()
                    if 'product selection' in text or 'actions' in text:
                        continue
                    
                    if col_idx < len(headers):
                        # Keep image HTML if present
                        img = cell.find('img')
                        if img:
                            row_data[headers[col_idx]] = str(cell)
                        else:
                            row_data[headers[col_idx]] = cell.get_text(strip=True)
                        col_idx += 1
                
                if row_data:
                    rows.append(row_data)
            
            # Create costed_data structure from stitched table
            costed_data = {
                'tables': [{
                    'headers': headers,
                    'rows': rows
                }],
                'factors': {},
                'session_id': session.get('session_id', '')
            }
            logger.info("Created costed_data from stitched table HTML (no costed data was available)")
        
        # Final check: if we still don't have costed_data, raise error
        if not costed_data:
            raise Exception('Costed data or stitched table not found. Please apply costing first.')
        
        # Create output directory
        session_id = session['session_id']
        output_dir = os.path.join('outputs', session_id, 'offers')
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate PDF
        output_file = os.path.join(output_dir, f'offer_{file_id}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')
        
        doc = SimpleDocTemplate(output_file, pagesize=A4,
                    topMargin=1.0*inch, bottomMargin=0.8*inch,
                    leftMargin=0.6*inch, rightMargin=0.6*inch)
        story = []
        
        # Title
        title = Paragraph('<font color="#1a365d">COMMERCIAL OFFER</font>', self.title_style)
        story.append(title)
        story.append(Spacer(1, 0.3*inch))
        
        # Company info (placeholder)
        company_info = Paragraph(
            """
            <b><font color="#1a365d">ALSHAYA ENTERPRISES</font></b><br/>
            <font color="#475569">P.O. Box 4451, Kuwait City</font><br/>
            <font color="#475569">Tel: +965 XXX XXXX | Email: info@alshayaenterprises.com</font>
        """,
            self.styles['Normal'])
        story.append(company_info)
        story.append(Spacer(1, 0.3*inch))
        
        # Date
        date_text = Paragraph(f"Date: {datetime.now().strftime('%B %d, %Y')}", self.styles['Normal'])
        story.append(date_text)
        story.append(Spacer(1, 0.5*inch))
        
        # Costing factors removed - confidential information
        
        # Log which data source was used
        if 'factors' in costed_data and costed_data.get('factors'):
            logger.info(f"Using costed_data with factors: {costed_data.get('factors')}")
        else:
            logger.info("Using costed_data (no factors recorded - may be from stitched table)")
        
        # Tables with images
        for idx, table_data in enumerate(costed_data['tables']):
            # Log sample data to verify costed prices are present
            if table_data.get('rows') and len(table_data['rows']) > 0:
                sample_row = table_data['rows'][0]
                # Find price/rate columns in sample row
                price_cols = [k for k in sample_row.keys() if any(term in k.lower() for term in ['rate', 'price', 'amount', 'total'])]
                if price_cols:
                    logger.info(f"Sample price values from first row: {[(col, sample_row.get(col)) for col in price_cols[:3]]}")
            
            header = Paragraph(f"<b><font color='#1a365d'>Item List {idx + 1}</font></b>", self.header_style)
            story.append(header)
            story.append(Spacer(1, 0.2*inch))
            
            # Get session and file info for images
            session_id = session['session_id']
            file_info = None
            uploaded_files = session.get('uploaded_files', [])
            for f in uploaded_files:
                if f['id'] == file_id:
                    file_info = f
                    break
            
            # Prepare table data with images
            table_rows = []
            
            # Headers - clean and format, exclude Action and Product Selection columns
            headers = table_data['headers']
            
            logger.info(f"Original headers (first 3): {headers[:3] if len(headers) > 3 else headers}")
            
            # Clean headers: use _sanitize_text to remove any object representations
            cleaned_headers = []
            header_mapping = {}  # Map cleaned header -> original header STRING (for row lookup)
            
            for h in headers:
                h_str = str(h) if h is not None else ''
                original_h_str = h_str  # Keep original STRING for mapping
                
                # Use the robust sanitize function
                h_clean = self._sanitize_text(h_str)
                
                # Only add non-empty headers
                if h_clean:
                    cleaned_headers.append(h_clean)
                    # Map clean text -> original string representation (for row.get() lookup)
                    header_mapping[h_clean] = original_h_str
                else:
                    logger.warning(f"Skipping empty/malformed header: {h_str[:50]}")
            
            logger.info(f"Cleaned headers (first 3): {cleaned_headers[:3] if len(cleaned_headers) > 3 else cleaned_headers}")
            
            # Filter out Action/Actions and Product Selection columns
            filtered_headers = [h for h in cleaned_headers if h.lower() not in ['action', 'actions', 'product selection', 'productselection']]
            
            logger.info(f"Filtered headers (first 3): {filtered_headers[:3] if len(filtered_headers) > 3 else filtered_headers}")
            logger.info(f"Filtered headers types: {[type(h).__name__ for h in filtered_headers[:3]]}")
            logger.info(f"Filtered headers repr: {[repr(h) for h in filtered_headers[:3]]}")
            
            # Use tiny header style for tables with many columns (10+) to fit in 1-2 lines max
            num_cols = len(filtered_headers)
            header_style = self.table_header_tiny_style if num_cols > 10 else self.table_header_style
            
            # Create header row - use Paragraph for text wrapping
            header_row = []
            for h in filtered_headers:
                # Sanitize and limit length
                h_clean = self._safe_cell(h, max_length=30)
                if not h_clean:
                    h_clean = 'Col'
                logger.info(f"Creating header for: '{h_clean}'")
                # Escape and create Paragraph
                h_escaped = html.escape(h_clean)
                try:
                    p = Paragraph(f"<b>{h_escaped}</b>", header_style)
                    header_row.append(p)
                except:
                    header_row.append(h_clean)
            table_rows.append(header_row)
            
            # Data rows - show only final costed prices with images
            for row in table_data['rows']:
                table_row = []
                
                for h in filtered_headers:
                    # Get original header for data lookup
                    original_h = header_mapping.get(h, h)
                    cell_value = row.get(original_h, '')
                    
                    # Skip original price fields
                    if '_original' in str(original_h):
                        continue
                    
                    # Check if this cell contains an image reference
                    if self.contains_image(cell_value):
                        # Extract ALL image paths from the cell
                        all_image_paths = self.extract_all_image_paths(cell_value, session_id, file_id)
                        
                        # Download any URLs first
                        valid_image_paths = []
                        for img_path in all_image_paths:
                            if img_path and img_path.startswith('http'):
                                from utils.image_helper import download_image
                                cached_path = download_image(img_path)
                                if cached_path and os.path.exists(cached_path):
                                    valid_image_paths.append(cached_path)
                            elif img_path and os.path.exists(img_path):
                                valid_image_paths.append(img_path)
                        
                        if valid_image_paths:
                            try:
                                from PIL import Image as PILImage
                                from reportlab.platypus import Table as InnerTable
                                
                                num_images = len(valid_image_paths)
                                
                                # Calculate image size based on number of images
                                # Smaller images when there are more of them
                                if num_images == 1:
                                    max_size = 1.2 * inch
                                elif num_images == 2:
                                    max_size = 0.7 * inch
                                elif num_images <= 4:
                                    max_size = 0.55 * inch
                                else:
                                    max_size = 0.45 * inch
                                
                                # Create image objects
                                image_objects = []
                                for img_path in valid_image_paths[:6]:  # Max 6 images
                                    try:
                                        pil_img = PILImage.open(img_path)
                                        img_width, img_height = pil_img.size
                                        
                                        # Scale to fit
                                        scale_ratio = min(max_size / img_width, max_size / img_height)
                                        final_width = img_width * scale_ratio
                                        final_height = img_height * scale_ratio
                                        
                                        img = RLImage(img_path, width=final_width, height=final_height)
                                        image_objects.append(img)
                                    except:
                                        pass
                                
                                if image_objects:
                                    if len(image_objects) == 1:
                                        # Single image - just add it
                                        table_row.append(image_objects[0])
                                    else:
                                        # Multiple images - arrange in grid (2 columns)
                                        grid_rows = []
                                        for i in range(0, len(image_objects), 2):
                                            if i + 1 < len(image_objects):
                                                grid_rows.append([image_objects[i], image_objects[i+1]])
                                            else:
                                                grid_rows.append([image_objects[i], ''])
                                        
                                        # Create inner table for image grid
                                        inner_table = InnerTable(grid_rows)
                                        inner_table.setStyle(TableStyle([
                                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                            ('LEFTPADDING', (0, 0), (-1, -1), 1),
                                            ('RIGHTPADDING', (0, 0), (-1, -1), 1),
                                            ('TOPPADDING', (0, 0), (-1, -1), 1),
                                            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
                                        ]))
                                        table_row.append(inner_table)
                                else:
                                    table_row.append('[Img]')
                            except Exception as e:
                                table_row.append('[Img]')
                        else:
                            table_row.append('[Img]')
                    else:
                        # Regular text cell - use Paragraphs for wrapping
                        h_lower = h.lower()
                        if 'descript' in h_lower:
                            cell_style = self.table_description_style
                            max_len = 300  # Keep full description
                        elif 'item' in h_lower or 'product' in h_lower:
                            cell_style = self.table_cell_style
                            max_len = 60
                        elif self.is_numeric_column(h):
                            cell_style = self.table_numeric_style
                            max_len = 12
                        else:
                            cell_style = self.table_cell_style
                            max_len = 20
                        
                        # Sanitize and limit length
                        final_value = self._safe_cell(cell_value, max_length=max_len)
                        
                        # Format numbers nicely
                        if self.is_numeric_column(h) and final_value:
                            try:
                                num_val = float(re.sub(r'[^\d.-]', '', final_value))
                                final_value = f"{num_val:,.2f}"
                            except:
                                pass
                        
                        if not final_value:
                            final_value = '-'
                        
                        # Create Paragraph for text wrapping
                        final_escaped = html.escape(final_value)
                        try:
                            p = Paragraph(final_escaped, cell_style)
                            table_row.append(p)
                        except:
                            table_row.append(final_value)
                
                table_rows.append(table_row)
            
            # Create ReportLab table with appropriate column widths using filtered headers
            col_widths = self.calculate_column_widths(filtered_headers, len(filtered_headers))
            
            # Split large tables into smaller chunks to prevent ReportLab overflow
            MAX_ROWS_PER_TABLE = 15  # Smaller chunks for stability
            
            # Enhanced table styling with WORDWRAP for text cells
            table_style = TableStyle([
                # Header styling
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d4af37')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 6),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 3),
                ('TOPPADDING', (0, 0), (-1, 0), 3),
                ('WORDWRAP', (0, 0), (-1, 0)),  # Enable word wrap for headers
                
                # Data rows styling
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 1), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 5),
                ('TOPPADDING', (0, 1), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 2),
                ('LEFTPADDING', (0, 1), (-1, -1), 2),
                ('RIGHTPADDING', (0, 1), (-1, -1), 2),
                ('WORDWRAP', (0, 1), (-1, -1)),  # Enable word wrap for data
                
                # Grid
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
                
                # Alternating row colors for better readability
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
            ])
            
            try:
                if len(table_rows) > MAX_ROWS_PER_TABLE:
                    # Split into chunks
                    header_row_data = table_rows[0]
                    data_rows = table_rows[1:]
                    
                    for i in range(0, len(data_rows), MAX_ROWS_PER_TABLE - 1):
                        chunk_rows = [header_row_data] + data_rows[i:i + MAX_ROWS_PER_TABLE - 1]
                        
                        # Use regular Table with splitByRow for page breaks
                        t = Table(chunk_rows, colWidths=col_widths, repeatRows=1, splitByRow=True)
                        t.setStyle(table_style)
                        story.append(t)
                        story.append(Spacer(1, 0.15*inch))
                else:
                    # Small table - use regular Table
                    t = Table(table_rows, colWidths=col_widths, repeatRows=1, splitByRow=True)
                    t.setStyle(table_style)
                    story.append(t)
            except Exception as table_error:
                logger.error(f"Failed to create table: {table_error}")
                # Fallback: create a simple text summary
                story.append(Paragraph(f"[Table with {len(table_rows)} rows - see Excel export for details]", self.styles['Normal']))
            
            story.append(Spacer(1, 0.4*inch))
        
        # Summary with updated VAT (5%)
        summary_header = Paragraph("<b><font color='#1a365d'>SUMMARY</font></b>", self.header_style)
        story.append(summary_header)
        story.append(Spacer(1, 0.2*inch))
        
        # Calculate totals
        subtotal = self.calculate_subtotal(costed_data['tables'])
        vat = subtotal * 0.05  # 5% VAT
        grand_total = subtotal + vat
        
        summary_data = [
            ['Subtotal:', f'{subtotal:,.2f}'],
            ['VAT (5%):', f'{vat:,.2f}'],
            ['', ''],  # Empty row for spacing
            ['Grand Total:', f'{grand_total:,.2f}']
        ]
        
        summary_table = Table(summary_data, colWidths=[4*inch, 2*inch])
        summary_style = TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, 2), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 2), 11),
            ('FONTNAME', (0, 3), (-1, 3), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 3), (-1, 3), 14),
            ('TEXTCOLOR', (0, 3), (-1, 3), colors.HexColor('#1a365d')),
            ('LINEABOVE', (0, 3), (-1, 3), 2, colors.HexColor('#d4af37')),
            ('TOPPADDING', (0, 3), (-1, 3), 10),
        ])
        summary_table.setStyle(summary_style)
        story.append(summary_table)
        
        # Terms and conditions
        story.append(Spacer(1, 0.5*inch))
        terms = Paragraph("""
            <b>Terms and Conditions:</b><br/>
            1. Prices are valid for 30 days from the date of this offer.<br/>
            2. Delivery time: 4-6 weeks from order confirmation.<br/>
            3. Payment terms: 50% advance, 50% before delivery.<br/>
            4. Warranty: As per manufacturer's warranty.<br/>
        """, self.styles['Normal'])
        story.append(terms)
        
        # Build PDF
        doc.build(story, onFirstPage=self._draw_header_footer, onLaterPages=self._draw_header_footer)
        
        return output_file
    
    def calculate_subtotal(self, tables):
        """Calculate subtotal from all tables - recalculate totals if needed"""
        subtotal = 0.0
        
        for table in tables:
            headers = table.get('headers', [])
            for row in table['rows']:
                # First, ensure totals are recalculated
                from utils.costing_engine import CostingEngine
                engine = CostingEngine()
                row = engine.recalculate_totals(row, headers)
                
                # Then sum up total/amount columns
                for key, value in row.items():
                    # Look for total/amount columns, exclude original values
                    key_lower = str(key).lower()
                    if (('total' in key_lower or ('amount' in key_lower and 'total' in key_lower)) and 
                        '_original' not in key_lower):
                        try:
                            num_value = float(str(value).replace(',', '').replace('OMR', '').replace('$', '').strip())
                            if not (num_value != num_value) and num_value > 0:  # Check for NaN
                                subtotal += num_value
                        except:
                            pass
        
        return subtotal
    
    def contains_image(self, cell_value):
        """Check if cell contains an image reference"""
        return '<img' in str(cell_value).lower() or 'img_in_' in str(cell_value).lower()
    
    def extract_image_path(self, cell_value, session_id, file_id):
        """Extract image path or URL from cell value"""
        try:
            # Look for img src pattern
            import re
            match = re.search(r'src=["\']([^"\']+)["\']', str(cell_value))
            if match:
                img_path_or_url = match.group(1)
                # If it's a URL, return it as-is (will be downloaded later)
                if img_path_or_url.startswith('http://') or img_path_or_url.startswith('https://'):
                    return img_path_or_url
                
                # Remove leading slash if present
                img_path_or_url = img_path_or_url.lstrip('/')
                # Build absolute path from workspace root
                if img_path_or_url.startswith('outputs'):
                    img_path = img_path_or_url
                else:
                    img_path = os.path.join('outputs', session_id, file_id, img_path_or_url)
                return img_path
            
            # Try to find image reference in text
            if 'img_in_' in str(cell_value):
                match = re.search(r'(imgs/img_in_[^"\s<>]+\.jpg)', str(cell_value))
                if match:
                    img_relative_path = match.group(1)
                    img_path = os.path.join('outputs', session_id, file_id, img_relative_path)
                    return img_path
        except Exception as e:
            pass
        
        return None
    
    def extract_all_image_paths(self, cell_value, session_id, file_id):
        """Extract ALL image paths or URLs from cell value (for multi-image cells)"""
        image_paths = []
        try:
            import re
            cell_str = str(cell_value)
            
            # Find all img src patterns
            matches = re.findall(r'src=["\']([^"\']+)["\']', cell_str)
            for img_path_or_url in matches:
                # If it's a URL, add it directly
                if img_path_or_url.startswith('http://') or img_path_or_url.startswith('https://'):
                    image_paths.append(img_path_or_url)
                else:
                    # Build absolute path
                    img_path_or_url = img_path_or_url.lstrip('/')
                    if img_path_or_url.startswith('outputs'):
                        image_paths.append(img_path_or_url)
                    else:
                        img_path = os.path.join('outputs', session_id, file_id, img_path_or_url)
                        image_paths.append(img_path)
            
            # Also find img_in_ patterns
            img_in_matches = re.findall(r'(imgs/img_in_[^"\s<>]+\.jpg)', cell_str)
            for img_relative_path in img_in_matches:
                img_path = os.path.join('outputs', session_id, file_id, img_relative_path)
                if img_path not in image_paths:
                    image_paths.append(img_path)
        except Exception as e:
            pass
        
        return image_paths
    
    def is_numeric_column(self, header):
        """Check if column likely contains numeric values"""
        numeric_keywords = ['qty', 'quantity', 'rate', 'price', 'amount', 'total', 'cost']
        return any(keyword in header.lower() for keyword in numeric_keywords)
    
    def calculate_column_widths(self, headers, num_cols):
        """Calculate dynamic column widths - prioritize image and description"""
        total_width = 7.5 * inch  # A4 page width minus margins
        
        # Identify column types and assign appropriate widths
        widths = []
        
        for header in headers:
            h_lower = header.lower()
            
            # Image/reference column - PRIORITY: Large for product images
            if 'img' in h_lower or 'image' in h_lower or 'indicative' in h_lower:
                widths.append(1.4 * inch)
            # Description column - PRIORITY: Large for full text
            elif 'descript' in h_lower or 'discript' in h_lower:
                widths.append(2.0 * inch)
            # Item/Product name - medium
            elif 'item' in h_lower or 'product' in h_lower:
                widths.append(0.7 * inch)
            # Location column - small
            elif 'location' in h_lower or 'loc' in h_lower:
                widths.append(0.4 * inch)
            # Serial number - very small
            elif 'sn' in h_lower or 'sl' in h_lower or h_lower in ['no', '#']:
                widths.append(0.25 * inch)
            # Unit column - very small
            elif h_lower == 'unit' or (h_lower.startswith('unit') and 'rate' not in h_lower and 'price' not in h_lower):
                widths.append(0.3 * inch)
            # Quantity columns - small
            elif 'qty' in h_lower or 'quantity' in h_lower:
                widths.append(0.35 * inch)
            # Rate/Price/Amount/Total - compact
            elif 'rate' in h_lower or 'price' in h_lower or 'amount' in h_lower or 'total' in h_lower:
                widths.append(0.5 * inch)
            # Supplier/Brand/Model - medium
            elif 'supplier' in h_lower or 'brand' in h_lower or 'model' in h_lower:
                widths.append(0.5 * inch)
            # All other columns - small
            else:
                widths.append(0.4 * inch)
        
        # Normalize to fit total width
        current_total = sum(widths)
        if current_total > total_width:
            scale_factor = total_width / current_total
            widths = [max(w * scale_factor, 0.35 * inch) for w in widths]  # Ensure minimum after scaling
        elif current_total < total_width * 0.95:
            scale_factor = (total_width * 0.98) / current_total
            widths = [w * scale_factor for w in widths]
        
        return widths
