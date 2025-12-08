import os
import re
import logging
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, PageBreak
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from datetime import datetime
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

class MASGenerator:
    """Generate Material Approval Sheets (MAS) with company template"""
    
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self.setup_custom_styles()
    
    def setup_custom_styles(self):
        """Setup custom styles for MAS"""
        self.title_style = ParagraphStyle(
            'MASTitle',
            parent=self.styles['Heading1'],
            fontSize=18,
            textColor=colors.HexColor('#1a365d'),
            spaceAfter=12,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        )
        
        self.header_style = ParagraphStyle(
            'MASHeader',
            fontSize=11,
            textColor=colors.black,
            fontName='Helvetica-Bold',
            spaceAfter=6
        )
        
        self.normal_style = ParagraphStyle(
            'MASNormal',
            fontSize=9,
            textColor=colors.black,
            leading=11,
            wordWrap='CJK'
        )

    def _get_logo_path(self):
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
        """Draw properly placed header logo and footer website for MAS PDF."""
        page_width, page_height = doc.pagesize
        gold = colors.HexColor('#d4af37')
        dark = colors.HexColor('#1a365d')

        # Logo centered in header with proper spacing
        logo_path = self._get_logo_path()
        if logo_path and os.path.exists(logo_path):
            try:
                w = 140  # Increased width
                h = 50   # Increased height for full logo visibility
                x = (page_width - w) / 2  # Center horizontally
                y = page_height - 60  # More space from top
                canv.drawImage(logo_path, x, y, width=w, height=h, preserveAspectRatio=True, mask='auto')
            except Exception:
                pass

        # Header gold line - positioned below the logo with proper spacing
        canv.setStrokeColor(gold)
        canv.setLineWidth(2)
        canv.line(doc.leftMargin, page_height - 70, page_width - doc.rightMargin, page_height - 70)

        # Footer with gold line and website centered
        canv.setStrokeColor(gold)
        canv.setLineWidth(2)
        canv.line(doc.leftMargin, doc.bottomMargin + 15, page_width - doc.rightMargin, doc.bottomMargin + 15)
        
        canv.setFillColor(dark)
        canv.setFont('Helvetica', 10)
        footer_text = 'https://alshayaenterprises.com'
        canv.drawCentredString(page_width / 2, doc.bottomMargin + 5, footer_text)
    
    def generate(self, file_id, session):
        """
        Generate Material Approval Sheet
        Returns: path to generated PDF
        """
        # Get file info and extracted data
        uploaded_files = session.get('uploaded_files', [])
        file_info = None
        
        for f in uploaded_files:
            if f['id'] == file_id:
                file_info = f
                break
        
        if not file_info:
            raise Exception('File not found')
        
        # Check if this is multi-budget and get product selections
        is_multibudget = file_info.get('multibudget', False)
        product_selections = file_info.get('product_selections', []) if is_multibudget else []
        tier = file_info.get('tier', 'budgetary') if is_multibudget else None
        
        # Priority: costed_data -> stitched_table -> extraction_result
        items = []
        session_id = session.get('session_id', '')
        
        if 'costed_data' in file_info:
            items = self.parse_items_from_costed_data(file_info['costed_data'], session, file_id,
                                                      is_multibudget, product_selections, tier)
        elif 'stitched_table' in file_info:
            items = self.parse_items_from_stitched_table(file_info['stitched_table'], session, file_id,
                                                        is_multibudget, product_selections, tier)
        elif 'extraction_result' in file_info:
            items = self.parse_items_from_extraction(file_info['extraction_result'], session, file_id)
        else:
            raise Exception('No data available. Please extract tables first.')
        
        if not items:
            raise Exception('No items found in the table. Please check your data.')
        
        # Create output directory
        output_dir = os.path.join('outputs', session_id, 'mas')
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate PDF
        output_file = os.path.join(output_dir, f'mas_{file_id}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')
        
        doc = SimpleDocTemplate(output_file, pagesize=A4,
                                topMargin=0.9*inch, bottomMargin=0.7*inch,
                                leftMargin=0.6*inch, rightMargin=0.6*inch)
        story = []
        
        # Create MAS page for each item
        for idx, item in enumerate(items):
            if idx > 0:
                story.append(PageBreak())
            story.extend(self.create_mas_page(item, idx + 1, len(items)))
        
        # Build PDF
        doc.build(story, onFirstPage=self._draw_header_footer, onLaterPages=self._draw_header_footer)
        
        return output_file
    
    def parse_items_from_costed_data(self, costed_data, session, file_id, is_multibudget=False,
                                     product_selections=None, tier=None):
        """Parse items from costed data"""
        items = []
        session_id = costed_data.get('session_id', session.get('session_id', ''))
        product_selections = product_selections or []
        
        tables = costed_data.get('tables', [])
        for table in tables:
            rows = table.get('rows', [])
            headers = table.get('headers', [])
            
            for row_idx, row in enumerate(rows):
                description = ''
                qty = ''
                unit = ''
                image_path = None
                
                # Iterate through row dictionary items
                image_paths = []
                reference_image_paths = []
                description_found = False
                
                # First pass: look for Brand Description, Brand Image, and reference images
                selected_product_image = None  # For multi-budget: Brand Image from costed table
                
                # PRIORITY 1: For multi-budget, FIRST check for Brand Description and Brand Image
                if is_multibudget:
                    for header, cell_value in row.items():
                        header_str = str(header) if header else ''
                        header_lower = header_str.lower().strip()
                        cell_value = str(cell_value) if cell_value else ''
                        
                        # Brand Description - highest priority for multi-budget
                        if not description_found:
                            # Check multiple variations of Brand Description header
                            if (header_lower == 'brand description' or 
                                header_lower == 'branddescription' or
                                (header_lower.startswith('brand') and 'description' in header_lower)):
                                description = re.sub(r'<[^>]+>', ' ', cell_value).strip()
                                description = re.sub(r'\s+', ' ', description)  # Normalize whitespace
                                if description and description.strip() and 'no description' not in description.lower():
                                    description_found = True
                                    logger.info(f"MAS: Found BRAND DESCRIPTION column '{header}' (length: {len(description)}): {description[:150]}...")
                                    break  # Found it, stop looking
                        
                        # Brand Image - highest priority for multi-budget
                        if not selected_product_image:
                            # Check multiple variations of Brand Image header
                            if (header_lower == 'brand image' or 
                                header_lower == 'brandimage' or
                                (header_lower.startswith('brand') and 'image' in header_lower and 'description' not in header_lower)):
                                if '<img' in cell_value:
                                    paths = self.extract_all_image_paths(cell_value, session_id, file_id)
                                    if paths:
                                        selected_product_image = paths[0]  # Use first Brand Image
                                        logger.info(f"MAS: Found BRAND IMAGE column '{header}' for multi-budget")
                
                # Second pass: Extract other fields (images, description fallback)
                for header, cell_value in row.items():
                    header_str = str(header) if header else ''
                    header_lower = header_str.lower().strip()
                    cell_value = str(cell_value) if cell_value else ''
                    
                    # Check for reference image in indicative image column (for multi-budget)
                    if is_multibudget:
                        # Look for indicative/reference image column (not Brand Image)
                        if (('indicative' in header_lower and 'image' in header_lower) or 
                            ('image' in header_lower and 'brand' not in header_lower and 'product' not in header_lower)):
                            if '<img' in cell_value:
                                paths = self.extract_all_image_paths(cell_value, session_id, file_id)
                                if paths:
                                    reference_image_paths.extend(paths)
                    # Check for other images (non-reference) - only if not multi-budget
                    elif not is_multibudget and '<img' in cell_value:
                        paths = self.extract_all_image_paths(cell_value, session_id, file_id)
                        if paths:
                            image_paths.extend(paths)
                    
                    # Priority: DESCRIPTION column (fallback ONLY if Brand Description not found)
                    if not description_found and ('descript' in header_lower or 'discript' in header_lower) and 'brand' not in header_lower:
                        description = re.sub(r'<[^>]+>', ' ', cell_value).strip()
                        description = re.sub(r'\s+', ' ', description)  # Normalize whitespace
                        description_found = True
                        logger.info(f"MAS: Found DESCRIPTION (fallback) (length: {len(description)}): {description[:150]}...")
                
                # Second pass: other fields and fallback for description
                for header, cell_value in row.items():
                    header_str = str(header) if header else ''
                    header_lower = header_str.lower()
                    cell_value = str(cell_value) if cell_value else ''
                    
                    # Fallback: if no description found, try item/product columns
                    if not description_found and any(h in header_lower for h in ['item', 'product']):
                        description = re.sub(r'<[^>]+>', ' ', cell_value).strip()
                        description = re.sub(r'\s+', ' ', description)
                    
                    # Extract other fields
                    if 'qty' in header_lower or 'quantity' in header_lower:
                        qty = re.sub(r'<[^>]+>', '', cell_value)
                    elif 'unit' in header_lower and 'rate' not in header_lower:
                        unit = re.sub(r'<[^>]+>', '', cell_value)
                
                # For multi-budget: download Brand Image if it's a URL
                if is_multibudget and selected_product_image and selected_product_image.startswith('http'):
                    from utils.image_helper import download_image
                    cached_path = download_image(selected_product_image)
                    if cached_path:
                        selected_product_image = cached_path
                
                # Use Brand Description and Brand Image from costed table for multi-budget
                final_description = description  # Already extracted Brand Description above for multi-budget
                final_image_paths = [selected_product_image] if (is_multibudget and selected_product_image) else (image_paths if image_paths else [])
                
                if final_description:
                    brand = self.extract_brand(final_description)
                    specifications = self.extract_specifications(final_description)
                    
                    item = {
                        'description': final_description,
                        'qty': qty,
                        'unit': unit,
                        'brand': brand,
                        'specifications': specifications,
                        'image_path': final_image_paths[0] if final_image_paths else None,  # Selected product image (big)
                        'image_paths': final_image_paths,  # Selected product images
                        'reference_image_path': reference_image_paths[0] if reference_image_paths else None,  # Reference image (small)
                        'reference_image_paths': reference_image_paths,  # All reference images
                        'is_multibudget': is_multibudget,  # Flag to indicate multi-budget
                        'finish': 'As per manufacturer standard',
                        'warranty': '5 Years'
                    }
                    items.append(item)
        
        return items
    
    def parse_items_from_stitched_table(self, stitched_table, session, file_id, is_multibudget=False,
                                        product_selections=None, tier=None):
        """Parse items from stitched HTML table"""
        items = []
        session_id = session.get('session_id', '')
        product_selections = product_selections or []
        
        html_content = stitched_table.get('html', '')
        if not html_content:
            return items
        
        soup = BeautifulSoup(html_content, 'html.parser')
        table = soup.find('table')
        
        if not table:
            return items
        
        rows = table.find_all('tr')
        if len(rows) < 2:
            return items
        
        # Extract headers from first row
        header_row = rows[0]
        headers = []
        for th in header_row.find_all(['th', 'td']):
            header_text = th.get_text(strip=True).lower()
            # Exclude Product Selection and Actions columns
            if header_text not in ['action', 'actions', 'product selection', 'productselection']:
                headers.append(header_text)
        
        # Process data rows
        for row_idx, row in enumerate(rows[1:]):
            cells = row.find_all('td')
            if not cells:
                continue
            
            row_data = {}
            col_idx = 0
            for idx, cell in enumerate(cells):
                # Skip Product Selection and Actions cells
                if cell.find(class_='product-selection-dropdowns') or cell.find('button'):
                    continue
                text = cell.get_text(strip=True).lower()
                if 'product selection' in text or 'actions' in text:
                    continue
                
                if col_idx < len(headers):
                    # Keep HTML for image detection
                    cell_html = str(cell)
                    row_data[headers[col_idx]] = cell_html
                    col_idx += 1
            
            # Extract fields
            description = ''
            qty = ''
            unit = ''
            reference_image_paths = []
            selected_product_image = None  # For multi-budget: Brand Image from costed table
            description_found = False
            
            for header, cell_value in row_data.items():
                # Ensure header is a string
                header_str = str(header).lower() if header else ''
                
                # For multi-budget: Priority 1 - Brand Description from costed table
                if is_multibudget and not description_found:
                    if 'brand description' in header_str or header_str == 'brand description':
                        description = re.sub(r'<[^>]+>', ' ', str(cell_value)).strip()
                        description = re.sub(r'\s+', ' ', description)  # Normalize whitespace
                        if description and description.strip() and 'no description' not in description.lower():
                            description_found = True
                            logger.info(f"MAS Stitched: Found BRAND DESCRIPTION (length: {len(description)}): {description[:150]}...")
                
                # For multi-budget: Priority 1 - Brand Image from costed table
                if is_multibudget and not selected_product_image:
                    if 'brand image' in header_str or header_str == 'brand image':
                        if '<img' in str(cell_value):
                            img_path = self.extract_image_path(str(cell_value), session_id, file_id)
                            if img_path:
                                selected_product_image = img_path
                                logger.info(f"MAS Stitched: Found BRAND IMAGE column '{header}' for multi-budget")
                
                # Check for reference images in indicative image column (for multi-budget)
                if is_multibudget and (('indicative' in header_str and 'image' in header_str) or 
                                      ('image' in header_str and 'brand' not in header_str and 'product' not in header_str)):
                    if '<img' in str(cell_value):
                        img_path = self.extract_image_path(str(cell_value), session_id, file_id)
                        if img_path:
                            reference_image_paths.append(img_path)
                # Check for other images (non-reference)
                elif '<img' in str(cell_value) and not is_multibudget:
                    img_path = self.extract_image_path(str(cell_value), session_id, file_id)
                    if img_path:
                        reference_image_paths.append(img_path)
                
                # Clean text
                cell_text = re.sub(r'<[^>]+>', '', str(cell_value)).strip()
                
                # Map to fields (fallback if Brand Description not found)
                if not description_found and any(h in header_str for h in ['descript', 'discript', 'item', 'product']) and 'brand' not in header_str:
                    description = cell_text
                elif 'qty' in header_str or 'quantity' in header_str:
                    qty = cell_text
                elif 'unit' in header_str and 'rate' not in header_str:
                    unit = cell_text
            
            # For multi-budget: download Brand Image if it's a URL
            if is_multibudget and selected_product_image and selected_product_image.startswith('http'):
                from utils.image_helper import download_image
                cached_path = download_image(selected_product_image)
                if cached_path:
                    selected_product_image = cached_path
            
            # Use Brand Description and Brand Image from costed table for multi-budget
            final_description = description  # Already extracted Brand Description above for multi-budget
            final_image_paths = [selected_product_image] if (is_multibudget and selected_product_image) else (reference_image_paths if reference_image_paths else [])
            
            if final_description:
                brand = self.extract_brand(final_description)
                specifications = self.extract_specifications(final_description)
                
                item = {
                    'description': final_description,
                    'qty': qty,
                    'unit': unit,
                    'brand': brand,
                    'specifications': specifications,
                    'image_path': final_image_paths[0] if final_image_paths else None,  # Selected product image (big)
                    'image_paths': final_image_paths,  # Selected product images
                    'reference_image_path': reference_image_paths[0] if reference_image_paths else None,  # Reference image (small)
                    'reference_image_paths': reference_image_paths,  # All reference images
                    'is_multibudget': is_multibudget,  # Flag to indicate multi-budget
                    'finish': 'As per manufacturer standard',
                    'warranty': '5 Years'
                }
                items.append(item)
        
        return items
    
    def parse_items_from_extraction(self, extraction_result, session, file_id):
        """Parse items from raw extraction result"""
        items = []
        session_id = session.get('session_id', '')
        
        for layout_result in extraction_result.get('layoutParsingResults', []):
            markdown_text = layout_result.get('markdown', {}).get('text', '')
            images = layout_result.get('markdown', {}).get('images', {})
            
            # Parse tables
            rows = self.extract_table_rows(markdown_text)
            
            for row in rows:
                description = row.get('description', row.get('item', row.get('product', 'N/A')))
                qty = row.get('qty', row.get('quantity', 'N/A'))
                unit = row.get('unit', '')
                
                brand = self.extract_brand(description)
                specifications = self.extract_specifications(description)
                
                # Get first image if available
                image_path = None
                if images:
                    first_img = list(images.values())[0]
                    # Ensure first_img is a string, not a list
                    if isinstance(first_img, (list, tuple)):
                        first_img = first_img[0] if first_img else ''
                    if isinstance(first_img, str) and first_img:
                        image_path = os.path.join('outputs', session_id, file_id, first_img)
                
                item = {
                    'description': description,
                    'qty': qty,
                    'unit': unit,
                    'brand': brand,
                    'specifications': specifications,
                    'image_path': image_path,
                    'finish': 'As per manufacturer standard',
                    'warranty': '5 Years'
                }
                items.append(item)
        
        return items
    
    def create_mas_page(self, item, item_num, total_items):
        """Create complete MAS page for one item"""
        story = []
        
        # Header with logo placeholder
        header_data = [
            ['Material Approval Sheet', f'Item {item_num} of {total_items}']
        ]
        header_table = Table(header_data, colWidths=[5.5*inch, 1.5*inch])
        header_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (0, 0), 14),
            ('FONTNAME', (1, 0), (1, 0), 'Helvetica'),
            ('FONTSIZE', (1, 0), (1, 0), 9),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(header_table)
        
        # Horizontal line
        line_data = [['']]
        line_table = Table(line_data, colWidths=[7*inch])
        line_table.setStyle(TableStyle([
            ('LINEBELOW', (0, 0), (-1, -1), 2, colors.HexColor('#d4af37')),
        ]))
        story.append(line_table)
        story.append(Spacer(1, 0.08*inch))  # Reduced from 0.15
        
        # Project info - more compact
        project_data = [
            ['Project:', '[Project Name]', 'MAS No:', f'MAS-{str(item_num).zfill(3)}'],
            ['Date:', datetime.now().strftime('%d/%m/%Y'), 'Rev:', '00'],
        ]
        project_table = Table(project_data, colWidths=[1*inch, 3*inch, 0.8*inch, 2.2*inch])
        project_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f0f0f0')),
            ('BACKGROUND', (2, 0), (2, -1), colors.HexColor('#f0f0f0')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(project_table)
        story.append(Spacer(1, 0.08*inch))  # Reduced from 0.15
        
        # Item details section
        details_title = Paragraph('<b>ITEM DETAILS</b>', self.header_style)
        story.append(details_title)
        story.append(Spacer(1, 0.05*inch))  # Reduced from 0.08
        
        # Limit description length to fit on one page
        description_text = item.get('description', 'N/A')
        if len(description_text) > 400:
            description_text = description_text[:397] + '...'
        
        # Smaller font for compact layout
        desc_style = ParagraphStyle('DescCompact', parent=self.normal_style, fontSize=7, leading=8)
        
        details_data = [
            ['Description:', Paragraph(description_text, desc_style)],
            ['Brand:', item.get('brand', 'To be specified')],
            ['Quantity:', f"{item.get('qty', 'N/A')} {item.get('unit', '')}"],
            ['Finish:', item.get('finish', 'As per manufacturer standard')],
            ['Warranty:', item.get('warranty', '5 Years')],
        ]
        
        details_table = Table(details_data, colWidths=[1.3*inch, 5.7*inch])
        details_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f8f9fa')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(details_table)
        story.append(Spacer(1, 0.08*inch))  # Reduced from 0.15
        
        # For multi-budget: Add small reference image on the right side below item details table
        is_multibudget = item.get('is_multibudget', False)
        reference_image_path = item.get('reference_image_path')
        if is_multibudget and reference_image_path:
            try:
                # Download if URL
                if reference_image_path.startswith('http'):
                    from utils.image_helper import download_image
                    cached_path = download_image(reference_image_path)
                    if cached_path:
                        reference_image_path = cached_path
                
                if reference_image_path and os.path.exists(reference_image_path):
                    # Small reference image on the right side with label below
                    ref_img = RLImage(reference_image_path, width=0.8*inch, height=0.6*inch)
                    ref_label = Paragraph("Reference Image", ParagraphStyle('RefLabel', fontSize=7, textColor=colors.grey, alignment=2))  # alignment=2 is RIGHT
                    # Create a table with empty left column and reference image on right
                    ref_table = Table([['', ref_img], ['', ref_label]], colWidths=[5.0*inch, 0.8*inch])
                    ref_table.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ]))
                    story.append(ref_table)
                    story.append(Spacer(1, 0.05*inch))
            except Exception as e:
                logger.warning(f"Could not add reference image to MAS: {e}")
        
        # Product image section - support multiple images in grid
        image_title = Paragraph('<b>PRODUCT IMAGE(S)</b>', self.header_style)
        story.append(image_title)
        story.append(Spacer(1, 0.05*inch))  # Reduced from 0.08
        
        image_paths = item.get('image_paths', [])
        if not image_paths and item.get('image_path'):
            image_paths = [item.get('image_path')]
        
        if image_paths:
            # Process and display images
            valid_images = []
            for image_path in image_paths[:9]:  # Max 9 images (increased from 4)
                # If it's a URL, download it first
                if image_path.startswith('http'):
                    from utils.image_helper import download_image
                    cached_path = download_image(image_path)
                    if cached_path:
                        image_path = cached_path
                
                if image_path and os.path.exists(image_path):
                    valid_images.append(image_path)
            
            # Adaptive Grid System
            num_images = len(valid_images)
            
            if num_images == 0:
                pass
            elif num_images == 1:
                # Single large image
                try:
                    from PIL import Image as PILImage
                    pil_img = PILImage.open(valid_images[0])
                    img_width, img_height = pil_img.size
                    aspect_ratio = img_height / img_width
                    target_width = 3.5 * inch  # Large single image
                    target_height = target_width * aspect_ratio
                    # Cap max height
                    if target_height > 3.0 * inch:
                        target_height = 3.0 * inch
                        target_width = target_height / aspect_ratio
                    img = RLImage(valid_images[0], width=target_width, height=target_height)
                    img.hAlign = 'CENTER'
                    story.append(img)
                except Exception as e:
                    logger.error(f"Failed to add image: {e}")
            
            else:
                # multiple images
                try:
                    from PIL import Image as PILImage
                    image_elements = []
                    
                    # Determine columns based on image count
                    if num_images <= 4:
                        cols = 2
                        col_width = 3.5 * inch
                        max_img_width = 3.2 * inch
                        max_img_height = 2.2 * inch
                    else:
                        cols = 3
                        col_width = 2.3 * inch
                        max_img_width = 2.1 * inch
                        max_img_height = 1.6 * inch # Reduce height for 3 rows to fit page
                    
                    for img_path in valid_images:
                        pil_img = PILImage.open(img_path)
                        img_width, img_height = pil_img.size
                        aspect_ratio = img_height / img_width
                        
                        target_width = max_img_width
                        target_height = target_width * aspect_ratio
                        
                        # strict height cap to ensure page fit
                        if target_height > max_img_height:
                            target_height = max_img_height
                            target_width = target_height / aspect_ratio
                            
                        img_elem = RLImage(img_path, width=target_width, height=target_height)
                        image_elements.append(img_elem)
                    
                    # Create grid rows
                    img_table_data = []
                    current_row = []
                    for i, img in enumerate(image_elements):
                        current_row.append(img)
                        if len(current_row) == cols:
                            img_table_data.append(current_row)
                            current_row = []
                    
                    # Fill last row with empty strings if not full
                    if current_row:
                        while len(current_row) < cols:
                            current_row.append('')
                        img_table_data.append(current_row)
                    
                    # Create table
                    img_table = Table(img_table_data, colWidths=[col_width] * cols)
                    img_table.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('LEFTPADDING', (0, 0), (-1, -1), 1),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 1),
                        ('TOPPADDING', (0, 0), (-1, -1), 2),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                    ]))
                    story.append(img_table)
                except Exception as e:
                    logger.error(f"Failed to create image grid: {e}")
        
        story.append(Spacer(1, 0.08*inch))  # Reduced from 0.15
        
        # Technical specifications - compact
        spec_title = Paragraph('<b>SPECIFICATIONS</b>', self.header_style)
        story.append(spec_title)
        story.append(Spacer(1, 0.04*inch))  # Reduced from 0.06
        
        specifications = item.get('specifications', [])
        if specifications:
            # Limit to 3 specs to fit on page (reduced from 4)
            specs_to_show = specifications[:3]
            spec_text = '<br/>'.join([f'• {spec}' for spec in specs_to_show])
            spec_para = Paragraph(spec_text, ParagraphStyle('SpecCompact', parent=self.normal_style, fontSize=7, leading=8))
            story.append(spec_para)
        else:
            compact_specs = '• As per manufacturer standard specifications<br/>• Comply with relevant standards'
            story.append(Paragraph(compact_specs, ParagraphStyle('SpecCompact', parent=self.normal_style, fontSize=7, leading=8)))
        
        story.append(Spacer(1, 0.08*inch))  # Reduced from 0.15
        
        # Approval section - more compact
        approval_title = Paragraph('<b>APPROVAL</b>', self.header_style)
        story.append(approval_title)
        story.append(Spacer(1, 0.04*inch))  # Reduced from 0.06
        
        approval_data = [
            ['Submitted By:', '', 'Date:', ''],
            ['', '', '', ''],
            ['Approved By:', '', 'Date:', ''],
            ['', '', '', ''],
        ]
        
        approval_table = Table(approval_data, colWidths=[1.2*inch, 2.5*inch, 0.7*inch, 2.6*inch])
        approval_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 2), (0, 2), 'Helvetica-Bold'),
            ('FONTNAME', (2, 0), (2, 0), 'Helvetica-Bold'),
            ('FONTNAME', (2, 2), (2, 2), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 7),  # Reduced from 8
            ('LINEBELOW', (1, 1), (1, 1), 1, colors.black),
            ('LINEBELOW', (3, 1), (3, 1), 1, colors.black),
            ('LINEBELOW', (1, 3), (1, 3), 1, colors.black),
            ('LINEBELOW', (3, 3), (3, 3), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        story.append(approval_table)
        story.append(Spacer(1, 0.06*inch))  # Reduced from 0.1
        
        # Remarks - compact with smaller font
        remarks = Paragraph('<b>Remarks:</b> _______________________________________________', ParagraphStyle('RemarksCompact', parent=self.normal_style, fontSize=7))
        story.append(remarks)
        
        return story
    
    def extract_all_image_paths(self, html_content, session_id, file_id):
        """Extract ALL image paths from HTML content (supports multiple images)"""
        image_paths = []
        matches = re.findall(r'src=["\']([^"\']+)["\']', html_content)
        
        for src in matches:
            # Handle URLs (http/https)
            if src.startswith('http://') or src.startswith('https://'):
                image_paths.append(src)
                continue
            
            # Handle leading slash
            if src.startswith('/'):
                src = src[1:]
            
            # Handle local paths
            if src.startswith('outputs/'):
                image_paths.append(src)
                continue
            
            # Handle relative path - ensure all parts are strings
            if isinstance(session_id, (list, tuple)):
                session_id = session_id[0] if session_id else ''
            if isinstance(file_id, (list, tuple)):
                file_id = file_id[0] if file_id else ''
            if isinstance(src, (list, tuple)):
                src = src[0] if src else ''
            
            full_path = os.path.join('outputs', str(session_id), str(file_id), str(src))
            if os.path.exists(full_path):
                image_paths.append(full_path)
            elif os.path.exists(str(src)):
                image_paths.append(str(src))
            else:
                image_paths.append(full_path)
        
        return image_paths if image_paths else None
    
    def extract_image_path(self, html_content, session_id, file_id):
        """Extract first image path from HTML content (for backward compatibility)"""
        paths = self.extract_all_image_paths(html_content, session_id, file_id)
        return paths[0] if paths else None
        
    def _extract_image_path_old(self, html_content, session_id, file_id):
        """Original single image extraction (kept for reference)"""
        match = re.search(r'src=["\']([^"\']+)["\']', html_content)
        if match:
            src = match.group(1)
            
            # Handle URLs (http/https) - return as-is for download logic to handle
            if src.startswith('http://') or src.startswith('https://'):
                return src
            
            # Handle leading slash
            if src.startswith('/'):
                src = src[1:]
            
            # Handle local paths
            if src.startswith('outputs/'):
                return src
            
            # Handle relative path - ensure all parts are strings
            if isinstance(session_id, (list, tuple)):
                session_id = session_id[0] if session_id else ''
            if isinstance(file_id, (list, tuple)):
                file_id = file_id[0] if file_id else ''
            if isinstance(src, (list, tuple)):
                src = src[0] if src else ''
            
            full_path = os.path.join('outputs', str(session_id), str(file_id), str(src))
            if os.path.exists(full_path):
                return full_path
            # Also check if src exists as-is
            if os.path.exists(str(src)):
                return str(src)
            return full_path  # Return even if doesn't exist yet, let download logic handle it
        return None
    
    def extract_table_rows(self, markdown_text):
        """Extract table rows from markdown"""
        lines = markdown_text.split('\n')
        rows = []
        headers = []
        
        for line in lines:
            if '|' in line:
                cells = [cell.strip() for cell in line.split('|') if cell.strip()]
                
                if not headers and cells:
                    headers = [h.lower() for h in cells]
                elif headers and cells and not all(c in '-: ' for c in ''.join(cells)):
                    if len(cells) == len(headers):
                        row = dict(zip(headers, cells))
                        rows.append(row)
        
        return rows
    
    def extract_brand(self, description):
        """Extract brand from description"""
        brands = ['Sedus', 'Narbutas', 'Sokoa', 'B&T', 'Herman Miller', 'Steelcase', 'Vitra', 'Knoll', 'Haworth']
        for brand in brands:
            if brand.lower() in description.lower():
                return brand
        
        # Try to find capitalized words as potential brands
        words = description.split()
        for word in words:
            if word and len(word) > 2 and word[0].isupper():
                return word
        
        return 'To be specified'
    
    def extract_specifications(self, description):
        """Extract specifications from description"""
        specs = []
        
        # Try to extract key specifications from description
        desc_lower = description.lower()
        
        # Material
        if any(mat in desc_lower for mat in ['wood', 'metal', 'fabric', 'leather', 'plastic', 'steel', 'aluminum']):
            specs.append('Material: As specified')
        
        # Finish
        if any(fin in desc_lower for fin in ['polished', 'matte', 'glossy', 'powder coated', 'chrome']):
            specs.append('Finish: As specified')
        
        # Always add these compact specs
        if len(specs) < 2:
            specs.append('Material/Finish: Per manufacturer standard')
        
        specs.append('Color: As per approved sample')
        specs.append('Compliance: Meet relevant standards')
        
        return specs[:4]  # Limit to 4 specs maximum
