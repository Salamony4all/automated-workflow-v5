import pandas as pd
import json
import re
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

class CostingEngine:
    """Apply costing factors to extracted tables"""
    
    def __init__(self):
        self.default_factors = {
            'net_margin': 0,
            'freight': 0,
            'customs': 0,
            'installation': 0,
            'exchange_rate': 1.0,
            'additional': 0
        }
    
    def apply_factors(self, file_id, factors, session, table_data=None):
        """
        Apply costing factors to the extracted table
        Args:
            file_id: The file ID
            factors: Dictionary of costing factors
            session: Flask session
            table_data: Optional - table data extracted from DOM (preferred method)
        Returns: Updated table with new prices
        """
        # Get file info from session
        uploaded_files = session.get('uploaded_files', [])
        file_info = None
        
        for f in uploaded_files:
            if f['id'] == file_id:
                file_info = f
                break
        
        if not file_info:
            raise Exception('File not found')
        
        # Use provided table_data if available, otherwise parse from extraction_result
        if table_data:
            tables_data = [table_data]
        elif 'extraction_result' in file_info:
            extraction_result = file_info['extraction_result']
            tables_data = self.parse_markdown_tables(extraction_result)
        else:
            raise Exception('No table data available')
        
        # Apply factors to each row
        updated_tables = []
        for table in tables_data:
            updated_table = self.apply_factors_to_table(table, factors)
            updated_tables.append(updated_table)
        
        # Store costed data in session
        file_info['costed_data'] = {
            'factors': factors,
            'tables': updated_tables,
            'original_table': table_data if table_data else tables_data[0],
            'session_id': session.get('session_id', '')
        }
        
        # Also store in global shared session for cross-app access
        if 'shared_tables' not in session:
            session['shared_tables'] = {}
        session['shared_tables']['costed_table'] = {
            'file_id': file_id,
            'costed_data': file_info['costed_data'],
            'timestamp': datetime.now().isoformat()
        }
        
        session.modified = True
        
        return updated_tables
    
    def parse_markdown_tables(self, extraction_result):
        """
        Parse markdown tables from extraction result
        """
        tables = []
        
        for layout_result in extraction_result.get('layoutParsingResults', []):
            markdown_text = layout_result.get('markdown', {}).get('text', '')
            
            # Split by tables
            table_blocks = self.extract_table_blocks(markdown_text)
            
            for block in table_blocks:
                table_data = self.markdown_table_to_dict(block)
                if table_data:
                    tables.append(table_data)
        
        return tables
    
    def extract_table_blocks(self, markdown_text):
        """
        Extract table blocks from markdown text
        """
        lines = markdown_text.split('\n')
        table_blocks = []
        current_table = []
        in_table = False
        
        for line in lines:
            if '|' in line:
                in_table = True
                current_table.append(line)
            else:
                if in_table and current_table:
                    table_blocks.append('\n'.join(current_table))
                    current_table = []
                in_table = False
        
        if current_table:
            table_blocks.append('\n'.join(current_table))
        
        return table_blocks
    
    def markdown_table_to_dict(self, table_text):
        """
        Convert markdown table to dictionary format
        """
        lines = [line.strip() for line in table_text.split('\n') if line.strip()]
        
        if len(lines) < 2:
            return None
        
        # Extract headers
        header_line = lines[0]
        headers = [h.strip() for h in header_line.split('|') if h.strip()]
        
        # Skip separator line (line with ---)
        data_lines = [line for line in lines[2:] if not all(c in '-|: ' for c in line)]
        
        rows = []
        for line in data_lines:
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if len(cells) == len(headers):
                row = dict(zip(headers, cells))
                rows.append(row)
        
        return {
            'headers': headers,
            'rows': rows
        }
    
    def apply_factors_to_table(self, table_data, factors):
        """
        Apply costing factors to each row in the table
        """
        if not table_data or 'rows' not in table_data:
            return table_data
        
        # Identify price/rate columns
        price_columns = self.identify_price_columns(table_data['headers'])
        
        updated_rows = []
        for row in table_data['rows']:
            updated_row = row.copy()
            
            for col in price_columns:
                if col in row:
                    original_price = self.extract_number(row[col])
                    
                    if original_price is not None:
                        # Apply all factors
                        new_price = original_price
                        
                        # Apply exchange rate
                        new_price *= factors.get('exchange_rate', 1.0)
                        
                        # Apply percentage-based factors
                        new_price *= (1 + factors.get('freight', 0) / 100)
                        new_price *= (1 + factors.get('customs', 0) / 100)
                        new_price *= (1 + factors.get('installation', 0) / 100)
                        new_price *= (1 + factors.get('net_margin', 0) / 100)
                        new_price *= (1 + factors.get('additional', 0) / 100)
                        
                        # Update the row with new price only (no _original)
                        updated_row[col] = f"{new_price:.2f}"
            
            # Recalculate totals if quantity and unit rate exist
            updated_row = self.recalculate_totals(updated_row, table_data['headers'])
            updated_rows.append(updated_row)
        
        return {
            'headers': table_data['headers'],
            'rows': updated_rows,
            'factors_applied': factors
        }
    
    def identify_price_columns(self, headers):
        """
        Identify columns that contain prices/rates
        """
        price_keywords = ['rate', 'price', 'unit rate', 'unit price', 'amount', 'total', 'cost']
        price_columns = []
        
        for header in headers:
            header_lower = header.lower()
            if any(keyword in header_lower for keyword in price_keywords):
                price_columns.append(header)
        
        return price_columns
    
    def extract_number(self, text):
        """
        Extract numeric value from text
        """
        if isinstance(text, (int, float)):
            return float(text)
        
        # Remove currency symbols and commas
        cleaned = re.sub(r'[^\d.-]', '', str(text))
        
        try:
            return float(cleaned)
        except:
            return None
    
    def recalculate_totals(self, row, headers):
        """
        Recalculate total columns based on quantity and unit rate
        """
        qty_col = None
        rate_col = None
        total_col = None
        
        for header in headers:
            header_lower = str(header).lower()
            # Find QTY column
            if ('qty' in header_lower or 'quantity' in header_lower) and 'unit' not in header_lower:
                qty_col = header
            # Find UNIT RATE column - more flexible matching
            elif ('rate' in header_lower or 'price' in header_lower) and 'unit' in header_lower:
                rate_col = header
            elif ('rate' in header_lower or 'price' in header_lower) and 'total' not in header_lower and 'amount' not in header_lower:
                # Fallback: any rate/price column that's not total
                if not rate_col:  # Only use if we haven't found a better match
                    rate_col = header
            # Find TOTAL/AMOUNT column
            elif 'total' in header_lower or ('amount' in header_lower and 'total' in header_lower):
                total_col = header
            elif header_lower == 'amount' and not total_col:
                # Use as fallback for total
                total_col = header
        
        if qty_col and rate_col and total_col:
            qty = self.extract_number(row.get(qty_col, 0))
            rate = self.extract_number(row.get(rate_col, 0))
            
            if qty is not None and rate is not None:
                total = qty * rate
                row[total_col] = f"{total:.2f}"
                logger.info(f"Recalculated total: {qty} * {rate} = {total} for columns {qty_col} * {rate_col} = {total_col}")
        
        return row
