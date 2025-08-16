import io
import base64
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.shared import OxmlElement, qn
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import Cm
from docx.shared import Inches

from odoo import api, fields, models, _
from odoo.exceptions import UserError


class SaleReportWizard(models.TransientModel):
    _name = 'sale.report.wizard'
    _description = 'Sale Order Report Wizard'

    order_id = fields.Many2one('sale.order', string='Sale Order', required=True)
    report_format = fields.Selection([
        ('pdf', 'PDF Format'),
        ('docx', 'Word Document (DOCX)')
    ], string='Report Format', default='pdf', required=True)

    def _add_dynamic_terms(self, doc, terms_text):
        """Add dynamic terms and conditions with proper formatting"""
        from docx.shared import RGBColor, Pt, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Split terms by lines and process
        lines = terms_text.split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Check if it's a section header (contains colon)
            if ':' in line and not line.startswith('•') and not line.startswith('-') and not line.startswith('  '):
                # Section title with checkmark
                section_para = doc.add_paragraph()
                checkmark_run = section_para.add_run("✓ ")
                checkmark_run.font.color.rgb = RGBColor(0, 128, 0)  # Green checkmark
                checkmark_run.font.bold = True

                title_run = section_para.add_run(line)
                title_run.font.bold = True
                title_run.font.size = Pt(11)
            else:
                # Regular item or bullet point
                item_para = doc.add_paragraph()

                # Add bullet point if not already present
                if not line.startswith('•') and not line.startswith('-') and not line.startswith('  '):
                    bullet_run = item_para.add_run("• ")
                    bullet_run.font.size = Pt(10)

                item_run = item_para.add_run(line.lstrip('•-').strip())
                item_run.font.size = Pt(10)

                # Indent bullet points
                item_para.paragraph_format.left_indent = Inches(0.5)
                item_para.paragraph_format.first_line_indent = Inches(-0.25)
    def _add_horizontal_line(self, paragraph):
        """Add a horizontal line to paragraph"""
        from docx.oxml.shared import OxmlElement, qn

        p = paragraph._element
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        pBdr.set(qn('w:bottom'), 'single')
        pBdr.set(qn('w:sz'), '6')
        pBdr.set(qn('w:space'), '1')
        pBdr.set(qn('w:color'), 'CCCCCC')
        pPr.append(pBdr)
    def _add_page_number(self, paragraph):
        """Add page number to paragraph"""
        from docx.shared import Pt, RGBColor
        from docx.oxml.shared import OxmlElement, qn

        page_run = paragraph.add_run('Page ')
        page_run.font.size = Pt(9)
        page_run.font.color.rgb = RGBColor(108, 117, 125)

        # Add page number field
        fldChar = OxmlElement('w:fldChar')
        fldChar.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('w:xml:space'), 'preserve')
        instrText.text = 'PAGE'

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')

        fldChar3 = OxmlElement('w:fldChar')
        fldChar3.set(qn('w:fldCharType'), 'end')

        r_element = paragraph.add_run()._element
        r_element.append(fldChar)
        r_element.append(instrText)
        r_element.append(fldChar2)
        r_element.append(fldChar3)

        # Add "of" text
        of_run = paragraph.add_run(' of ')
        of_run.font.size = Pt(9)
        of_run.font.color.rgb = RGBColor(108, 117, 125)

        # Add total pages field
        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'begin')

        instrText2 = OxmlElement('w:instrText')
        instrText2.set(qn('w:xml:space'), 'preserve')
        instrText2.text = 'NUMPAGES'

        fldChar5 = OxmlElement('w:fldChar')
        fldChar5.set(qn('w:fldCharType'), 'separate')

        fldChar6 = OxmlElement('w:fldChar')
        fldChar6.set(qn('w:fldCharType'), 'end')

        r_element2 = paragraph.add_run()._element
        r_element2.append(fldChar4)
        r_element2.append(instrText2)
        r_element2.append(fldChar5)
        r_element2.append(fldChar6)
    def _apply_blue_background_to_cell(self, cell):
        """Apply blue background to table cell"""
        from docx.shared import RGBColor
        from docx.oxml.shared import OxmlElement, qn

        if cell.paragraphs and cell.paragraphs[0].runs:
            run = cell.paragraphs[0].runs[0]
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.bold = True

            shading = OxmlElement('w:shd')
            shading.set(qn('w:fill'), '5DADE2')
            run._element.get_or_add_rPr().append(shading)
    def _add_custom_styles(self, doc):
        """Add custom styles matching the PDF"""
        from docx.shared import Pt, RGBColor
        from docx.enum.style import WD_STYLE_TYPE
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        styles = doc.styles

        # Blue header style
        try:
            blue_header_style = styles['Blue Header']
        except KeyError:
            blue_header_style = styles.add_style('Blue Header', WD_STYLE_TYPE.PARAGRAPH)

        blue_header_font = blue_header_style.font
        blue_header_font.name = 'Arial'
        blue_header_font.size = Pt(12)
        blue_header_font.bold = True
        blue_header_font.color.rgb = RGBColor(255, 255, 255)  # White text

        blue_header_format = blue_header_style.paragraph_format
        blue_header_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        blue_header_format.space_after = Pt(6)
        blue_header_format.space_before = Pt(6)
    def _add_header_footer(self, doc):
        """Add header and footer matching the PDF style"""
        from docx.shared import Pt, RGBColor, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        section = doc.sections[0]

        # Header
        header = section.header
        header_para = header.paragraphs[0]
        header_para.clear()

        # Create simple header layout without complex tables
        # Left side - Logo placeholder
        logo_para = header.add_paragraph()
        logo_para.add_run("[LOGO]").font.size = Pt(10)
        logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Right side - Company info
        company_para = header.add_paragraph()
        company_run = company_para.add_run("Ai Plus")
        company_run.font.size = Pt(24)
        company_run.font.bold = True
        company_run.font.color.rgb = RGBColor(93, 173, 226)
        company_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        tagline_para = header.add_paragraph("Make Your Home More Comfortable")
        tagline_para.runs[0].font.size = Pt(10)
        tagline_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Footer
        footer = section.footer
        footer_para = footer.paragraphs[0]
        footer_para.clear()
        self._add_page_number(footer_para)
        footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    def _remove_table_borders(self, table):
        """Remove borders from table"""
        try:
            from docx.oxml.ns import nsdecls
            from docx.oxml import parse_xml

            for row in table.rows:
                for cell in row.cells:
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    tcBorders = parse_xml(
                        r'<w:tcBorders {}><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'.format(
                            nsdecls('w')))
                    tcPr.append(tcBorders)
        except:
            pass  # Skip border removal if it fails
    def _add_customer_details(self, doc, order):
        """Add customer details section"""
        from docx.shared import Inches, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Section header with blue background
        header = doc.add_paragraph('Customer Details')
        self._apply_blue_background(header)

        # Add customer details as simple paragraphs instead of table
        customer_data = [
            ('Customer Name:', order.partner_id.name or '-'),
            ('Company Name:', order.partner_id.parent_id.name if order.partner_id.parent_id else '-'),
            ('Phone Number:', order.partner_id.phone or order.partner_id.mobile or '-'),
            ('Project:', order.client_order_ref or '-'),
            ('Service Type:', 'Sales Order'),
        ]

        for label, value in customer_data:
            detail_para = doc.add_paragraph()
            label_run = detail_para.add_run(label + " ")
            label_run.font.bold = True
            label_run.font.size = Pt(10)

            value_run = detail_para.add_run(str(value))
            value_run.font.size = Pt(10)

        doc.add_paragraph()  # Add space
    def _add_financial_offer(self, doc, order):
        """Add financial offer section"""
        from docx.shared import Inches, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Section header
        header = doc.add_paragraph('Financial Offer')
        self._apply_blue_background(header)

        # Create simple table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'

        # Header row
        hdr_cells = table.rows[0].cells
        headers = ['Description', 'Quantity', 'Unit Price', 'Total']

        for i, header_text in enumerate(headers):
            if i < len(hdr_cells):
                para = hdr_cells[i].paragraphs[0]
                para.clear()
                run = para.add_run(header_text)
                run.font.bold = True
                run.font.size = Pt(10)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Get currency
        currency = order.currency_id or order.company_id.currency_id
        currency_symbol = currency.symbol if currency else '$'

        # Add order lines
        for line in order.order_line:
            if not line.display_type and line.product_id:
                row_cells = table.add_row().cells

                # Description
                desc = line.name or ''
                if line.product_id.default_code:
                    desc = f"{line.product_id.default_code} - {desc}"

                row_cells[0].paragraphs[0].clear()
                desc_run = row_cells[0].paragraphs[0].add_run(desc)
                desc_run.font.size = Pt(10)

                # Quantity
                row_cells[1].paragraphs[0].clear()
                qty_run = row_cells[1].paragraphs[0].add_run(f"{line.product_uom_qty:.2f}")
                qty_run.font.size = Pt(10)
                row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Unit Price
                row_cells[2].paragraphs[0].clear()
                price_run = row_cells[2].paragraphs[0].add_run(f"{line.price_unit:,.0f} {currency_symbol}")
                price_run.font.size = Pt(10)
                row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Total
                row_cells[3].paragraphs[0].clear()
                total_run = row_cells[3].paragraphs[0].add_run(f"{line.price_subtotal:,.0f} {currency_symbol}")
                total_run.font.size = Pt(10)
                row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add total row
        total_row = table.add_row()
        total_cells = total_row.cells

        # Total Amount label (span 3 columns)
        total_cells[0].paragraphs[0].clear()
        total_label_run = total_cells[0].paragraphs[0].add_run("Total Amount")
        total_label_run.font.bold = True
        total_label_run.font.size = Pt(11)
        total_label_run.font.color.rgb = RGBColor(255, 255, 255)
        total_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Clear middle cells
        total_cells[1].paragraphs[0].clear()
        total_cells[2].paragraphs[0].clear()

        # Total value
        total_cells[3].paragraphs[0].clear()
        total_value_run = total_cells[3].paragraphs[0].add_run(f"{order.amount_total:,.0f} {currency_symbol}")
        total_value_run.font.bold = True
        total_value_run.font.size = Pt(11)
        total_value_run.font.color.rgb = RGBColor(255, 255, 255)
        total_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Apply blue background to total row
        self._apply_blue_background_to_cells([total_cells[0], total_cells[3]])

        doc.add_paragraph()  # Add space
    def _add_additional_service(self, doc, order):
        """Add additional service section"""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Section header
        additional_header = doc.add_paragraph('Additional Service "optional"')
        additional_run = additional_header.runs[0] if additional_header.runs else additional_header.add_run(
            'Additional Service "optional"')
        additional_run.font.bold = True
        additional_run.font.size = Pt(12)

        # Simple table for additional services
        additional_table = doc.add_table(rows=2, cols=2)
        additional_table.style = 'Table Grid'

        # Header row
        header_cells = additional_table.rows[0].cells
        header_cells[0].text = "Description"
        header_cells[1].text = "Unit Price"

        # Make headers bold and centered
        for cell in header_cells:
            para = cell.paragraphs[0]
            if para.runs:
                para.runs[0].font.bold = True
                para.runs[0].font.size = Pt(10)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Data row
        data_cells = additional_table.rows[1].cells
        data_cells[0].text = "Extended Warranty Service"

        currency = order.currency_id or order.company_id.currency_id
        currency_symbol = currency.symbol if currency else '$'
        data_cells[1].text = f"500 {currency_symbol}"
        data_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Set font size for data
        for cell in data_cells:
            para = cell.paragraphs[0]
            if para.runs:
                para.runs[0].font.size = Pt(10)

        doc.add_paragraph()
    def _apply_blue_background_to_cells(self, cells):
        """Apply blue background to specific table cells"""
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        try:
            for cell in cells:
                shading_elm = parse_xml(r'<w:shd {} w:fill="5DADE2"/>'.format(nsdecls('w')))
                cell._tc.get_or_add_tcPr().append(shading_elm)
        except Exception:
            pass  # Skip if XML manipulation fails
    def _add_document_content(self, doc):
        """Generate document content matching the PDF style"""
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Quotation header
        header_para = doc.add_paragraph()
        header_run = header_para.add_run(f"Quotation Number: #{self.order_id.name or ''}")
        header_run.font.bold = True
        header_run.font.size = Pt(11)
        header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        date_para = doc.add_paragraph()
        if self.order_id.date_order:
            date_text = self.order_id.date_order.strftime('%d/%m/%Y')
        else:
            date_text = ''
        date_run = date_para.add_run(f"Date: {date_text}")
        date_run.font.bold = True
        date_run.font.size = Pt(11)
        date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        doc.add_paragraph()  # Space

        # Customer Details Section
        self._add_customer_details(doc, self.order_id)

        # Financial Offer Section
        self._add_financial_offer(doc, self.order_id)

        # Additional Service Section
        self._add_additional_service(doc, self.order_id)

        # Company information
        self._add_company_info(doc)

        # Terms and Conditions
        if self.order_id.note and self.order_id.note.strip():
            doc.add_page_break()
            self._add_terms_conditions(doc)
    def _add_company_info(self, doc):
        """Add company contact information"""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        company = self.env.company
        contact_lines = []

        if company.street:
            address_parts = [company.street]
            if company.city:
                address_parts.append(company.city)
            if company.country_id:
                address_parts.append(company.country_id.name)
            contact_lines.append(f"Address: {' - '.join(address_parts)}")

        if company.phone:
            contact_lines.append(f"Phone Number: {company.phone}")
        if company.email:
            contact_lines.append(f"Email: {company.email}")
        if company.website:
            contact_lines.append(f"Website: {company.website}")

        if contact_lines:
            contact_para = doc.add_paragraph('\n'.join(contact_lines))
            if contact_para.runs:
                contact_para.runs[0].font.size = Pt(9)

        # Tax registration
        if company.vat:
            tax_para = doc.add_paragraph(f"Tax Registration Number: {company.vat}")
            if tax_para.runs:
                tax_para.runs[0].font.size = Pt(9)
            tax_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    def _add_terms_conditions(self, doc):
        """Add terms and conditions section"""
        from docx.shared import RGBColor, Pt, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        order = self.order_id

        # Section header
        header = doc.add_paragraph('Terms & Conditions')
        self._apply_blue_background(header)

        if order.note:
            # Add terms as simple formatted text
            terms_para = doc.add_paragraph(order.note)
            if terms_para.runs:
                terms_para.runs[0].font.size = Pt(10)
    def _apply_blue_background(self, paragraph):
        """Apply blue background to paragraph"""
        from docx.shared import RGBColor, Pt
        from docx.oxml.shared import OxmlElement, qn

        # Ensure paragraph has content
        if not paragraph.runs:
            paragraph.add_run()

        run = paragraph.runs[0]
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.font.bold = True
        run.font.size = Pt(12)

        # Add blue background using simpler method
        try:
            shading = OxmlElement('w:shd')
            shading.set(qn('w:fill'), '5DADE2')
            run._element.get_or_add_rPr().append(shading)
        except Exception:
            pass  # Skip background if it fails
    def action_generate_report(self):
        """Generate report based on selected format"""
        if self.report_format == 'pdf':
            return self._generate_pdf_report()
        elif self.report_format == 'docx':
            return self._generate_docx_report()
    def _generate_pdf_report(self):
        """Generate PDF report using the existing template"""
        return self.env.ref('sale_category_filter.action_report_saleorder_custom').report_action(self.order_id)
    def _add_simple_company_info(self, doc):
        """Add simple company information"""
        from docx.shared import Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        company = self.env.company

        # Address
        if company.street:
            address_parts = [company.street]
            if company.city:
                address_parts.append(company.city)
            if company.country_id:
                address_parts.append(company.country_id.name)  # Fixed: was country_id.name

            address_para = doc.add_paragraph(f"Address: {' - '.join(address_parts)}")
            address_para.runs[0].font.size = Pt(9)

        # Phone
        if company.phone:
            phone_para = doc.add_paragraph(f"Phone Number: {company.phone}")
            phone_para.runs[0].font.size = Pt(9)

        # Email
        if company.email:
            email_para = doc.add_paragraph(f"Email: {company.email}")
            email_para.runs[0].font.size = Pt(9)

        # Website
        if company.website:
            website_para = doc.add_paragraph(f"Website: {company.website}")
            website_para.runs[0].font.size = Pt(9)

        # Tax registration
        if company.vat:
            tax_para = doc.add_paragraph(f"Tax Registration Number: {company.vat}")
            tax_para.runs[0].font.size = Pt(9)
            tax_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    def _generate_docx_report(self):
        """Generate DOCX report with minimal complexity"""
        try:
            from docx import Document
            from docx.shared import Inches, Pt, RGBColor
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            import io
            import base64
        except ImportError:
            raise UserError(
                _('The python-docx library is not installed. Please install it using: pip3 install python-docx'))

        try:
            # Create document
            doc = Document()

            # Set page margins
            sections = doc.sections
            for section in sections:
                section.top_margin = Inches(1.2)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

            # Add simple header
            self._add_simple_header(doc)

            # Generate document content
            self._add_simple_content(doc)

            # Save document to memory
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)

            # Create safe filename
            order_name = self.order_id.name or 'quotation'
            safe_name = ''.join(c for c in order_name if c.isalnum() or c in '-_')
            if not safe_name:
                safe_name = 'quotation'

            # Create attachment
            attachment = self.env['ir.attachment'].create({
                'name': f'Quotation_{safe_name}.docx',
                'type': 'binary',
                'datas': base64.b64encode(doc_buffer.getvalue()),
                'mimetype': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            })

            return {
                'type': 'ir.actions.act_url',
                'url': f'/web/content/{attachment.id}?download=true',
                'target': 'self',
            }

        except Exception as e:
            raise UserError(_('Error generating DOCX report: %s') % str(e))
    def _add_simple_header(self, doc):
        """Add simple header"""
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Company header
        company_para = doc.add_paragraph()
        company_run = company_para.add_run("Ai Plus")
        company_run.font.size = Pt(24)
        company_run.font.bold = True
        company_run.font.color.rgb = RGBColor(93, 173, 226)
        company_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        tagline_para = doc.add_paragraph("Make Your Home More Comfortable")
        tagline_para.runs[0].font.size = Pt(10)
        tagline_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        doc.add_paragraph()  # Space
    def _add_simple_financial_table(self, doc, order):
        """Add simple financial table"""
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Create basic table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'

        # Add headers
        headers = ['Description', 'Quantity', 'Unit Price', 'Total']
        header_cells = table.rows[0].cells

        for i in range(min(len(headers), len(header_cells))):
            cell = header_cells[i]
            para = cell.paragraphs[0]
            para.clear()
            run = para.add_run(headers[i])
            run.font.bold = True
            run.font.size = Pt(10)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Get currency
        currency = order.currency_id or order.company_id.currency_id
        currency_symbol = currency.symbol if currency else '$'

        # Add order lines
        for line in order.order_line:
            if line.product_id:
                row_cells = table.add_row().cells

                # Description
                desc = line.name or ''
                if line.product_id.default_code:
                    desc = f"{line.product_id.default_code} - {desc}"
                row_cells[0].text = desc[:50]  # Limit length

                # Quantity
                row_cells[1].text = f"{line.product_uom_qty:.2f}"
                row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Unit Price
                row_cells[2].text = f"{line.price_unit:,.0f} {currency_symbol}"
                row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Total
                row_cells[3].text = f"{line.price_subtotal:,.0f} {currency_symbol}"
                row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add total row
        total_cells = table.add_row().cells
        total_cells[0].text = "Total Amount"
        total_cells[1].text = ""
        total_cells[2].text = ""
        total_cells[3].text = f"{order.amount_total:,.0f} {currency_symbol}"

        # Make total row bold
        for cell in total_cells:
            para = cell.paragraphs[0]
            if para.runs:
                para.runs[0].font.bold = True
                para.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    def _add_simple_content(self, doc):
        """Add document content matching PDF style"""
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        order = self.order_id

        # Quotation header
        header_para = doc.add_paragraph()
        header_run = header_para.add_run(f"Quotation Number: #{order.name or ''}")
        header_run.font.bold = True
        header_run.font.size = Pt(11)
        header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Date
        date_para = doc.add_paragraph()
        if order.date_order:
            date_text = order.date_order.strftime('%d/%m/%Y')
        else:
            date_text = ''
        date_run = date_para.add_run(f"Date: {date_text}")
        date_run.font.bold = True
        date_run.font.size = Pt(11)
        date_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        doc.add_paragraph()  # Space

        # Customer Details Section
        self._add_customer_details_table(doc, order)

        # Financial Offer Section
        self._add_improved_financial_table(doc, order)

        # Additional Service
        self._add_additional_service_section(doc, order)

        # Company info
        self._add_simple_company_info(doc)

        # Terms and conditions
        if order.note and order.note.strip():
            doc.add_page_break()
            self._add_terms_conditions_section(doc, order)
    def _add_customer_details_table(self, doc, order):
        """Add customer details as a proper table like PDF"""
        from docx.shared import Inches, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        # Customer Details Header
        customer_header = doc.add_paragraph("Customer Details")
        self._apply_simple_blue_background(customer_header)

        # Create customer details table
        customer_table = doc.add_table(rows=5, cols=2)
        customer_table.style = 'Table Grid'

        # Customer data
        customer_data = [
            ('Customer Name', order.partner_id.name or '-'),
            ('Company Name', order.partner_id.parent_id.name if order.partner_id.parent_id else '-'),
            ('Phone Number', order.partner_id.phone or order.partner_id.mobile or '-'),
            ('Project', order.client_order_ref or '-'),
            ('Service Type', 'Sales Order')
        ]

        # Fill table data
        for i, (label, value) in enumerate(customer_data):
            if i < len(customer_table.rows):
                row = customer_table.rows[i]

                # Label cell with gray background
                label_cell = row.cells[0]
                label_para = label_cell.paragraphs[0]
                label_para.clear()
                label_run = label_para.add_run(label)
                label_run.font.bold = True
                label_run.font.size = Pt(10)

                # Add gray background to label cell
                try:
                    shading_elm = parse_xml(r'<w:shd {} w:fill="F8F9FA"/>'.format(nsdecls('w')))
                    label_cell._tc.get_or_add_tcPr().append(shading_elm)
                except:
                    pass

                # Value cell
                value_cell = row.cells[1]
                value_para = value_cell.paragraphs[0]
                value_para.clear()
                value_run = value_para.add_run(str(value))
                value_run.font.size = Pt(10)

        doc.add_paragraph()  # Space
    def _add_improved_financial_table(self, doc, order):
        """Add financial table with subtotals like PDF"""
        from docx.shared import Inches, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        # Financial Offer Header
        financial_header = doc.add_paragraph("Financial Offer")
        self._apply_simple_blue_background(financial_header)

        # Create financial table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'

        # Add headers with gray background
        headers = ['Description', 'Quantity', 'Unit Price', 'Total']
        header_cells = table.rows[0].cells

        for i, header_text in enumerate(headers):
            if i < len(header_cells):
                cell = header_cells[i]
                para = cell.paragraphs[0]
                para.clear()
                run = para.add_run(header_text)
                run.font.bold = True
                run.font.size = Pt(10)
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Add gray background to header
                try:
                    shading_elm = parse_xml(r'<w:shd {} w:fill="E9ECEF"/>'.format(nsdecls('w')))
                    cell._tc.get_or_add_tcPr().append(shading_elm)
                except:
                    pass

        # Get currency
        currency = order.currency_id or order.company_id.currency_id
        currency_symbol = currency.symbol if currency else '$'

        # Add order lines
        subtotal = 0
        for line in order.order_line:
            if line.product_id:
                row_cells = table.add_row().cells

                # Description
                desc = line.name or ''
                if line.product_id.default_code:
                    desc = f"{line.product_id.default_code} - {desc}"

                desc_para = row_cells[0].paragraphs[0]
                desc_para.clear()
                desc_run = desc_para.add_run(desc[:60])
                desc_run.font.size = Pt(10)

                # Quantity
                qty_para = row_cells[1].paragraphs[0]
                qty_para.clear()
                qty_run = qty_para.add_run(f"{line.product_uom_qty:.2f}")
                qty_run.font.size = Pt(10)
                qty_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Unit Price
                price_para = row_cells[2].paragraphs[0]
                price_para.clear()
                price_run = price_para.add_run(f"{line.price_unit:,.0f} {currency_symbol}")
                price_run.font.size = Pt(10)
                price_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Total
                total_para = row_cells[3].paragraphs[0]
                total_para.clear()
                total_run = total_para.add_run(f"{line.price_subtotal:,.0f} {currency_symbol}")
                total_run.font.size = Pt(10)
                total_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                subtotal += line.price_subtotal

        # Add subtotal row if there are multiple lines
        if len([l for l in order.order_line if l.product_id]) > 1:
            subtotal_cells = table.add_row().cells

            # Clear cells
            for cell in subtotal_cells:
                cell.paragraphs[0].clear()

            # Subtotal label (span first 3 columns)
            subtotal_para = subtotal_cells[2].paragraphs[0]
            subtotal_run = subtotal_para.add_run("Subtotal:")
            subtotal_run.font.bold = True
            subtotal_run.font.size = Pt(10)
            subtotal_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # Subtotal value
            subtotal_value_para = subtotal_cells[3].paragraphs[0]
            subtotal_value_run = subtotal_value_para.add_run(f"{subtotal:,.0f} {currency_symbol}")
            subtotal_value_run.font.bold = True
            subtotal_value_run.font.size = Pt(10)
            subtotal_value_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add total row with blue background
        total_cells = table.add_row().cells

        # Clear cells
        for cell in total_cells:
            cell.paragraphs[0].clear()

        # Total Amount label (span first 3 columns)
        total_label_para = total_cells[2].paragraphs[0]
        total_label_run = total_label_para.add_run("Total Amount")
        total_label_run.font.bold = True
        total_label_run.font.size = Pt(11)
        total_label_run.font.color.rgb = RGBColor(255, 255, 255)
        total_label_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Total value
        total_value_para = total_cells[3].paragraphs[0]
        total_value_run = total_value_para.add_run(f"{order.amount_total:,.0f} {currency_symbol}")
        total_value_run.font.bold = True
        total_value_run.font.size = Pt(11)
        total_value_run.font.color.rgb = RGBColor(255, 255, 255)
        total_value_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Apply blue background to total row
        try:
            for cell in [total_cells[2], total_cells[3]]:
                shading_elm = parse_xml(r'<w:shd {} w:fill="5DADE2"/>'.format(nsdecls('w')))
                cell._tc.get_or_add_tcPr().append(shading_elm)
        except:
            pass

        doc.add_paragraph()  # Space
    def _add_additional_service_section(self, doc, order):
        """Add additional service section with table like PDF"""
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        # Additional Service Header
        additional_header = doc.add_paragraph("Additional Service \"optional\"")
        additional_header.runs[0].font.bold = True
        additional_header.runs[0].font.size = Pt(12)

        # Create additional service table
        additional_table = doc.add_table(rows=2, cols=2)
        additional_table.style = 'Table Grid'

        # Header row with gray background
        header_row = additional_table.rows[0]
        header_cells = header_row.cells

        # Description header
        desc_para = header_cells[0].paragraphs[0]
        desc_para.clear()
        desc_run = desc_para.add_run("Description")
        desc_run.font.bold = True
        desc_run.font.size = Pt(10)
        desc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Unit Price header
        price_para = header_cells[1].paragraphs[0]
        price_para.clear()
        price_run = price_para.add_run("Unit Price")
        price_run.font.bold = True
        price_run.font.size = Pt(10)
        price_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add gray background to headers
        try:
            for cell in header_cells:
                shading_elm = parse_xml(r'<w:shd {} w:fill="E9ECEF"/>'.format(nsdecls('w')))
                cell._tc.get_or_add_tcPr().append(shading_elm)
        except:
            pass

        # Data row
        data_row = additional_table.rows[1]
        data_cells = data_row.cells

        # Service description
        service_para = data_cells[0].paragraphs[0]
        service_para.clear()
        service_run = service_para.add_run("Extended Warranty Service")
        service_run.font.size = Pt(10)

        # Service price
        currency = order.currency_id or order.company_id.currency_id
        currency_symbol = currency.symbol if currency else '$'

        price_para = data_cells[1].paragraphs[0]
        price_para.clear()
        price_run = price_para.add_run(f"500 {currency_symbol}")
        price_run.font.size = Pt(10)
        price_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()  # Space
    def _add_terms_conditions_section(self, doc, order):
        """Add terms and conditions with proper formatting"""
        from docx.shared import RGBColor, Pt, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Terms & Conditions Header
        terms_header = doc.add_paragraph("Terms & Conditions")
        self._apply_simple_blue_background(terms_header)

        if order.note:
            # Process terms with proper bullet formatting
            lines = order.note.split('\n')
            current_section = None

            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Check for main section headers
                if any(keyword in line for keyword in
                       ['Payment Terms:', 'Warranty:', 'Delivery/Completion Timeline:', 'Validity period:',
                        'The Offer Not Include:']):
                    # Add section with checkmark
                    section_para = doc.add_paragraph()
                    checkmark_run = section_para.add_run("✓ ")
                    checkmark_run.font.color.rgb = RGBColor(40, 167, 69)  # Green
                    checkmark_run.font.bold = True
                    checkmark_run.font.size = Pt(11)

                    section_run = section_para.add_run(line)
                    section_run.font.bold = True
                    section_run.font.size = Pt(11)

                elif line.startswith(('•', '-', '60%', '30%', '10%')):
                    # Bullet points
                    bullet_para = doc.add_paragraph()
                    bullet_run = bullet_para.add_run("• ")
                    bullet_run.font.size = Pt(10)

                    content_run = bullet_para.add_run(line.lstrip('•-').strip())
                    content_run.font.size = Pt(10)

                    # Indent bullet points
                    bullet_para.paragraph_format.left_indent = Inches(0.5)
                    bullet_para.paragraph_format.first_line_indent = Inches(-0.25)

                else:
                    # Regular content
                    content_para = doc.add_paragraph(line)
                    if content_para.runs:
                        content_para.runs[0].font.size = Pt(10)
                    content_para.paragraph_format.left_indent = Inches(0.25)
    def _apply_simple_blue_background(self, paragraph):
        """Apply blue background matching PDF style"""
        from docx.shared import RGBColor, Pt
        from docx.oxml.shared import OxmlElement, qn

        if not paragraph.runs:
            paragraph.add_run()

        run = paragraph.runs[0]
        run.font.color.rgb = RGBColor(255, 255, 255)  # White text
        run.font.bold = True
        run.font.size = Pt(12)

        # Apply blue background
        try:
            shading = OxmlElement('w:shd')
            shading.set(qn('w:fill'), '5DADE2')  # Blue background
            run._element.get_or_add_rPr().append(shading)
        except:
            pass