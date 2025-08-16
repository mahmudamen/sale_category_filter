from odoo import api, fields, models


class SaleOrder(models.Model):
    _inherit = 'sale.order'

    categ_id = fields.Many2one(
        'product.category',
        string='Product Category',
        help='Select a product category to filter products in order lines'
    )
    # Additional fields for the custom report
    project_name = fields.Char(
        string='Project Name',
        help='Project name for this quotation'
    )

    service_type = fields.Char(
        string='Service Type',
        help='Type of service being provided'
    )

    def action_open_report_wizard(self):
        """Open report generation wizard"""
        return {
            'name': 'Generate Custom Report',
            'type': 'ir.actions.act_window',
            'res_model': 'sale.report.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_order_id': self.id,
            }
        }

    @api.onchange('categ_id')
    def _onchange_categ_id(self):
        """Clear product selection in order lines when category changes"""
        if self.categ_id:
            # Auto-fill service type based on category
            self.service_type = self.categ_id.name

            for line in self.order_line:
                if line.product_id and line.product_id.categ_id:
                    # Check if product is not in the selected category or its children
                    category_ids = [int(x) for x in line.product_id.categ_id.parent_path.split('/') if x]
                    if self.categ_id.id not in category_ids:
                        line.product_id = False
                        line.product_template_id = False





class SaleOrderLine(models.Model):
    _inherit = 'sale.order.line'


    @api.model
    def default_get(self, fields_list):
        result = super().default_get(fields_list)
        return result