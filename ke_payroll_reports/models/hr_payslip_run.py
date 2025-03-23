# -*- coding: utf-8 -*-

from odoo import models, fields


class HrPayslipRun(models.Model):
    _inherit = "hr.payslip.run"

    # adding company to payslip batch
    company_id = fields.Many2one('res.company')