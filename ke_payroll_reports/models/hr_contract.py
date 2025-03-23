# -*- coding: utf-8 -*-

from odoo import models, fields


class HrContract(models.Model):
    _inherit = "hr.contract"

    car = fields.Boolean(string="Car Benefit")

    house = fields.Boolean(string="Housing Benefit")