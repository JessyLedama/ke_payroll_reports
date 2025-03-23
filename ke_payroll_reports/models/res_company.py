# -*- coding: utf-8 -*-

from odoo import models, fields


class ResCompany(models.Model):
    _inherit = "res.company"

    # adding required details for Kenyan localisation
    nssf = fields.Char(string="NSSF Number", required=True)
    nhif = fields.Char(string="NHIF Number", required=True)