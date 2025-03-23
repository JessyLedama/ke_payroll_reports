# -*- coding: utf-8 -*-

from odoo import models, fields, api


class HrEmployee(models.Model):
    _inherit = "hr.employee"

    employee_no = fields.Char(string="Internal Number")
    
    kra_pin = fields.Char(string="KRA PIN")
    
    nssf = fields.Char(string="NSSF")
    
    nhif = fields.Char(string="NHIF")
    
    account_number = fields.Char(string="Account Number")
    
    bank_code = fields.Char(string="Bank Code")
    
    bank_branch = fields.Char(string="Bank Branch")

    disability = fields.Boolean(string="Does Employee Have Disability?", help="Check this box if the employee has a disability and is registered in the Council of Persons with Disability and has a certificate of exemption from the Commissioner of Domestic Taxes.")

    resident = fields.Boolean(string="Resident", help="Check this box if the employee is a resident of Kenya. Such employees are entitled to a tax relief and insurance relief (if they have any)")

    emp_type = fields.Selection([
        ('primary', 'Primary'), 
        ('secondary', 'Secondary')
        ], 
        string="Employee Type", help="[primary] - Select this option if this is the primary employment for the employee. [secondary] - Select this option if this is the secondary employment for the employee. The default is [primary]")

    global_income = fields.Float(string="Global Income (Non Full Time Director)", help="Please record the Global Income of a Non Full time Service director. This amount will be used in computing the taxable pay as per the law")

    helb = fields.Boolean(string="HELB Loan", help="Check this box if the employee is paying HELB.")

    helb_rate = fields.Float(string="HELB Monthly Amount", help="When a new employee with a HELB loan is employed, you must notify HELB in writing within 3 months. They will then advise on the monthly rate. Enter the rate here.")

    @api.model
    def default_get(self, fields_list):
        defaults = super().default_get(fields_list)
        defaults['emp_type'] = 'primary'  # Set default dynamically
        return defaults