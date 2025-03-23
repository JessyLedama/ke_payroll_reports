# -*- coding: utf-8 -*-
{
    'name': "Kenyan Payroll Financial Reports",

    'summary': """
       Generate KRA Returns, NSSF, NHIF, HELB, PAYROLL SUMMARY, NET PAY and HOUSING LEVY reports for Payslip batches. 
       """,

    'description': """
        Generate KRA Returns, NSSF, SHIF, HELB, PAYROLL SUMMARY, NET PAY and HOUSING LEVY reports for Payslip batches. 

       This module enables you to generate financial reports for payslip batches. The deductions are input as Salary Rules in Odoo Payroll, which you use in your Salary Structure. Since a Batch Payslip already has the structure and rules applied to your payslips, the system will try to find NSSF, NHIF, HELB etc from your salary rules.
    """,

    'author': "SIMI Technologies",
    'website': "http://simitechnologies.co.ke",

    'category': 'Payroll',
    'version': '0.1',

    # any module necessary for this one to work correctly
    'depends': ['hr', 'hr_payroll_community'],

    # always loaded
    'data': [
        "views/hr_payslip_run_views.xml",
        "views/res_company.xml",
        "views/hr_employee.xml",
        "views/hr_contract.xml",
    ],

    'price': 450,
    'currency': 'USD',
    'license': 'LGPL-3.0',
}
