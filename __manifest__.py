{
    "name": "Informe Excel de Recepciones de Compra",
    "summary": "Informe de recepciones de proveedor por rango de fechas con OC, documento y paquetes",
    "version": "15.0.1.0.0",
    "category": "Purchases",
    "author": "Custom",
    "license": "LGPL-3",
    "depends": [
        "purchase_stock",
        "reportes",
    ],
    "external_dependencies": {
        "python": ["xlsxwriter"],
    },
    "data": [
        "security/ir.model.access.csv",
        "views/purchase_reception_report_wizard_views.xml",
    ],
    "installable": True,
    "application": False,
}
