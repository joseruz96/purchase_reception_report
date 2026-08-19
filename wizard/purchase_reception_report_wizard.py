# -*- coding: utf-8 -*-
import base64
import io
import re
import unicodedata
from datetime import datetime, time

import pytz
import xlsxwriter

from odoo import _, api, fields, models
from odoo.exceptions import UserError, ValidationError


class PurchaseReceptionReportWizard(models.TransientModel):
    _name = "purchase.reception.report.wizard"
    _description = "Informe Excel de Recepciones de Compra"

    date_from = fields.Date(
        string="Fecha desde",
        required=True,
        default=lambda self: fields.Date.context_today(self).replace(day=1),
    )
    date_to = fields.Date(
        string="Fecha hasta",
        required=True,
        default=fields.Date.context_today,
    )
    partner_id = fields.Many2one(
        "res.partner",
        string="Proveedor",
        domain="[('supplier_rank', '>', 0)]",
    )
    company_id = fields.Many2one(
        "res.company",
        string="Compañía",
        required=True,
        default=lambda self: self.env.company,
    )
    file_data = fields.Binary(readonly=True)
    file_name = fields.Char(readonly=True)

    @api.constrains("date_from", "date_to")
    def _check_dates(self):
        for wizard in self:
            if wizard.date_from and wizard.date_to and wizard.date_from > wizard.date_to:
                raise ValidationError(_("La fecha desde no puede ser posterior a la fecha hasta."))

    @staticmethod
    def _normalize(value):
        value = value or ""
        value = unicodedata.normalize("NFKD", str(value))
        value = "".join(char for char in value if not unicodedata.combining(char))
        return value.strip().lower()

    def _date_bounds_utc(self):
        """Convierte el rango de fechas del usuario a UTC para buscar date_done."""
        self.ensure_one()
        tz_name = self.env.user.tz or "UTC"
        tz = pytz.timezone(tz_name)

        start_local = tz.localize(datetime.combine(self.date_from, time.min))
        end_local = tz.localize(datetime.combine(self.date_to, time.max))

        start_utc = start_local.astimezone(pytz.UTC).replace(tzinfo=None)
        end_utc = end_local.astimezone(pytz.UTC).replace(tzinfo=None)
        return start_utc, end_utc

    def _find_text_field(self, record, technical_candidates, label_candidates):
        """Busca un campo de texto por nombre técnico y, si no existe, por etiqueta visible."""
        for field_name in technical_candidates:
            if field_name in record._fields:
                value = getattr(record, field_name, False)
                if value:
                    return value

        normalized_labels = {self._normalize(label) for label in label_candidates}
        for field_name, field in record._fields.items():
            if field.type not in ("char", "text", "html"):
                continue
            label = self._normalize(field.string)
            if label in normalized_labels or any(candidate in label for candidate in normalized_labels):
                value = getattr(record, field_name, False)
                if value:
                    return value
        return ""

    @staticmethod
    def _plain_text(value):
        if not value:
            return ""
        text = str(value)
        text = re.sub(r"<br\s*/?>", " ", text, flags=re.I)
        text = re.sub(r"</p>", " ", text, flags=re.I)
        text = re.sub(r"<[^>]+>", "", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def _get_observation(self, picking):
        value = self._find_text_field(
            picking,
            technical_candidates=(
                "observacion",
                "observaciones",
                "observation",
                "observations",
                "x_observacion",
                "x_observaciones",
                "x_studio_observacion",
                "x_studio_observaciones",
                "note",
            ),
            label_candidates=(
                "Observación",
                "Observaciones",
            ),
        )
        return self._plain_text(value)

    def _get_purchase_order(self, picking):
        if "purchase_id" in picking._fields and picking.purchase_id:
            return picking.purchase_id

        move_field = "move_lines" if "move_lines" in picking._fields else "move_ids_without_package"
        moves = getattr(picking, move_field, self.env["stock.move"])
        if moves and "purchase_line_id" in moves._fields:
            orders = moves.mapped("purchase_line_id.order_id")
            if len(orders) == 1:
                return orders

        if picking.origin:
            order = self.env["purchase.order"].search([
                ("name", "=", picking.origin),
                ("company_id", "=", picking.company_id.id),
            ], limit=1)
            if order:
                return order
        return self.env["purchase.order"]

    @staticmethod
    def _get_float(record, field_names, default=0.0):
        for name in field_names:
            if name in record._fields:
                try:
                    return float(getattr(record, name, 0.0) or 0.0)
                except (TypeError, ValueError):
                    continue
        return default

    def _get_product_destination(self, product):
        """Destino = tipo/variante de producto usado actualmente por la operación."""
        values = product.product_template_variant_value_ids.mapped("name")
        if values:
            return ", ".join(values)

        template = product.product_tmpl_id
        value = self._find_text_field(
            product,
            technical_candidates=(
                "tipo_producto",
                "type_product",
                "product_type_custom",
                "destino",
                "destination",
            ),
            label_candidates=("Tipo de producto", "Destino"),
        )
        if value:
            return self._plain_text(value)

        value = self._find_text_field(
            template,
            technical_candidates=(
                "tipo_producto",
                "type_product",
                "product_type_custom",
                "destino",
                "destination",
            ),
            label_candidates=("Tipo de producto", "Destino"),
        )
        return self._plain_text(value)

    def _get_factor_pulgada_base(self, product):
        tag = self._get_product_destination(product)
        largo = self._get_float(product, ("largo", "long", "length"), 0.0)

        registry_model = self.env.registry.get("stock.mantenedor_cubicacion")
        if registry_model and tag:
            Mantenedor = self.env["stock.mantenedor_cubicacion"]
            domain = [("nombre", "=", tag.upper())]
            if "largo" in Mantenedor._fields and largo:
                domain.append(("largo", "=", largo))
            mantenedor = Mantenedor.search(domain, limit=1)
            if mantenedor and "factor" in mantenedor._fields:
                try:
                    return float(mantenedor.factor or 0.0)
                except (TypeError, ValueError):
                    pass

        # Respaldo proporcional al largo de referencia de 3,20 m.
        if largo:
            return largo / 3.2
        return 1.0

    def _compute_factor_and_inches(self, product, quantity):
        quantity = float(quantity or 0.0)
        espesor = self._get_float(product, ("espesor", "thickness"), 0.0)
        ancho = self._get_float(product, ("ancho", "width"), 0.0)
        largo = self._get_float(product, ("largo", "long", "length"), 0.0)

        is_mm = False
        for field_name in ("isMmProducto", "isMmProduct"):
            if field_name in product._fields:
                is_mm = bool(getattr(product, field_name))
                break

        if is_mm:
            factor = ((espesor / 1000.0) * (ancho / 1000.0) * largo) * 48.43
        else:
            factor_base = self._get_factor_pulgada_base(product)
            factor = (factor_base * espesor * ancho) / 10.0

        return factor, factor * quantity

    def _get_purchase_line(self, move, purchase, product):
        if move and "purchase_line_id" in move._fields and move.purchase_line_id:
            return move.purchase_line_id
        if purchase:
            lines = purchase.order_line.filtered(lambda line: line.product_id == product)
            if lines:
                return lines[0]
        return self.env["purchase.order.line"]

    def _package_rows_from_move_lines(self, picking, purchase):
        """Una fila por paquete/producto recibido usando qty_done de stock.move.line."""
        groups = {}
        move_lines = picking.move_line_ids if "move_line_ids" in picking._fields else self.env["stock.move.line"]

        for line in move_lines:
            qty = float(line.qty_done or 0.0)
            if qty <= 0:
                continue
            package = line.result_package_id if "result_package_id" in line._fields else self.env["stock.quant.package"]
            if not package:
                continue
            product = line.product_id
            purchase_line = self._get_purchase_line(line.move_id, purchase, product)
            key = (package.id, product.id, purchase_line.id or 0)
            if key not in groups:
                groups[key] = {
                    "package": package,
                    "product": product,
                    "quantity": 0.0,
                    "purchase_line": purchase_line,
                }
            groups[key]["quantity"] += qty

        return list(groups.values())

    def _package_rows_fallback(self, picking, purchase):
        """Respaldo para instalaciones que gestionan paquetes desde package_level_ids_details."""
        rows = []
        if "package_level_ids_details" not in picking._fields:
            return rows

        for level in picking.package_level_ids_details:
            package = level.package_id
            if not package:
                continue
            quants = package.quant_ids.filtered(lambda quant: quant.quantity > 0)
            for product in quants.mapped("product_id"):
                quantity = sum(quants.filtered(lambda q: q.product_id == product).mapped("quantity"))
                purchase_line = self._get_purchase_line(False, purchase, product)
                rows.append({
                    "package": package,
                    "product": product,
                    "quantity": quantity,
                    "purchase_line": purchase_line,
                })
        return rows

    def _get_report_rows(self):
        self.ensure_one()
        start_utc, end_utc = self._date_bounds_utc()

        domain = [
            ("state", "=", "done"),
            ("picking_type_id.code", "=", "incoming"),
            ("date_done", ">=", start_utc),
            ("date_done", "<=", end_utc),
            ("company_id", "=", self.company_id.id),
        ]
        if self.partner_id:
            domain.append(("partner_id", "=", self.partner_id.id))

        pickings = self.env["stock.picking"].search(domain, order="date_done, name")
        rows = []

        for picking in pickings:
            if "returned" in picking._fields and picking.returned:
                continue

            purchase = self._get_purchase_order(picking)
            # Evita incluir devoluciones de clientes u otros ingresos no vinculados a compras.
            if not purchase:
                continue

            observation = self._get_observation(picking)
            package_rows = self._package_rows_from_move_lines(picking, purchase)
            if not package_rows:
                package_rows = self._package_rows_fallback(picking, purchase)

            local_datetime = fields.Datetime.context_timestamp(self, picking.date_done) if picking.date_done else False

            for package_data in package_rows:
                product = package_data["product"]
                quantity = package_data["quantity"]
                purchase_line = package_data["purchase_line"]
                price = float(purchase_line.price_unit or 0.0) if purchase_line else 0.0
                factor, inches = self._compute_factor_and_inches(product, quantity)
                total = quantity * price
                price_per_inch = total / inches if inches else 0.0

                rows.append({
                    "supplier": (purchase.partner_id or picking.partner_id).display_name,
                    "purchase_order": purchase.name or "",
                    "reception": picking.name or "",
                    "document": observation,
                    "package": package_data["package"].name or "",
                    "description": product.display_name or product.name or "",
                    "quantity": quantity,
                    "factor": factor,
                    "inches": inches,
                    "price": price,
                    "date": local_datetime.replace(tzinfo=None) if local_datetime else False,
                    "destination": self._get_product_destination(product),
                    "total": total,
                    "price_per_inch": price_per_inch,
                })

        return rows

    def action_generate_excel(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise ValidationError(_("La fecha desde no puede ser posterior a la fecha hasta."))

        rows = self._get_report_rows()
        if not rows:
            raise UserError(_(
                "No se encontraron paquetes recibidos desde órdenes de compra en el rango seleccionado."
            ))

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        worksheet = workbook.add_worksheet("Recepciones")

        title_fmt = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "align": "center",
            "valign": "vcenter",
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E78",
        })
        subtitle_fmt = workbook.add_format({
            "italic": True,
            "align": "center",
            "font_color": "#404040",
        })
        header_fmt = workbook.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#4472C4",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })
        text_fmt = workbook.add_format({"border": 1, "valign": "top"})
        center_fmt = workbook.add_format({"border": 1, "align": "center", "valign": "top"})
        integer_fmt = workbook.add_format({"border": 1, "num_format": "#,##0", "align": "right"})
        decimal_fmt = workbook.add_format({"border": 1, "num_format": "#,##0.000", "align": "right"})
        money_fmt = workbook.add_format({"border": 1, "num_format": "$#,##0", "align": "right"})
        money_decimal_fmt = workbook.add_format({"border": 1, "num_format": "$#,##0.00", "align": "right"})
        date_fmt = workbook.add_format({"border": 1, "num_format": "dd/mm/yyyy", "align": "center"})
        total_label_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        total_num_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "num_format": "#,##0.000"})
        total_money_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "num_format": "$#,##0"})

        headers = [
            "Proveedor",
            "Pedido de compra",
            "Recepción",
            "Guía / Factura (Observación)",
            "Código paquete",
            "Descripción",
            "Pzas",
            "Factor",
            "Pulgadas",
            "Precio",
            "Fecha",
            "Destino",
            "Total $",
            "$/pulg",
        ]

        worksheet.merge_range(0, 0, 0, len(headers) - 1, "INFORME DE RECEPCIONES DE COMPRA", title_fmt)
        worksheet.merge_range(
            1, 0, 1, len(headers) - 1,
            "Período: %s al %s" % (
                fields.Date.to_string(self.date_from),
                fields.Date.to_string(self.date_to),
            ),
            subtitle_fmt,
        )

        header_row = 3
        for col, header in enumerate(headers):
            worksheet.write(header_row, col, header, header_fmt)

        data_start = header_row + 1
        for row_idx, row in enumerate(rows, start=data_start):
            worksheet.write(row_idx, 0, row["supplier"], text_fmt)
            worksheet.write(row_idx, 1, row["purchase_order"], center_fmt)
            worksheet.write(row_idx, 2, row["reception"], center_fmt)
            worksheet.write(row_idx, 3, row["document"], text_fmt)
            worksheet.write(row_idx, 4, row["package"], center_fmt)
            worksheet.write(row_idx, 5, row["description"], text_fmt)
            worksheet.write_number(row_idx, 6, row["quantity"], integer_fmt)
            worksheet.write_number(row_idx, 7, row["factor"], decimal_fmt)
            worksheet.write_number(row_idx, 8, row["inches"], decimal_fmt)
            worksheet.write_number(row_idx, 9, row["price"], money_fmt)
            if row["date"]:
                worksheet.write_datetime(row_idx, 10, row["date"], date_fmt)
            else:
                worksheet.write(row_idx, 10, "", center_fmt)
            worksheet.write(row_idx, 11, row["destination"], text_fmt)
            worksheet.write_number(row_idx, 12, row["total"], money_fmt)
            worksheet.write_number(row_idx, 13, row["price_per_inch"], money_decimal_fmt)

        last_data_row = data_start + len(rows) - 1
        total_row = last_data_row + 1
        worksheet.merge_range(total_row, 0, total_row, 5, "TOTALES", total_label_fmt)
        worksheet.write_formula(total_row, 6, "=SUM(G%d:G%d)" % (data_start + 1, last_data_row + 1), total_label_fmt)
        worksheet.write(total_row, 7, "", total_label_fmt)
        worksheet.write_formula(total_row, 8, "=SUM(I%d:I%d)" % (data_start + 1, last_data_row + 1), total_num_fmt)
        worksheet.write(total_row, 9, "", total_label_fmt)
        worksheet.write(total_row, 10, "", total_label_fmt)
        worksheet.write(total_row, 11, "", total_label_fmt)
        worksheet.write_formula(total_row, 12, "=SUM(M%d:M%d)" % (data_start + 1, last_data_row + 1), total_money_fmt)
        worksheet.write_formula(
            total_row,
            13,
            "=IF(I%d=0,0,M%d/I%d)" % (total_row + 1, total_row + 1, total_row + 1),
            money_decimal_fmt,
        )

        worksheet.autofilter(header_row, 0, last_data_row, len(headers) - 1)
        worksheet.freeze_panes(data_start, 0)
        worksheet.set_row(0, 26)
        worksheet.set_row(header_row, 34)
        widths = [22, 17, 17, 30, 19, 42, 10, 11, 13, 14, 12, 22, 16, 14]
        for col, width in enumerate(widths):
            worksheet.set_column(col, col, width)

        workbook.close()
        output.seek(0)

        filename = "recepciones_%s_%s.xlsx" % (
            fields.Date.to_string(self.date_from).replace("-", ""),
            fields.Date.to_string(self.date_to).replace("-", ""),
        )
        self.write({
            "file_data": base64.b64encode(output.read()),
            "file_name": filename,
        })

        return {
            "type": "ir.actions.act_url",
            "url": "/web/content?model=%s&id=%s&field=file_data&filename_field=file_name&download=true" % (
                self._name,
                self.id,
            ),
            "target": "self",
        }
