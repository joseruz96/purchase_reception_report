# Informe Excel de Recepciones de Compra - Odoo 15

Módulo independiente para descargar recepciones de proveedor por rango de fechas.

## Columnas del Excel

- Proveedor
- Pedido de compra
- Recepción
- Guía / Factura (campo Observación de la recepción)
- Código paquete
- Descripción
- Pzas
- Factor
- Pulgadas
- Precio
- Fecha
- Destino (tipo/variante del producto)
- Total $
- $/pulg

## Fuente de datos

- Recepción: `stock.picking` realizada (`done`) de tipo `incoming`.
- Orden de compra: `stock.picking.purchase_id`, líneas de compra o `origin` como respaldo.
- Paquete: `stock.move.line.result_package_id`, con `package_level_ids_details` como respaldo.
- Cantidad: `stock.move.line.qty_done`.
- Documento: campo Observación/Observaciones de `stock.picking`; se detecta por nombre técnico o etiqueta.
- Destino: valores de variante `product_template_variant_value_ids`; se contemplan campos personalizados como respaldo.
- Precio: `purchase.order.line.price_unit`.
- Factor/Pulgadas: reutiliza la lógica operacional existente y `stock.mantenedor_cubicacion` cuando está disponible.

## Uso

Compras -> Informe de Recepciones

Seleccione Fecha desde / Fecha hasta y presione **Descargar Excel**.
