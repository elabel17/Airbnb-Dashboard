# Finanzas Personales (versión simple) · Excel 365

`Finanzas_Personales.xlsx` funciona sin macros y usa DOP como moneda base.

## Registrar un movimiento

En **Registro** llenas 4 datos: **Fecha, Monto, Categoría y Cuenta**. Hay dos columnas opcionales:

- **Comercio**: alimenta el ranking de lugares.
- **Hacia**: solo para transferencias y pagos de tarjeta.

El resto se calcula a partir de la categoría y la cuenta: tipo de movimiento, grupo, fijo o variable, moneda, análisis de comida y gastos hormiga.

| Situación | Cómo se registra |
|---|---|
| Compra con tarjeta de crédito | Gasto normal con Cuenta = «Tarjeta de crédito». |
| Pago de la tarjeta | Categoría «Pago de tarjeta», Hacia = la tarjeta. No cuenta como gasto. |
| Mover dinero entre tus cuentas | «Transferencia entre cuentas» con Hacia. No cuenta como ingreso ni como gasto. |
| Depositar o retirar de trading | «Depósito / retiro de trading» con Hacia. No cuenta como gasto. |
| Ganancia o pérdida de trading | «Ganancia / pérdida de trading», con monto + o −. Se reporta aparte. |
| Devolución o reembolso | Monto en negativo con la misma categoría del gasto. |
| Cuenta en dólares | Escribe el monto en USD. Se convierte con la tasa de Ajustes. |

## Hojas

- **Registro**: tus movimientos. La columna *Estado* avisa si falta algo.
- **Resumen**: tablero con período, ingresos, gasto, ahorro, gasto de hoy, semana y mes, gasto por grupo, evolución, métodos de pago, lugares, comida comprada frente a cocinar, gastos hormiga, trading y flujo de efectivo.
- **Metas**: límites mensuales por grupo, objetivos de ahorro y agenda de suscripciones. Todo empieza vacío.
- **Ajustes**: tasa USD, umbral de gasto hormiga, cuentas con saldo inicial opcional y categorías editables.

## Datos DEMO

Las filas lilas del Registro (B6:G106) y las suscripciones «DEMO» de Metas son ficticias. Para empezar con tus datos, selecciónalas y pulsa Suprimir.

## Regenerar

```bash
python build_finanzas.py Finanzas_Personales.xlsx   # requiere openpyxl
```
