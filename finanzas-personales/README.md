# Finanzas Personales: sistema de control en Excel 365

`Finanzas_Personales.xlsx` es un sistema de registro y análisis financiero personal. Funciona sin macros ni dependencias externas, con moneda base DOP y soporte para USD.

## Hojas

| Hoja | Para qué sirve |
|---|---|
| **Inicio** | Navegación, estado del sistema, reglas de registro y cómo borrar los datos DEMO. |
| **Registro** | Tabla central. Cada fila es una transacción. Tiene listas desplegables, validaciones y una columna *Estado* que avisa de errores. |
| **Dashboard** | KPIs, filtros (período, mes, rango personalizado, categoría y cuenta), gráficos, top 10, comparación mensual y tres vistas separadas: flujo de efectivo, gasto real y trading. |
| **Alimentación** | Comida comprada frente a súper e ingredientes, costo por desayuno, almuerzo y cena, frecuencia, días de mayor gasto y un simulador de ahorro. |
| **Gastos hormiga** | Gastos variables por debajo de un umbral editable: cantidad, frecuencia, acumulado, peso sobre el total y proyecciones. |
| **Recurrentes** | Agenda de suscripciones con equivalente mensual, próximo cobro y último cargo registrado. No suma en los reportes. |
| **Presupuestos** | Etapa 2: límites diarios, semanales, mensuales y por categoría, con alertas. También objetivos (fondo de emergencia, boda, inversión). Todo empieza vacío. |
| **Cuentas** | Saldos iniciales opcionales, saldo actual, deuda de tarjetas, conciliación y gasto por método. |
| **Categorías / Config** | Listas editables y parámetros (tasa USD→DOP, cuenta predeterminada). |
| **Motor** | Cálculos intermedios para rankings y series. No se edita. |

## Reglas contables incorporadas

- **Compra con tarjeta de crédito**: se registra como *Gasto* en la fecha de consumo, con la tarjeta como cuenta.
- **Pago de la tarjeta**: se registra como *Pago tarjeta* (banco → tarjeta). Reduce la deuda y no cuenta como un segundo gasto.
- **Transferencias entre cuentas propias**: no cuentan como ingreso ni como gasto.
- **Trading**: los depósitos y retiros son transferencias. Las ganancias y pérdidas realizadas son *Resultado trading* (+/−). Las evaluaciones, plataformas y comisiones son gasto real y además se reportan como costo operativo de trading.
- **Reembolsos**: restan del gasto de su categoría.
- **Saldos**: solo se calculan si se escribe un saldo inicial real. La hoja no inventa saldos.

## Datos de demostración

El Registro incluye 189 filas ficticias (agosto a octubre de 2026). Están marcadas con «DEMOSTRACIÓN» y fondo lila. Recurrentes incluye 5 más. Para borrarlas, selecciona `B6:N194` en Registro y pulsa *Suprimir*. Las fórmulas se mantienen.

## Regenerar el archivo

```bash
python build_finanzas.py Finanzas_Personales.xlsx
```

El script requiere `openpyxl`. Al abrir el archivo, Excel recalcula todas las fórmulas.
