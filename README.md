# Generador de Liquidaciones y Acuerdos
### Fisco de la Provincia de Buenos Aires — Fuero Comercial de la Nación

Herramienta web para generar liquidaciones judiciales y acuerdos de pago en apremios fiscales del Fuero Comercial.

---

## Funcionalidades

### Liquidación
Calcula automáticamente todos los rubros de costas:

| Rubro | Detalle |
|-------|---------|
| Tasa de justicia | 3 % sobre el monto |
| Sobre tasa | **5 % o 10 %** según el caso |
| Honorarios | % elegido sobre el monto, con **mínimo aplicable** |
| Honorarios Fiscalía | 40 % de honorarios |
| Honorarios apoderado/a | 60 % de honorarios |
| Aportes previsionales | 12 % sobre honorarios |
| Gastos de juicio | $ 24.875 / $ 49.750 / $ 33.998 (automático) |
| Servicios registrales | Opcional: $ 84.000 / $ 168.000 / $ 252.000 |

### Honorarios mínimos
- **$ 149.250** → gastos se eligen libremente ($ 24.875 o $ 49.750)
- **$ 101.994** → gastos se fijan automáticamente en **$ 33.998**

### Acuerdo de pago
Genera el texto completo del acuerdo con:
- Datos del contribuyente (DNI / CUIT)
- Apoderado/a fiscal (Natalia, Mauricio o Maximina)
- Juzgado N° 1 o N° 2
- Hasta 6 títulos ejecutivos
- Montos de la liquidación incorporados al texto si se calculó previamente

---

## Estructura del proyecto

```
GeneradorAcuerdos/
├── index.html          # Página principal
├── css/
│   └── style.css       # Estilos
├── js/
│   └── liquidacion.js  # Lógica de cálculo y generación
└── README.md
```

---

## Uso

Clonar o descargar el repositorio y abrir `index.html` en un navegador. No requiere servidor ni dependencias externas (solo carga la fuente Rubik desde Google Fonts).

```bash
git clone https://github.com/TU_USUARIO/GeneradorAcuerdos.git
cd GeneradorAcuerdos
# Abrir index.html en el navegador
```

---

## Flujo recomendado

1. Completar la **pestaña Liquidación** con el monto, la sobre tasa, el porcentaje y mínimo de honorarios, los gastos y servicios registrales.
2. Hacer clic en **Calcular liquidación** → revisar el desglose.
3. Ir a la **pestaña Acuerdo de pago** → completar los datos del contribuyente.
4. Hacer clic en **Generar acuerdo** → el texto se completa automáticamente con los montos calculados.
5. **Imprimir** o copiar el texto para el expediente.
