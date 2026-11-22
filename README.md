# Plantilles fulls de càlcul (multi-format)

Plantilles amb fórmules per a LibreOffice / OpenOffice Calc, WPS Office Sheet i Excel.
Organitzades per exercicis i exportades a diversos formats.

## Estructura

```
exercici1/          # Checklist + Factura
  checklist.ods / .xlsx / .csv
  factura.ods / .xlsx / .pdf

exercici3/          # Notes del professorat
  quadern_notes_professorat.xlsx / .ods
  llista_notes_finals.xlsx / .ods / .csv / .pdf
```

## Contingut

### Exercici 1
- **Checklist**: recompte Sí/No amb `COUNTA` i `COUNTIF`
- **Factura**: subtotals, descomptes, IVA, totals amb `SUM` / `SUMIF`

### Exercici 3
- **Quadern de notes**: pesos configurables, mitjana ponderada (`SUMPRODUCT`), qualificacions, estadístiques + gràfic
- **Llista notes finals**: 1a/2a/3a avaluació, nota final, mitjana, qualificació

## Formats disponibles
| Format | Ús |
|--------|-----|
| `.ods` | LibreOffice / OpenOffice Calc |
| `.xlsx` | WPS Office Sheet, Excel, LibreOffice |
| `.csv`  | Intercanvi de dades |
| `.pdf`  | Visualització / entrega |

## Scripts generadors
- `create_templates.py` → checklist + factura (ODS)
- `create_notes_templates.py` → quadern + llista notes (XLSX)

```bash
pip install odfpy openpyxl
python3 create_templates.py
python3 create_notes_templates.py
```

## Repositori
https://github.com/RamisJoanSigma/plantilles-calc-ods
