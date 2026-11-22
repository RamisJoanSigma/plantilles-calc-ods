# Plantilles fulls de càlcul (ODS + XLSX)

Plantilles amb fórmules per a LibreOffice Calc, WPS Office Sheet i Excel.

## Contingut

### 1. Checklist (`checklist.ods`)
- Llista de tasques amb columna **Sí/No**
- Recomptes automàtics: total, fets, pendents (`COUNTA`, `COUNTIF`)

### 2. Factura (`factura.ods`)
- Factura completa amb subtotals, descomptes, IVA i total
- Fórmules per línia i resum per tipus d'IVA (`SUM`, `SUMIF`)

### 3. Quadern de notes del professorat (`quadern_notes_professorat.xlsx`)
- Full **Notes**: alumnes, 8 activitats amb pesos configurables, mitjana ponderada, nota final, qualificació textual
- Full **Estadístiques**: distribució de qualificacions + gràfic
- Full **Instruccions**
- Fórmules: `SUMPRODUCT`, `SUMIF`, `ROUND`, `COUNTIF`, `AVERAGEIF`, `MAX`, `MIN`
- Format condicional (escala de color a les notes)

### 4. Llista de notes finals (`llista_notes_finals.xlsx`)
Columnes:
| Nom i cognoms | Correu | Nota 1a aval. | Nota 2a aval. | Nota 3a aval. | Nota final | Mitjana | Qualificació |

- Nota final = mitjana de les 3 avaluacions
- Resum del grup (mitjana, aprovats, suspesos, %…)

## Com generar

```bash
# ODS (checklist + factura)
pip install odfpy
python3 create_templates.py

# XLSX (quadern + llista notes)
python3 create_notes_templates.py
```

## Repositori
https://github.com/RamisJoanSigma/plantilles-calc-ods
