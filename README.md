# Plantilles LibreOffice Calc (ODS)

Plantilles creades amb fórmules per a:

1. **Checklist** (`checklist.ods`) – Llista de tasques amb recompte automàtic
2. **Factura** (`factura.ods`) – Factura completa amb càlculs d'IVA i descomptes

## Checklist

- Columna de descripció de tasques
- Columna **Fet?** amb valors **Sí** / **No** (utilitza validació de dades tipus llista)
- Recomptes automàtics amb fórmules:
  - **Total elements**: `=COUNTA(B5:B24)`
  - **Fets (Sí)**: `=COUNTIF(C5:C24;"Sí")`
  - **Pendents (No)**: `=COUNTIF(C5:C24;"No")`

### Com activar la llista desplegable Sí/No

1. Selecciona el rang `C5:C24`
2. Menú **Dades → Validesa**
3. Criteri: **Llista**
4. Entrades: `Sí;No`
5. Accepta

## Factura

Basada en l'estructura de la factura d'exemple. Inclou:

| Fórmula | Descripció |
|---------|------------|
| `[1] Subtotal` | `=Unitats × Preu Unitari` |
| `[2] Total descompte` | `=Subtotal × % Descompte` |
| `[3] Total IVA` | `=(Subtotal − Descompte) × % IVA` |
| `[4] Total amb IVA` | `=Subtotal − Descompte + IVA` |
| `[5] Import brut` | `=SUM(Subtotals)` |
| `[6] Total descomptes` | `=SUM(Total descompte)` |
| `[7] Base Imponible` | `SUMIF` per cada tipus d'IVA |
| `[8] Import IVA` | `=Base × Tipus IVA` |

El **TOTAL FACTURA** suma la columna de totals amb IVA.

## Com generar les plantilles

```bash
python3 create_templates.py
```

Requereix `odfpy` (`pip install odfpy`).

## Fitxers

- `checklist.ods` / `checklist.xlsx`
- `factura.ods` / `factura.xlsx`
- `create_templates.py` – script generador

Obre els fitxers `.ods` amb **LibreOffice Calc**.
