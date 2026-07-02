# Thinkswiss

## Peru — HRNEC 2030 (National Circular Economy Roadmap)

National Circular Economy Roadmap to 2030 document (*Hoja de Ruta Nacional de Economía Circular al 2030*), coded for the 4P Index. Approved by DS 003-2025-MINAM.

### Downloads

- **4P Index coding (Excel):** [Peru_HRNEC_2030_4P_Index_Coding.xlsx](https://github.com/shivrain/Thinkswiss/raw/cursor/peru-hrnec-2030-4p-index-d350/output/peru-hrnec-2030/Peru_HRNEC_2030_4P_Index_Coding.xlsx)
- **English translation (Word):** [Peru_HRNEC_2030_English_Translation.docx](https://github.com/shivrain/Thinkswiss/raw/cursor/peru-hrnec-2030-4p-index-d350/output/peru-hrnec-2030/Peru_HRNEC_2030_English_Translation.docx)

### Source

- MINAM (PDF): https://cdn.www.gob.pe/uploads/document/file/9548248/6507211-hoja-de-ruta-nacional-de-economia-circular-al-2030.pdf
- Approving decree: https://www.gob.pe/institucion/minam/normas-legales/6507211-003-2025-minam

### Regenerate

```bash
python3 scripts/generate_peru_hrnec_2030_deliverables.py
```

### Format

- Columns A–W per 4P Index coding prompt v2 (bilingual policy name in column A; no separate English or country columns)
- Bilingual cells: Spanish original / English translation in the same cell
