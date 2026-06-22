#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Decree 97/2020."""

import os
from datetime import date

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_DIR = "/workspace/output/mozambique"
os.makedirs(OUTPUT_DIR, exist_ok=True)

POLICY_FIELDS = {
    "policy_name": (
        "Regulamento de Gestão e Ordenamento da Zona Costeira e das Praias "
        "(Decreto n.º 97/2020)"
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/moz213251.pdf",
    "policy_year": 2020,
    "policy_objective": (
        "Define principles and rules for sustainable integrated management and planning "
        "of coastal zones and beaches; conserve sensitive ecosystems; prevent and combat "
        "marine and coastal pollution; ensure urban solid waste is not discharged to "
        "beaches, sea or watercourses; and regulate beach use, concessions and coastal activities."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (Regulamento) approved by Council of Ministers Decree n.º 97/2020 "
        "under Sea Law (Lei n.º 20/2019) Arts. 22 and 96 — sub-legislative instrument revoking "
        "Decree 45/2006 Arts. 53–61 on beach management."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "fisheries, waste management, industry, tourism, municipalities, water, packaging"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "consumption, disposal, environmental leakage"
    ),
    "policy_budget": 0.50,
    "policy_budget_text": (
        "Art. 4: polluter-pays and user-pays principles for coastal zone management; "
        "Annex II: fines measured in minimum public-sector salary units (2–24 salaries); "
        "no dedicated plastic pollution budget allocation in decree text."
    ),
    "policy_score": "[auto]",
}

# Six consolidated instruments — beach classification chapters not overcoded line-by-line.
INSTRUMENTS = [
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 4(d–f, k): Establishes environmental responsibility — polluters must restore or "
            "compensate natural resource damage; zero-waste principle requiring full valorization "
            "of waste produced or found in coastal zones and beaches; polluter-pays and user-pays "
            "principles for coastal management decisions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 4f: "desperdício zero" — full valorization of resíduos in coastal zone and beaches; '
            'Art. 4k: "princípio do poluidor – pagador" — obligation to repair or compensate '
            'environmental damage; Art. 4d: prevention systems for acts harmful to the environment.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Strategic principles in regulation — operationalised through Arts. 35, 50 and Annex II. "
            "Not coded separately: co-responsibilization, civic education (Art. 4e)."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 34, Art. 50(1b), Art. 10(e): Prohibits dumping, abandoning, burying or burning "
            "any solid or liquid waste in coastal zones and beaches; requires elimination of "
            "pollution sources; beach administrators must ensure urban solid waste is not "
            "discharged to beaches, sea, watercourses or hazardous locations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 50(1b): ban on "lançar, abandonar, despejar, enterrar ou queimar qualquer tipo '
            'de resíduos, sólidos ou líquidos"; Art. 10(e): urban waste must not be "lançados em '
            'praias, no mar, e em cursos e corpos de água"; Annex II #6: 3 minimum salaries fine.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary marine litter prevention instrument — direct beach/sea discharge ban. "
            "Revokes Decree 45/2006 beach provisions (Decree Art. 4). Not coded separately: "
            "Art. 34 education/monitoring activities."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 50(1h): Prohibits use of glass and plastic packaging in bathing zones "
            "(zonas balneares), except in duly licensed catering establishments."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 50(1h): "o uso de embalagens de vidro e plásticos nas zonas balneares, com '
            'excepção dos estabelecimentos de restauração devidamente licenciados"; Annex II #4: '
            '3 minimum salaries and seizure of glass packaging (plastic covered by Art. 50.1h).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Explicit plastic packaging restriction in bathing areas — one of only two direct "
            "plastic mentions in the regulation. Complements Decree 16/2015 plastic bag control."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 35–36, Art. 11(d), Art. 10(d): Beach users must collect and deposit solid waste "
            "in containers or carry it to nearest receptacle; economic operators responsible for "
            "waste from their activities; mandatory beach waste collection and cleaning services; "
            "beach sand cleaning and container emptying assigned to concessionaires or maritime authority."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 35(1): users must deposit resíduos in "contentores, ecopontos e baldes"; '
            'Art. 35(2): operators responsible for activity waste; Art. 11(1d): mandatory '
            '"recolha de resíduos e de limpeza de praia"; Art. 36: concessionaires clean '
            'concession areas, authority cleans remaining beach areas.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Operational waste management on beaches — links to Decree 94/2014 urban waste rules. "
            "Not coded separately: mandatory sanitary facilities (Art. 11e–f)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 47: Fishers and aquaculturists in coastal zones must correctly manage waste "
            "with special focus on all types of plastics including fishing gear — ropes, nets, "
            "buoys, lines, hooks, fishing equipment, vessel debris — and keep fishing camps and "
            "mariculture infrastructure clean."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 47(1a): "gerir correctamente os resíduos, com especial enfoque para todo o tipo '
            'de plásticos, incluindo... cordas, redes, boiás, linhas, anzóis, equipamentos de pesca, '
            'restos e ou destroços de embarcações"; Art. 47(2): sanctionable under Annex II.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Explicit plastic/fishing gear waste obligation — addresses ghost gear and coastal "
            "fishing camp litter. Annex II #16: 2–6 minimum salaries for operator non-compliance."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 48, Art. 51–53, Annex II: Coordinated coastal/beach fiscalization to prevent "
            "harmful activities; violations sanctioned with fines (minimum salary units), mandatory "
            "restoration/indemnification, accessory penalties including equipment seizure, licence "
            "suspension and infrastructure reversion to the State."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 51(1): fines plus "medidas de recuperação ou de indemnização obrigatória"; '
            'Art. 52(a): "apreensão de equipamentos e produtos"; Annex II #6 waste dumping: '
            '3 salaries; #7 illegal dumps in sensitive ecosystems: 12 salaries; 20-day payment deadline.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Enforcement layer for Arts. 35, 47, 50 prohibitions. Not coded separately: "
            "community fiscalization agents (Art. 48.3), alcohol/noise infractions."
        ),
    },
]

COLUMNS = [
    "policy_name", "policy_url", "policy_year", "policy_objective", "policy_target",
    "policy_target_text", "policy_type", "policy_type_justification", "policy_integration",
    "policy_sectors_list", "policy_circularity", "policy_lifecycle_phases_list",
    "policy_budget", "policy_budget_text", "policy_score",
    "instrument_type", "instrument_lifecycle_stage", "instrument_description",
    "instrument_in_force", "instrument_implementation", "instrument_implementation_text",
    "instrument_score", "comments",
]

COLUMN_LABELS = {
    "policy_name": "A — policy_name",
    "policy_url": "B — policy_url",
    "policy_year": "C — policy_year",
    "policy_objective": "D — policy_objective",
    "policy_target": "E — policy_target",
    "policy_target_text": "F — policy_target_text",
    "policy_type": "G — policy_type",
    "policy_type_justification": "H — policy_type_justification",
    "policy_integration": "I — policy_integration",
    "policy_sectors_list": "J — policy_sectors_list",
    "policy_circularity": "K — policy_circularity",
    "policy_lifecycle_phases_list": "L — policy_lifecycle_phases_list",
    "policy_budget": "M — policy_budget",
    "policy_budget_text": "N — policy_budget_text",
    "policy_score": "O — policy_score [auto]",
    "instrument_type": "P — instrument_type",
    "instrument_lifecycle_stage": "Q — instrument_lifecycle_stage",
    "instrument_description": "R — instrument_description",
    "instrument_in_force": "S — instrument_in_force",
    "instrument_implementation": "T — instrument_implementation",
    "instrument_implementation_text": "U — instrument_implementation_text",
    "instrument_score": "V — instrument_score [auto]",
    "comments": "W — comments",
}


def build_dataframe():
    return pd.DataFrame([{**POLICY_FIELDS, **i} for i in INSTRUMENTS], columns=COLUMNS)


def export_csv(df, basename):
    path = os.path.join(OUTPUT_DIR, f"{basename}.csv")
    df.rename(columns=COLUMN_LABELS).to_csv(path, index=False, encoding="utf-8-sig")
    return path


def export_xlsx(df, basename):
    path = os.path.join(OUTPUT_DIR, f"{basename}.xlsx")
    labelled = df.rename(columns=COLUMN_LABELS)
    labelled.to_excel(path, index=False, sheet_name="4P Index Coding")
    wb = load_workbook(path)
    ws = wb.active
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = min(45, max(12, len(str(ws.cell(1, col).value or "")) * 0.9))
    ws.freeze_panes = "A2"
    wb.save(path)
    return path


def add_heading(doc, text, level=1):
    doc.add_heading(text, level=level)


def add_para(doc, text):
    doc.add_paragraph(text)


def create_translation_doc():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading(
        "Regulation on Management and Planning of Coastal Zones and Beaches", 0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Decree No. 97/2020 of 4 October")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Regulamento de Gestão e Ordenamento da Zona Costeira "
        "e das Praias (Decreto n.º 97/2020)\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run("Source: https://faolex.fao.org/docs/pdf/moz213251.pdf")

    sections = [
        ("Decree No. 97/2020 — Enacting Provisions", (
            "Under combined provisions of Sea Law (Law 20/2019) Articles 22 and 96, the Council of "
            "Ministers decrees:\n\n"
            "Article 1. The Regulation on Management and Planning of Coastal Zones and Beaches is "
            "approved as an annex forming an integral part of this Decree.\n\n"
            "Article 2. The Minister responsible for the maritime sector shall adopt procedures for "
            "correct management and planning of coastal zones and beaches.\n\n"
            "Article 3. The Minister may delegate specific matters to State representation bodies "
            "and decentralized entities.\n\n"
            "Article 4. Articles 53–61 of Decree 45/2006 (marine pollution regulation) on beach "
            "management are revoked.\n\n"
            "Article 5. This Decree enters into force on the date of publication."
        )),
        ("Chapter I — General Provisions", [
            ("Object and scope (Arts. 2–3)", (
                "Article 2 (Object): The Regulation defines principles and rules for sustainable "
                "integrated management and planning of coastal zones and beaches, including:\n"
                "a) guarantee of public use of the maritime-terrestrial public domain;\n"
                "b) protection of coastline, beaches, dunes, native vegetation, mangroves, wetlands "
                "and seagrass beds;\n"
                "c) safety of bathers and tourists;\n"
                "d) water and sand quality and biodiversity health;\n"
                "e) private sector, NGO and community participation rules;\n"
                "f) concession criteria; g) use compatibility; h) beach qualification;\n"
                "i) bathing beach classification; j) risk-zone safeguards.\n\n"
                "Article 3 (Scope): Applies to maritime public domain within 100 metres inland from "
                "the high-water line, and to all natural and legal persons using coastal zones."
            )),
            ("Principles (Art. 4)", (
                "Key principles include: sustainability; integrated territorial management; "
                "co-responsibility; environmental responsibility (polluters restore/compensate); "
                "civic and environmental education; zero waste (full valorization of coastal waste); "
                "rational environmental component use; social equity; prevention and precaution; "
                "polluter-pays and user-pays; coordinated scientific approach."
            )),
            ("Policy instruments (Art. 7)", (
                "POEM, POLMAR, EGIZC, Mangrove Strategy, PNGA, Biodiversity Strategy, Integrated "
                "Urban Solid Waste Strategy, and Environmental Quality and Effluents Regulation apply "
                "in an articulated manner to coastal zone management."
            )),
            ("Beach administration duties (Art. 10)", (
                "Beach administrators must: protect coastal ecosystems; prevent and combat marine "
                "and coastal pollution; ensure adequate urban solid waste management (collection, "
                "transport, treatment, storage, final disposal); guarantee waste is not discharged "
                "to beaches, sea or watercourses; and promote applied marine/coastal research."
            )),
        ]),
        ("Chapter III — Pollution Prevention and Conservation", [
            ("General pollution rules (Art. 34)", (
                "Permanent education, awareness, cleaning, monitoring and fiscalization activities "
                "must ensure environmental quality. All parties must focus on eliminating pollution "
                "sources and preventing disposal of any solid or liquid waste that could harm "
                "environmental quality, health or biodiversity."
            )),
            ("Solid waste management (Arts. 35–36)", (
                "Article 35: Beach users must collect food and other solid waste and deposit it in "
                "containers, ecopoints or bins, or carry it to the nearest receptacle. Economic "
                "operators are responsible for waste from their activities.\n\n"
                "Article 36: Urban equipped beaches — concessionaires clean concession areas, "
                "maritime authority cleans remaining areas. Unequipped/restricted beaches — maritime "
                "authority, delegated entities or local communities clean as defined case by case."
            )),
            ("Mandatory beach services (Art. 11)", (
                "Beaches must provide: lifeguard rescue; first aid post; police post; waste "
                "collection and beach cleaning; sanitary facilities; and changing rooms."
            )),
        ]),
        ("Chapter V — Fishing Camps and Aquaculture (Art. 47)", (
            "Fishers and aquaculturists must observe good environmental management rules, notably:\n"
            "a) correctly manage waste with special focus on all types of plastics including fishing "
            "ropes, nets, buoys, lines, hooks, equipment and vessel debris; keep fishing camps and "
            "mariculture infrastructure clean;\n"
            "b) not destroy dunes or coastal vegetation;\n"
            "c) collaborate with authorities on conservation best practices.\n\n"
            "Non-compliance is sanctionable under Annex II."
        )),
        ("Chapter VI — Prohibitions, Fiscalization and Sanctions", [
            ("Prohibitions (Art. 50)", (
                "Prohibited throughout coastal zones and beaches:\n"
                "a) sand extraction except return to beach;\n"
                "b) dumping, abandoning, burying or burning any solid or liquid waste;\n"
                "c) destruction of sensitive ecosystems;\n"
                "d) creating dumps in beach, dune or mangrove ecosystems;\n"
                "h) use of glass and plastic packaging in bathing zones except licensed catering;\n"
                "m) vehicle traffic on dunes and beach sand except as legally permitted;\n"
                "and other conduct affecting safety, signage and environmental quality."
            )),
            ("Fiscalization (Arts. 48–49)", (
                "Coordinated fiscalization by all competent entities; officials may draw up notices "
                "and seize materials; community agents may assist. All users and operators must "
                "cooperate with fiscalization agents."
            )),
            ("Sanctions (Arts. 51–53, Annex II highlights)", (
                "Fines measured in minimum public-sector salaries, plus mandatory restoration/"
                "indemnification; criminal sanctions may also apply. Accessory penalties include "
                "equipment seizure, licence suspension (up to 2 years), compulsory removal and "
                "establishment closure.\n\n"
                "Annex II selected fines:\n"
                "— Waste dumping: 3 minimum salaries;\n"
                "— Illegal dumps in sensitive ecosystems: 12 salaries;\n"
                "— Glass packaging in bathing zones: 3 salaries and seizure (Art. 50.1h also covers "
                "plastic packaging);\n"
                "— Economic operator non-compliance: 2–6 salaries;\n"
                "— Payment deadline: 20 days."
            )),
        ]),
        ("Note on Plastics", (
            "Plastics are explicitly mentioned in Art. 50.1h (bathing-zone packaging ban) and Art. 47 "
            "(fishing/aquaculture plastic waste including gear). The zero-waste principle (Art. 4f) "
            "and waste dumping prohibition (Art. 50.1b) address coastal plastic leakage more broadly. "
            "Operational waste services link to Decree 94/2014 and POEM action plans."
        )),
    ]

    for heading, content in sections:
        add_heading(doc, heading, level=1)
        if isinstance(content, str):
            add_para(doc, content)
        else:
            for subheading, body in content:
                add_heading(doc, subheading, level=2)
                add_para(doc, body)

    doc.add_page_break()
    add_heading(doc, "4P Index Coding Summary", level=1)
    add_para(
        doc,
        f"Policy: {POLICY_FIELDS['policy_name']}\n"
        f"Instruments coded: {len(INSTRUMENTS)} (consolidated — beach classification not overcoded)\n"
        f"See: 4P_Index_Mozambique_Decree_97_2020_Coastal_Zones.xlsx",
    )

    path = os.path.join(
        OUTPUT_DIR, "Mozambique_Decree_97_2020_Coastal_Zones_English_Translation.docx"
    )
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Decree 97/2020 Coastal Zones — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Decree 97/2020 (Coastal Zones &amp; Beaches)</h1>
  <p>Generated {date.today().isoformat()}. Six consolidated instrument rows.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Decree_97_2020_Coastal_Zones_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_decree_97_2020_coastal_zones.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Decree_97_2020_Coastal_Zones"
    df = build_dataframe()
    paths = [
        export_csv(df, basename),
        export_xlsx(df, basename),
        create_translation_doc(),
        create_index_html(basename),
    ]
    print("Generated:")
    for p in paths:
        print(f"  {p}")


if __name__ == "__main__":
    main()
