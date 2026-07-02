#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru Law 30884 (single-use plastics)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-ley-30884")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BRANCH = "cursor/peru-ley-30884-4p-index-d350"
REPO_RAW = f"https://github.com/shivrain/Thinkswiss/raw/{BRANCH}"

POLICY_NAME_ES = (
    "Ley N.° 30884 – Ley que regula el plástico de un solo uso y los recipientes o envases "
    "descartables"
)
POLICY_NAME_EN = (
    "Law No. 30884 – Law that Regulates Single-Use Plastic and Disposable Containers or "
    "Packaging"
)
POLICY_URL = (
    "https://www.gob.pe/institucion/congreso-de-la-republica/normas-legales/1122664-30884"
)


def bi(es: str, en: str) -> str:
    return f"{es} / {en}"


POLICY = {
    "policy_name": bi(POLICY_NAME_ES, POLICY_NAME_EN),
    "policy_url": POLICY_URL,
    "policy_year": 2018,
    "policy_objective": bi(
        "Ley marco nacional que regula el plástico de un solo uso, otros plásticos no "
        "reutilizables y envases de poliestireno expandido (tecnopor) para alimentos y "
        "bebidas, con el fin de reducir el impacto del plástico, la basura marina/fluvial/"
        "lacustre y contaminantes similares en la salud y el ambiente mediante prohibiciones "
        "progresivas, contenido reciclado mínimo en botellas PET, registro de "
        "productores/importadores, impuesto al consumo de bolsas plásticas y educación "
        "ciudadana.",
        "Flagship national law regulating single-use plastic, other non-reusable plastics "
        "and expanded-polystyrene (styrofoam) food/beverage containers, to reduce plastic "
        "impacts, marine/fluvial/lacustrine plastic litter and similar pollutants on health "
        "and the environment through progressive prohibitions, minimum recycled content in "
        "PET bottles, producer/importer registry, plastic-bag consumption tax and citizen "
        "education.",
    ),
    "policy_target": 1,
    "policy_target_text": bi(
        "Art. 2.1: reemplazo progresivo de bolsas en 36 meses. Art. 3.1: prohibiciones a "
        "120 días en áreas protegidas, playas y administración estatal. Art. 3.2: a 12 meses "
        "prohibición de bolsas <900 cm² y espesor <50 µm; sorbetes; bolsas con aditivos "
        "fragmentadores. Art. 3.3: a 36 meses prohibición total de bolsas no reutilizables, "
        "vajilla polimérica no reciclable y tecnopor. Art. 10.1: mínimo 15% PET-PCR en "
        "botellas en 3 años. Art. 12.5: impuesto gradual S/0.10 (2019) a S/0.50 (2023+). "
        "Art. 12.10: vigencia impuesto 1 ago. 2019. Glosario: bolsas reutilizables mínimo 15 "
        "usos.",
        "Art. 2.1: progressive bag replacement within 36 months. Art. 3.1: prohibitions at "
        "120 days in protected areas, beaches and state administration. Art. 3.2: at 12 "
        "months prohibition of bags <900 cm² and thickness <50 µm; straws; bags with "
        "fragmentation additives. Art. 3.3: at 36 months full ban on non-reusable bags, "
        "non-recyclable polymer tableware and styrofoam. Art. 10.1: minimum 15% post-consumer "
        "recycled PET in bottles within 3 years. Art. 12.5: graduated tax from PEN 0.10 "
        "(2019) to PEN 0.50 (2023+). Art. 12.10: tax effective 1 Aug 2019. Glossary: "
        "reusable bags minimum 15 uses.",
    ),
    "policy_type": 1.0,
    "policy_type_justification": bi(
        "Ley N.° 30884 aprobada por el Congreso de la República (8 dic. 2018), promulgada "
        "18 dic. 2018 y publicada 19 dic. 2018. No enmendada (reglamento modificado "
        "2019/2020). Clasificado 1.0 (legislación parlamentaria).",
        "Law No. 30884 enacted by Congress (8 Dec 2018), promulgated 18 Dec 2018 and "
        "published 19 Dec 2018. Not amended (implementing regulation amended 2019/2020). "
        "Classified 1.0 (parliamentary legislation).",
    ),
    "policy_integration": 1,
    "policy_sectors_list": bi(
        "retail, comercio, manufactura, alimentos y bebidas, turismo, administración pública, "
        "ambiente, educación, reciclaje, empaques, gobiernos locales y regionales",
        "retail, commerce, manufacturing, food & beverage, tourism, public administration, "
        "environment, education, recycling, packaging, local and regional governments",
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": bi(
        "producción, consumo, reciclaje, disposición, fuga ambiental",
        "production, consumption, recycling, disposal, environmental leakage",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": bi(
        "Art. 12: créase impuesto al consumo de bolsas de plástico (gradual S/0.10–S/0.50); "
        "'constituye ingreso del tesoro público' administrado por SUNAT. DCF Primera: política "
        "pública puede incluir estímulos e incentivos tributarios o no tributarios. DCF "
        "Tercera: sin recursos adicionales del tesoro para entidades estatales.",
        "Art. 12: creates plastic-bag consumption tax (graduated PEN 0.10–0.50); 'constitutes "
        "public treasury revenue' administered by SUNAT. Final Complementary Provision 1: "
        "public policy may include stimulus and tax/non-tax incentives. Third FCP: no "
        "additional treasury resources for state entities.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 1: objeto — marco regulatorio sobre plástico de un solo uso, plásticos no "
            "reutilizables y envases de tecnopor; finalidad reducir impacto de plástico y "
            "basura marina/fluvial/lacustre.",
            "Art. 1: purpose — regulatory framework on single-use plastic, non-reusable "
            "plastics and styrofoam containers; aim to reduce plastic impact and "
            "marine/fluvial/lacustrine litter.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "Marco legal vinculante. Sin autoridad o sanción específica en Art. 1.",
            "Binding legal framework. No specific authority or penalty in Art. 1.",
        ),
        "comments": bi(
            "Ley bandera nacional de plásticos de un solo uso en Perú.",
            "Flagship national single-use plastics law in Peru.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 2.1: supermercados, comercios y similares 'deben reemplazar en forma "
            "progresiva' bolsas no reutilizables por reutilizables u otras sin microplásticos "
            "— plazo 36 meses.",
            "Art. 2.1: supermarkets, retailers and similar establishments 'must progressively "
            "replace' non-reusable bags with reusable or other non-microplastic alternatives "
            "— within 36 months.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben reemplazar'. Plazo 36 meses. OEFA/PRODUCE/INDECOPI fiscalizan Art. 8. "
            "Sanciones Art. 9 Ley 28611.",
            "'Must replace'. 36-month deadline. OEFA/PRODUCE/INDECOPI enforce Art. 8. "
            "Sanctions Art. 9 under Law 28611.",
        ),
        "comments": bi(
            "Reducción progresiva de bolsas plásticas en comercio minorista.",
            "Progressive retail plastic-bag reduction.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 2.2: establecimientos 'deben cobrar, por cada bolsa que entregan, como "
            "mínimo una suma equivalente al precio del mercado', informando explícitamente "
            "al consumidor.",
            "Art. 2.2: establishments 'must charge, for each bag delivered, at least an amount "
            "equivalent to the market price', with explicit consumer information.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben cobrar' precio mínimo de mercado. INDECOPI protege consumidor Art. 8. "
            "Complementa impuesto Art. 12.",
            "'Must charge' at least market price. INDECOPI consumer protection Art. 8. "
            "Complements Art. 12 tax.",
        ),
        "comments": bi(
            "Medida económica de desincentivo al uso de bolsas plásticas.",
            "Economic disincentive for plastic bag use.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": bi(
            "Art. 3.1: a 120 días se prohíbe adquisición/uso/comercialización de bolsas "
            "poliméricas, sorbetes y tecnopor en áreas naturales protegidas, patrimonio "
            "cultural/natural, museos, playas del litoral y Amazonía, y administración estatal.",
            "Art. 3.1: at 120 days acquisition/use/commercialization of polymer bags, straws "
            "and styrofoam is prohibited in protected natural areas, cultural/natural heritage, "
            "museums, coastal and Amazon beaches, and state administration.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "'Se prohíbe'. Plazo 120 días. SERNANP, MINCUL, OEFA y gobiernos locales "
            "fiscalizan Art. 8. Sanciones Art. 136 Ley 28611.",
            "'Is prohibited'. 120-day deadline. SERNANP, MINCUL, OEFA and local governments "
            "enforce Art. 8. Sanctions Law 28611 Art. 136.",
        ),
        "comments": bi(
            "Primera fase de prohibición en áreas sensibles y playas — fuga de plásticos.",
            "First-phase ban in sensitive areas and beaches — plastic leakage.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 3.2(a): a 12 meses se prohíbe fabricación, importación, distribución y "
            "comercialización de bolsas poliméricas con área <900 cm² y espesor <50 µm.",
            "Art. 3.2(a): within 12 months manufacture, import, distribution and "
            "commercialization of polymer bags with area <900 cm² and thickness <50 µm is "
            "prohibited.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "'Se prohíbe' con umbrales cuantificados 900 cm² y 50 µm. PRODUCE/OEFA "
            "fiscalización. Tipificado Art. 9.",
            "'Is prohibited' with quantified thresholds 900 cm² and 50 µm. PRODUCE/OEFA "
            "enforcement. Typified Art. 9.",
        ),
        "comments": bi(
            "Prohibición de bolsas plásticas pequeñas y delgadas.",
            "Ban on small thin plastic bags.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 3.2(b): a 12 meses se prohíbe fabricación, importación, distribución y "
            "uso de sorbetes de base polimérica (pajitas, pitillos, popotes, cañitas), salvo "
            "excepciones Art. 4.3.",
            "Art. 3.2(b): within 12 months manufacture, import, distribution and use of "
            "polymer-based straws is prohibited, except Art. 4.3 exceptions.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "'Se prohíbe' plazo 12 meses. Excepciones médicas/discapacidad Art. 4.3.",
            "'Is prohibited' 12-month deadline. Medical/disability exceptions Art. 4.3.",
        ),
        "comments": bi(
            "Prohibición nacional de sorbetes/pajitas plásticas.",
            "National ban on plastic straws.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 3.2(c): a 12 meses se prohíbe bolsas no biodegradables con aditivos que "
            "catalizan fragmentación en microplástico.",
            "Art. 3.2(c): within 12 months non-biodegradable bags with additives catalyzing "
            "fragmentation into microplastics are prohibited.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "'Se prohíbe' oxodegradables/microplásticos. Vinculado certificado Art. 11.",
            "'Is prohibited' oxo-degradable/microplastic bags. Linked to Art. 11 certificate.",
        ),
        "comments": bi(
            "Restricción de plásticos con aditivos fragmentadores.",
            "Restriction on fragmentation-additive plastics.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 3.3(a): a 36 meses se prohíbe fabricación, importación, distribución y "
            "consumo de bolsas plásticas no reutilizables que generen microplásticos o no "
            "aseguren valorización.",
            "Art. 3.3(a): within 36 months manufacture, import, distribution and consumption "
            "of non-reusable plastic bags generating microplastics or not ensuring "
            "valorization is prohibited.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Se prohíbe' plazo 36 meses. Alineado con Art. 2.1. Reglamento define "
            "progresividad.",
            "'Is prohibited' 36-month deadline. Aligned with Art. 2.1. Regulation defines "
            "progressivity.",
        ),
        "comments": bi(
            "Prohibición total final de bolsas de un solo uso.",
            "Final comprehensive single-use bag ban.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 3.3(b): a 36 meses se prohíbe platos, vasos y vajilla polimérica no "
            "reciclable para alimentos y bebidas.",
            "Art. 3.3(b): within 36 months non-recyclable polymer plates, cups and tableware "
            "for food and beverages are prohibited.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Se prohíbe' vajilla descartable polimérica. Plazo 36 meses.",
            "'Is prohibited' disposable polymer tableware. 36-month deadline.",
        ),
        "comments": bi(
            "Prohibición de utensilios plásticos desechables para comida.",
            "Ban on disposable plastic food service items.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 3.3(c): a 36 meses se prohíbe fabricación, importación, distribución y uso "
            "de recipientes/envases y vasos de poliestireno expandido (tecnopor) para alimentos "
            "y bebidas; reglamento establece progresividad para MYPE.",
            "Art. 3.3(c): within 36 months manufacture, import, distribution and use of "
            "expanded-polystyrene (styrofoam) containers/cups for food and beverages is "
            "prohibited; regulation sets MSME progressivity.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Se prohíbe' tecnopor. Comisión técnica DCF Octava evalúa impacto MYPE.",
            "'Is prohibited' styrofoam. Eighth FCP technical commission evaluates MSME impact.",
        ),
        "comments": bi(
            "Prohibición de envases de poliestireno expandido (tecnopor).",
            "Ban on expanded-polystyrene (styrofoam) containers.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 5.1: INACAL aprueba NTP y reglamentos técnicos peruanos de productos "
            "poliméricos en 240 días; especificaciones para bolsas reutilizables y "
            "biodegradables según NTP 900.080.",
            "Art. 5.1: INACAL approves Peruvian technical standards and technical regulations "
            "for polymer products within 240 days; specifications for reusable and biodegradable "
            "bags per NTP 900.080.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "INACAL designado. Plazo 240 días. DS refrendado MINAM/PRODUCE. PRODUCE "
            "fiscaliza reglamentos técnicos Art. 8.",
            "INACAL designated. 240-day deadline. Supreme decree countersigned MINAM/PRODUCE. "
            "PRODUCE enforces technical regulations Art. 8.",
        ),
        "comments": bi(
            "Estándares técnicos de diseño para alternativas al plástico de un solo uso.",
            "Design standards for single-use plastic alternatives.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 6: MINAM con PRODUCE y SUNAT implementa registro de fabricantes, "
            "importadores y distribuidores en 120 días; inscripción obligatoria y reporte "
            "anual de información estadística.",
            "Art. 6: MINAM with PRODUCE and SUNAT implements registry of manufacturers, "
            "importers and distributors within 120 days; mandatory registration and annual "
            "statistical reporting.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Implementa un registro' y 'se inscriben' en 120 días. MINAM/PRODUCE/SUNAT "
            "designados. Monitoreo estadístico Art. 6.3.",
            "'Implements a registry' and 'must register' within 120 days. MINAM/PRODUCE/SUNAT "
            "designated. Statistical monitoring Art. 6.3.",
        ),
        "comments": bi(
            "Registro nacional de productores/importadores de plásticos regulados.",
            "National registry of regulated plastic producers/importers.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 7: MINAM, MINEDU, PRODUCE y gobiernos descentralizados desarrollan "
            "educación y sensibilización; establecimientos implementan estrategias de "
            "educación ambiental; 'Día Internacional Libre de Bolsas' (3 julio) y 'Día del "
            "Reciclaje del Plástico' (miércoles).",
            "Art. 7: MINAM, MINEDU, PRODUCE and decentralized governments develop education "
            "and awareness; establishments implement environmental education strategies; "
            "International Plastic Bag Free Day (3 July) and Plastic Recycling Day "
            "(Wednesdays).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Autoridades designadas. 'Desarrollan acciones' y 'deben implementar'. Sin "
            "sanción específica en Art. 7.",
            "Authorities designated. 'Develop actions' and 'must implement'. No specific "
            "penalty in Art. 7.",
        ),
        "comments": bi(
            "Campañas de consumo responsable y reciclaje de plástico.",
            "Responsible consumption and plastic recycling campaigns.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 8–9: OEFA supervisa/fiscaliza/sanciona obligaciones ambientales; PRODUCE "
            "reglamentos técnicos; INDECOPI consumidor; SERNANP áreas protegidas; gobiernos "
            "locales/regionales; infracciones tipificadas en reglamento con sanciones Art. 136 "
            "Ley 28611.",
            "Art. 8–9: OEFA supervises/enforces/sanctions environmental obligations; PRODUCE "
            "technical regulations; INDECOPI consumer protection; SERNANP protected areas; "
            "local/regional governments; infractions typified in regulation with Law 28611 "
            "Art. 136 sanctions.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "Múltiples autoridades designadas Art. 8. 'Constituyen infracciones' Art. 9. "
            "Multas Ley 28611. Reporte anual DCF Séptima.",
            "Multiple authorities designated Art. 8. 'Constitute infractions' Art. 9. Fines "
            "under Law 28611. Annual report Seventh FCP.",
        ),
        "comments": bi(
            "Régimen sancionador multisectorial para incumplimiento de prohibiciones plásticas.",
            "Multisector sanction regime for plastic-ban non-compliance.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 10.1–10.4: fabricantes de botellas PET 'deben obligatoriamente incluir' "
            "mínimo 15% PET-PCR en composición; envasadores e importadores deben cumplir; "
            "vigencia en 3 años desde publicación.",
            "Art. 10.1–10.4: PET bottle manufacturers 'must mandatorily include' minimum 15% "
            "post-consumer recycled PET in composition; packagers and importers must comply; "
            "effective within 3 years of publication.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben obligatoriamente incluir' 15% PET-PCR. Plazo 3 años. Exclusión hotfill/"
            "retornables Art. 10.5. PRODUCE/OEFA fiscalización.",
            "'Must mandatorily include' 15% post-consumer recycled PET. 3-year deadline. "
            "Hotfill/returnable exclusion Art. 10.5. PRODUCE/OEFA enforcement.",
        ),
        "comments": bi(
            "Contenido reciclado mínimo en botellas PET — diseño circular.",
            "Minimum recycled content in PET bottles — circular design.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 11: productores e importadores de plásticos biodegradables 'deben contar "
            "con' certificado de biodegradabilidad de laboratorio acreditado; bienes exceptuados "
            "del impuesto Art. 12.",
            "Art. 11: producers and importers of biodegradable plastics 'must have' "
            "biodegradability certificate from accredited laboratory; goods exempted from "
            "Art. 12 tax.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben contar con' certificado. Fiscalización mercado Art. 11.3. NTP 900.080 "
            "glosario.",
            "'Must have' certificate. Market enforcement Art. 11.3. NTP 900.080 glossary.",
        ),
        "comments": bi(
            "Estándar de biodegradabilidad para alternativas al plástico convencional.",
            "Biodegradability standard for alternatives to conventional plastic.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 12: créase impuesto al consumo de bolsas de plástico; gradual S/0.10 (2019) "
            "a S/0.50 (2023+); agentes de percepción IGV; vigencia 1 agosto 2019; ingreso "
            "tesoro público SUNAT.",
            "Art. 12: creates plastic-bag consumption tax; graduated from PEN 0.10 (2019) to "
            "PEN 0.50 (2023+); VAT withholding agents; effective 1 August 2019; public "
            "treasury revenue administered by SUNAT.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "'Créase el impuesto'. Cronograma cuantificado Art. 12.5. SUNAT administra y "
            "percibe. Comprobante de pago Art. 12.6. Exención biodegradables Art. 11.4.",
            "'Creates the tax'. Quantified schedule Art. 12.5. SUNAT administers and collects. "
            "Invoice disclosure Art. 12.6. Biodegradable exemption Art. 11.4.",
        ),
        "comments": bi(
            "Impuesto al consumo de bolsas plásticas — principal instrumento económico.",
            "Plastic bag consumption tax — main economic instrument.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "DCF Segunda: Poder Ejecutivo promueve formalización de actores de cadena de "
            "valor del plástico incluyendo recicladores; gobiernos locales incorporan "
            "recuperación de plásticos en programas de segregación y recolección selectiva con "
            "participación de recicladores.",
            "Second FCP: Executive promotes formalization of plastic value-chain actors "
            "including recyclers; local governments incorporate plastic recovery in segregation "
            "and selective-collection programs with recycler participation.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Promueve la formalización' y 'deben incorporar'. Vinculado Ley 29419 y DL 1278. "
            "Sin sanción directa.",
            "'Promotes formalization' and 'must incorporate'. Linked to Law 29419 and DL 1278. "
            "No direct penalty.",
        ),
        "comments": bi(
            "Integración de recicladores en gestión de plásticos postconsumo.",
            "Integration of waste pickers in post-consumer plastic management.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "DCF Cuarta: establecimientos comerciales y bienes regulados 'deben exhibir' "
            "anuncio informativo visible según reglamento.",
            "Fourth FCP: commercial establishments and regulated goods 'must display' a "
            "visible informative notice per regulation.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Deben exhibir' mensaje informativo. INDECOPI/OEFA pueden fiscalizar información "
            "al consumidor.",
            "'Must display' informative message. INDECOPI/OEFA may oversee consumer "
            "information.",
        ),
        "comments": bi(
            "Etiquetado/información al consumidor sobre restricciones plásticas.",
            "Consumer information/labelling on plastic restrictions.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "DCF Sexta: mediante DS se 'podrá ampliar' sustitución progresiva y prohibición "
            "de otros bienes poliméricos e incrementar porcentaje de material reciclado en "
            "envases.",
            "Sixth FCP: by supreme decree the Executive 'may expand' progressive substitution "
            "and prohibition of other polymer goods and increase recycled-material percentage "
            "in packaging.",
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "Poder habilitante ('podrá ampliar'). MINAM/PRODUCE. Depende de DS posterior — "
            "in_force 0 salvo DS emitido.",
            "Enabling power ('may expand'). MINAM/PRODUCE. Depends on subsequent supreme "
            "decree — in_force 0 unless decree issued.",
        ),
        "comments": bi(
            "Habilitación para ampliar restricciones a más plásticos vía reglamento.",
            "Enabling expansion of restrictions to additional plastics via regulation.",
        ),
    },
]

COLUMNS = [
    ("A", "policy_name"),
    ("B", "policy_url"),
    ("C", "policy_year"),
    ("D", "policy_objective"),
    ("E", "policy_target"),
    ("F", "policy_target_text"),
    ("G", "policy_type"),
    ("H", "policy_type_justification"),
    ("I", "policy_integration"),
    ("J", "policy_sectors_list"),
    ("K", "policy_circularity"),
    ("L", "policy_lifecycle_phases_list"),
    ("M", "policy_budget"),
    ("N", "policy_budget_text"),
    ("O", "policy_score"),
    ("P", "instrument_type"),
    ("Q", "instrument_lifecycle_stage"),
    ("R", "instrument_description"),
    ("S", "instrument_in_force"),
    ("T", "instrument_implementation"),
    ("U", "instrument_implementation_text"),
    ("V", "instrument_score"),
    ("W", "comments"),
]


def policy_score_row(instrument_type: float) -> float:
    return round(
        (
            POLICY["policy_type"]
            + POLICY["policy_integration"]
            + POLICY["policy_circularity"]
            + POLICY["policy_budget"]
            + instrument_type
        )
        / 5,
        3,
    )


def instrument_score_row(instrument_type: float, implementation: float) -> float:
    return round((instrument_type + implementation) / 2, 3)


def build_excel(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "4P Index Coding"
    ws.append([f"{col} — {name}" for col, name in COLUMNS])
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for inst in INSTRUMENTS:
        ws.append(
            [
                POLICY["policy_name"],
                POLICY["policy_url"],
                POLICY["policy_year"],
                POLICY["policy_objective"],
                POLICY["policy_target"],
                POLICY["policy_target_text"],
                POLICY["policy_type"],
                POLICY["policy_type_justification"],
                POLICY["policy_integration"],
                POLICY["policy_sectors_list"],
                POLICY["policy_circularity"],
                POLICY["policy_lifecycle_phases_list"],
                POLICY["policy_budget"],
                POLICY["policy_budget_text"],
                policy_score_row(inst["instrument_type"]),
                inst["instrument_type"],
                inst["instrument_lifecycle_stage"],
                inst["instrument_description"],
                inst["instrument_in_force"],
                inst["instrument_implementation"],
                inst["instrument_implementation_text"],
                instrument_score_row(inst["instrument_type"], inst["instrument_implementation"]),
                inst["comments"],
            ]
        )

    widths = {
        "A": 40, "B": 38, "C": 8, "D": 44, "E": 8, "F": 44, "G": 8, "H": 40,
        "I": 10, "J": 40, "K": 10, "L": 36, "M": 10, "N": 42, "O": 10, "P": 10,
        "Q": 18, "R": 50, "S": 10, "T": 12, "U": 44, "V": 10, "W": 40,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    wb.save(path)


def add_article(doc: Document, es: str, en: str) -> None:
    p = doc.add_paragraph()
    r = p.add_run(es)
    r.bold = True
    doc.add_paragraph(en).paragraph_format.space_after = Pt(8)


def build_word(path: Path) -> None:
    doc = Document()
    for margin in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(doc.sections[0], margin, Inches(1))

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t = title.add_run(f"{POLICY_NAME_EN}\n{POLICY_NAME_ES}")
    t.bold = True
    t.font.size = Pt(14)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run(
        "Enacted: 8 December 2018 | Promulgated: 18 December 2018 | Published: 19 December 2018\n"
        "Not amended (implementing regulation amended 2019/2020)\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Law No. 30884 on single-use plastics "
        "and disposable containers."
    )

    sections = [
        (
            "Article 1 — Purpose",
            "El objeto de la ley es establecer el marco regulatorio sobre el plástico de un solo "
            "uso... reduciendo el impacto adverso del plástico... basura marina plástica...",
            "The law establishes the regulatory framework on single-use plastic, non-reusable "
            "plastics and styrofoam food/beverage containers, reducing adverse plastic impacts "
            "and marine/fluvial/lacustrine plastic litter.",
        ),
        (
            "Article 2 — Progressive bag reduction",
            "deben reemplazar en forma progresiva... treinta y seis (36) meses... deben cobrar, "
            "por cada bolsa... precio del mercado...",
            "Retailers must progressively replace non-reusable polymer bags within 36 months and "
            "charge at least market price per bag with explicit consumer information.",
        ),
        (
            "Article 3 — Prohibitions (phased)",
            "120 días... áreas naturales protegidas, playas... 12 meses... bolsas... 900 cm2... "
            "50 micras... sorbetes... 36 meses... tecnopor...",
            "Phased prohibitions: 120 days in protected areas/beaches/state admin; 12 months on "
            "small/thin bags and straws; 36 months on all non-reusable bags, non-recyclable "
            "tableware and styrofoam food/beverage containers.",
        ),
        (
            "Article 6 — Producer registry",
            "implementa un registro de fabricantes, importadores y distribuidores... ciento veinte "
            "(120) días...",
            "MINAM with PRODUCE and SUNAT implements manufacturer/importer/distributor registry "
            "within 120 days with annual statistical reporting.",
        ),
        (
            "Article 10 — Recycled PET content",
            "deben obligatoriamente incluir... material PET reciclado postconsumo (PET-PCR) en al "
            "menos quince por ciento (15%)... tres (3) años...",
            "PET bottle manufacturers must mandatorily include at least 15% post-consumer recycled "
            "PET within 3 years of publication.",
        ),
        (
            "Article 12 — Plastic bag consumption tax",
            "Créase el impuesto al consumo de las bolsas de plástico... S/ 0.10 en el 2019... "
            "S/ 0.50 en el 2023... 1 de agosto del año 2019.",
            "Creates plastic-bag consumption tax graduated from PEN 0.10 (2019) to PEN 0.50 "
            "(2023+), effective 1 August 2019, administered by SUNAT.",
        ),
        (
            "Articles 8–9 — Enforcement",
            "OEFA... supervisión, fiscalización y sanción... Constituyen infracciones el "
            "incumplimiento... artículo 136 de la Ley 28611...",
            "OEFA, PRODUCE, INDECOPI, SERNANP and local/regional governments enforce the law; "
            "non-compliance constitutes infractions sanctioned under General Environment Law "
            "Art. 136.",
        ),
        (
            "Final provisions — Recyclers and information",
            "formalización de los recicladores... Programas de Segregación... deben exhibir un "
            "anuncio...",
            "Executive promotes recycler formalization; local governments integrate plastic "
            "recovery in selective collection; establishments must display visible informative "
            "notices.",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: Law 30884 is Peru's flagship national single-use plastics law. It complements "
        "DS 013-2018-MINAM (Executive Branch entities only) and is implemented through Supreme "
        "Decree regulations (approved 2019, amended 2020). DS 014-2017-MINAM and DL 1278 provide "
        "the broader waste-management framework."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_Law_30884_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_Law_30884_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")
    print()
    print("Download links:")
    print(f"  Excel: {REPO_RAW}/output/peru-ley-30884/Peru_Law_30884_4P_Index_Coding.xlsx")
    print(f"  Word:  {REPO_RAW}/output/peru-ley-30884/Peru_Law_30884_English_Translation.docx")


if __name__ == "__main__":
    main()
