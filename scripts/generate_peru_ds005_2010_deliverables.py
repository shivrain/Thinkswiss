#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru DS 005-2010-MINAM (Law 29419 Regulation)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-ds-005-2010-minam")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

POLICY_NAME_ES = (
    "Reglamento de la Ley Nº 29419, Ley que regula la actividad de los recicladores "
    "(Decreto Supremo Nº 005-2010-MINAM)"
)
POLICY_NAME_EN = (
    "Regulation of Law No. 29419 (Waste Pickers Law) — Supreme Decree No. 005-2010-MINAM"
)
POLICY_URL = (
    "https://sinia.minam.gob.pe/normas/reglamento-ley-no-29419-ley-que-regula-actividad-recicladores"
)
COUNTRY = "Peru"

POLICY = {
    "policy_name_es": POLICY_NAME_ES,
    "policy_name_en": POLICY_NAME_EN,
    "policy_url": POLICY_URL,
    "country": COUNTRY,
    "policy_year": 2022,
    "policy_objective": (
        "Implementing regulation for Law No. 29419 establishing detailed procedures for "
        "formalization, training, occupational health protection, municipal planning, selective "
        "collection, and commercialization by recycler organizations. Explicitly authorizes "
        "collection of all plastics and other recyclable inorganic/organic waste streams, and "
        "sets technical, administrative, incentive, and sanction frameworks for recycler "
        "associations recovering plastic and other recyclables."
    ),
    "policy_target": 1,
    "policy_target_text": (
        "Art. 29: 'diseñará un Plan Técnico Operativo para la Recolección Selectiva... con la "
        "finalidad de... lograr cubrir progresivamente la totalidad de predios ubicados en su "
        "jurisdicción' and 'modificar las estrategias para lograr cubrir integralmente el 100% "
        "de los predios ubicados en su jurisdicción'. Art. 48: donation certificate for "
        "enterprises delivering 'más de 500 kg de residuos sólidos reaprovechables al año'. "
        "Art. 47: household segregation and delivery 'al menos una vez por semana'."
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Supreme Decree (Decreto Supremo Nº 005-2010-MINAM) approved by the Executive Power "
        "(President Alan García, 2 June 2010), countersigned by Ministers of Environment and "
        "Health. Last amended by Supreme Decree Nº 001-2022-MINAM (9 January 2022). This is "
        "sub-legislative executive regulation, not parliamentary legislation."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "waste management, recycling, municipalities, consumption, education, health, industry, "
        "retail, packaging"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "consumption, recycling, disposal, environmental leakage",
    "policy_budget": 1,
    "policy_budget_text": (
        "Art. 54: 'El FONAM implementará y administrará un Fondo de Garantía para la operación "
        "del programa de crédito dirigido a los recicladores... captar la cooperación financiera "
        "internacional y nacional, donaciones y otros del sector público y privado.' DS "
        "001-2022-MINAM Art. 6: implementation financed from institutional budgets of involved "
        "entities."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 2: Regulation is of mandatory national compliance ('cumplimiento obligatorio a "
            "nivel nacional') for natural/juridical persons engaged in selective management of "
            "non-hazardous solid waste and institutional actors under Law 29419."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Applies nationally to defined actors. Mandatory language 'cumplimiento obligatorio'. "
            "Enforcement delegated to municipalities/OEFA in later articles. No single authority "
            "designated in Art. 2 itself."
        ),
        "comments": "Overarching binding scope instrument for the regulation.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 5: MINAM promotes integrated waste management and recycler formalization; "
            "systematizes formalization program information from municipalities via SINIA; "
            "promotes environmental education with municipalities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "MINAM designated. Art. 5.2 requires municipalities to remit information and MINAM "
            "to systematize/disseminate via SINIA (monitoring/data). Uses 'promover' for some "
            "functions; no explicit penalties in Art. 5."
        ),
        "comments": "National coordination and information systematization instrument.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 6: Ministry of Health implements vaccination programs (Tetanus and Hepatitis B) "
            "and sanitary surveillance of selective waste management, targeting formalized "
            "recycler organization members."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Ministry of Health designated ('implementar', 'desarrollar acciones de vigilancia "
            "sanitaria'). Surveillance is monitoring mechanism. No explicit fines in Art. 6."
        ),
        "comments": "Health protection regulatory standard for formalized recyclers.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 7.1–7.2: Municipalities must elaborate and implement the Program of "
            "Formalization of Recyclers and Selective Collection; incorporate recycling situational "
            "study, operational plan, and public education program into PIGARS or solid waste "
            "management plans."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Municipalities designated ('deberán elaborar e implementar', 'incorporar'). Planning "
            "obligations unconditional on municipalities. Art. 7.3 adds annual reporting "
            "(monitoring). No explicit penalties in 7.1–7.2."
        ),
        "comments": "Municipal planning and program implementation obligation.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 7.3: Municipalities must include implementation report on the Formalization "
            "Program in annual solid waste reports, remitted to MINAM within Q1 each year per "
            "Annex 1 format."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Municipalities must report ('deberá ser remitido al Ministerio del Ambiente'). "
            "MINAM recipient. Standardized Annex 1 format (monitoring/data collection). "
            "Mandatory reporting deadline."
        ),
        "comments": "Monitoring/reporting instrument with standardized national format.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 7.5: Municipalities implement incentive programs promoting source segregation "
            "of solid waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Municipalities must implement ('implementan programas'). Responsible authority "
            "designated. Program design flexible. No explicit penalties or audit requirements."
        ),
        "comments": "Economic incentive instrument for source segregation including plastics.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 7.6: Municipalities fiscalize segregation, selective collection, and recycler "
            "formalization activities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Municipalities designated to 'fiscalizar'. Art. 51 links fiscalization to fines, "
            "vehicle seizure, and closure (enforcement). Monitoring through fiscalization actions."
        ),
        "comments": "Enforcement-linked regulatory fiscalization duty on municipalities.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 7.7–7.8: Municipalities systematize recycler organization records on "
            "quantities recycled and conduct the Registry of Authorized Recycler Organizations "
            "legally established and registered."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Municipal registry and data systematization ('sistematizar', 'conducir el Registro'). "
            "Art. 38 (as amended 2022) requires reporting to MINAM via SIGERSOL for National "
            "Registry inclusion (monitoring)."
        ),
        "comments": "Governance/registry instrument. Art. 38–39 amended by DS 001-2022-MINAM.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 8: Obligations on recycler organizations with legal personality: participate in "
            "operational plans, register in municipal program, comply with routes/schedules, submit "
            "monthly recovery reports, ensure vaccination and PPE, register recycled quantities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Recycler organizations subject to mandatory duties ('deben', 'cumplir', 'remitir'). "
            "Monthly reporting (monitoring). Municipal oversight under Art. 7.6/50–51 "
            "(enforcement). Some conditional elements (e.g., program participation)."
        ),
        "comments": "Bundled regulatory obligations on formalized recycler organizations.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 9: MINAM, Health Ministry, provincial and district municipalities coordinate "
            "functions to achieve Law and Regulation objectives."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Institutional actors designated. 'Coordinarán' — coordination without explicit "
            "penalties or monitoring mechanisms in Art. 9."
        ),
        "comments": "Inter-institutional governance coordination framework.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 11–13: Mandatory personal protective equipment and uniforms for formalized "
            "recyclers by activity type (selective collection/transport vs. conditioning), "
            "including gloves, masks, footwear, reflective uniforms, and maternity overol "
            "requirements."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Debe contar con' / 'uso obligatorio' / 'deberá contar obligatoriamente'. Employers "
            "must provide PPE for dependent workers at no cost (Art. 10). Municipal "
            "fiscalization under Art. 7.6. No explicit fine amounts in Arts. 11–13."
        ),
        "comments": "Regulatory occupational safety and equipment standards.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 15–16: Classification and technical specifications for conventional and "
            "non-conventional collection vehicles (tricycles, motofurgones, carts, boats), "
            "including municipal authorization, SOAT, and licensing requirements."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Mandatory vehicle characteristics ('deben cumplir', 'deberán contar'). Provincial "
            "municipal authorization required for conventional vehicles. Art. 51.2 sanctions "
            "unauthorized vehicles (enforcement)."
        ),
        "comments": "Regulatory technical standards for collection/transport equipment.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 19: Authorized waste types for recycler collection explicitly include "
            "'Plástico: todos los plásticos' alongside paper, cardboard, metals, glass, rubber, "
            "textiles, and organics. Hazardous waste excluded."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Sólo podrán hacer recolección de' — restrictive mandatory list. Art. 51.4 sanctions "
            "hazardous waste collection; Art. 51.1 sanctions informal practices (enforcement). "
            "Municipal fiscalization applies."
        ),
        "comments": "Direct explicit plastics provision. Key plastics-relevance anchor.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 20–21: Segregation permitted at source, commercialization infrastructure, and "
            "landfill treatment plants; selective collection limited to recoverable domestic, "
            "commercial, and similar waste, conducted by formalized organizations per municipal "
            "operational plans without environmental exposure."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Generators must ensure sanitary handling at source (Art. 20.1 'deberá asegurar'). "
            "Municipalities must set technical conditions in operational plans (Art. 21). No "
            "explicit penalties in Arts. 20–21."
        ),
        "comments": "Regulatory segregation and selective collection framework.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 22: Organizations conducting storage and conditioning must comply with "
            "operational standards for recoverable waste infrastructure under Law 27314 and its "
            "regulation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "'Deberán cumplir con los requisitos y estándares operacionales' — mandatory. "
            "Cross-reference to general solid waste standards. Municipal/OEFA enforcement via "
            "broader framework."
        ),
        "comments": "Regulatory infrastructure/operations standard for conditioning sites.",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 23: Municipalities promote treatment plants within sanitary landfills for "
            "compost, biodigesters, segregation facilities; formalized recyclers may segregate "
            "recoverable waste with required PPE enforced by facility operators."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Municipalities 'promueven la implementación'. Operators must enforce PPE ('debiendo "
            "el operador exigir su uso'). Promotional rather than mandatory construction."
        ),
        "comments": "Infrastructure promotion instrument for landfill segregation facilities.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 28: Municipalities must prepare Recycling Situational Study identifying "
            "segregation potential, recycler census, commercialization chain, and municipal "
            "capacities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Deberán elaborar' — mandatory planning study. Municipal authority designated. "
            "Statistical/census components (monitoring). Unconditional obligation on municipalities."
        ),
        "comments": "Governance planning requirement preceding operational plan design.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 29: Municipalities must design Operational Plan for Selective Collection with "
            "progressive coverage of all properties in jurisdiction (100% target), route "
            "sectorization, annual physical/budget targets, and biennial updates."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Diseñará un Plan Técnico Operativo' with quantified 100% property coverage target. "
            "Annual goals and biennial update requirement (monitoring). Municipal authority "
            "designated. No explicit penalty in Art. 29."
        ),
        "comments": "Regulatory planning instrument with quantifiable coverage targets.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 30: Municipalities must incorporate program elements via municipal ordinance, "
            "institutional operational plan, TUPA procedures, sanctions regime, incentives, and "
            "fiscalization mechanisms."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Deberán ser incorporados' — mandatory municipal legal implementation package. "
            "Includes sanctions regime incorporation (Art. 30.4) and fiscalization (Art. 30.6). "
            "Municipalities designated."
        ),
        "comments": "Regulatory municipal transposition obligation for the national regulation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 31 & 34 (as amended DS 001-2022): Technical and administrative requirements for "
            "recycler formalization — training certificate, vaccination card, PPE, vehicles, "
            "municipal ID card; Art. 34 sets formalization application requirements and 15-day "
            "evaluation deadline with formalization certificate."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Deben pertenecer', 'deberán cumplir', 'deben presentar' — mandatory requirements. "
            "Municipal evaluation within 15 business days (monitoring). Carnet issuance. Art. 34 "
            "text replaced by DS 001-2022-MINAM; in force since 2022."
        ),
        "comments": "Core formalization regulatory instrument. Arts. 34, 38, 39 amended 2022.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 35: Recycler profile requirements — minimum age 18; pregnant independent "
            "recyclers limited to work until 7th month of pregnancy with medical controls for "
            "reincorporation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Deberán cumplir' — mandatory age and pregnancy restrictions. Art. 51.3 sanctions "
            "collection by minors (enforcement). Municipal authority for program admission."
        ),
        "comments": "Regulatory social protection and eligibility standard.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 36: All recyclers must mandatorily participate in training program with minimum "
            "four 3-hour modules (integrated waste management, occupational safety, business "
            "management, social skills), delivered by SENATI or other institutions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Todos los recicladores deberán participar obligatoriamente' — unconditional "
            "mandatory training. SENATI/institutions designated. Quantified module duration "
            "(monitoring standard). Required for formalization (Art. 31/34)."
        ),
        "comments": "Regulatory mandatory training standard with quantified minimum hours.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 37: All recyclers must be vaccinated against Hepatitis B and Tetanus; "
            "municipalities promote mass vaccination with Health Ministry; health facilities "
            "must issue vaccination cards."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Todos los recicladores deberán ser vacunados' — mandatory. Health establishments "
            "must issue cards ('deberán emitir'). Required for formalization. Municipal "
            "coordination role designated."
        ),
        "comments": "Regulatory health requirement linked to formalization.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 38–39 (as amended DS 001-2022): Municipal Registry of Recycler Organizations "
            "authorizes selective collection, transport, and conditioning; registration requires "
            "formalization certificate, sworn declarations on equipment/vehicles, 15-day "
            "evaluation, ID cards; information reported to MINAM via SIGERSOL for National "
            "Registry."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "Municipality administers registry ('administra el Registro'). Mandatory reporting to "
            "MINAM/SIGERSOL within 10 business days (2022 framework). 15-day evaluation "
            "(monitoring). ID cards mandatory during work. Art. 51 sanctions unauthorized "
            "informal activity (enforcement). Amended text operative since 2022."
        ),
        "comments": (
            "Key formalization/registry instrument. 2010 version required 3-year renewal; 2022 "
            "amendment restructured requirements and linked to National Registry."
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 40–43: MINAM, Education Ministry, and municipalities must design and implement "
            "Public Education and Communication Program on source segregation, recycling culture, "
            "and recycler role; includes school and community projects and mass-media campaigns."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. 40: 'deberán diseñar e implementar un Programa' — binding on MINAM/municipalities. "
            "Responsible authorities designated. No explicit penalties for non-compliance."
        ),
        "comments": "Information/education instrument for segregation and recycling awareness.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 47: Municipalities progressively establish incentive bonus for households "
            "segregating recoverable waste and delivering to authorized recyclers at least once "
            "weekly, providing tariff discounts on municipal fees via municipal ordinance."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "'Establecerá progresivamente un bono de incentivo' — phased economic incentive. "
            "Weekly delivery frequency specified. Municipal ordinance required. 'Progresivamente' "
            "reduces unconditionality."
        ),
        "comments": "Economic incentive for household source segregation.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 48: Environmental incentive program for enterprises/institutions segregating at "
            "source and donating recoverables to recycler organizations at least twice weekly; "
            "donation certificate issued for donations exceeding 500 kg/year."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Quantified 500 kg/year threshold and twice-weekly donation minimum. Municipal "
            "recognition incentive. Donation certificate requirement (monitoring/documentation). "
            "Municipal program responsibility designated."
        ),
        "comments": "Economic/recognition incentive with quantified donation threshold.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 49: MINAM awards National Recycling Prize on 1 June (National Recycler Day) to "
            "persons/entities distinguished for recycling integrating environmental, social, and "
            "economic benefits."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "MINAM designated to 'realizará la premiación'. No enforcement or monitoring "
            "mechanisms specified."
        ),
        "comments": "Voluntary recognition/awareness instrument.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 50–51: Municipal environmental fiscalization (vigilance, control, verification) "
            "over all regulated persons; municipalities must approve Sanctions Regime with "
            "differentiated fines, vehicle seizure, and closure for informal recycling, "
            "unauthorized vehicles, minor collection, hazardous waste collection, and other "
            "listed infringements. OEFA supervises compliance."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "Art. 50: all persons 'están sometida a las acciones de fiscalización'. Art. 51: "
            "'deberán aprobar su Régimen de Aplicación de Sanciones' with 'multas', "
            "'incautación de vehículos', 'clausura de locales'. OEFA evaluation/supervision. "
            "Municipal and OEFA authorities designated."
        ),
        "comments": "Core enforcement instrument with explicit pecuniary sanctions and seizures.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 52: Program must include citizen environmental vigilance component with "
            "methodology, service evaluation surveys, and reporting of non-compliance on "
            "segregation and collection obligations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Program 'contemplará un componente de vigilancia ambiental ciudadana'. Surveys and "
            "denunciation reporting (monitoring). No designated single authority or penalties in "
            "Art. 52."
        ),
        "comments": "Citizen monitoring component within formalization program.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 53: Municipalities must promote Local Recycling Working Table among public "
            "institutions, business representatives, and recyclers; promote annual regional "
            "recycling fairs and clean technology exchange."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "'Deberán impulsar' — municipal promotional duty. No explicit enforcement or "
            "monitoring in Art. 53."
        ),
        "comments": "Governance/multi-stakeholder coordination for recycling value chain.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 54: FONAM promotes credit program for recyclers and implements/administers "
            "Guarantee Fund capturing international/national cooperation, donations, and public/"
            "private resources; monitors and honors guarantees."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "FONAM designated to implement fund and administer guarantees with 'monitoreo' of "
            "guarantees (monitoring). No explicit penalties. Fund structure is dedicated "
            "(economic instrument)."
        ),
        "comments": "Economic credit/guarantee fund instrument from parent Law Art. 11.",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 55: Promotion of public/private land cession for recycler organizations to "
            "install conditioning infrastructure for recoverable solid waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "'Se promoverá' — promotional, not mandatory land cession. No authority, enforcement, "
            "or monitoring specified."
        ),
        "comments": "Infrastructure promotion instrument for conditioning facilities.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Final Complementary Provisions: Municipalities must implement Formalization Program "
            "within 12 months of regulation publication; municipalities with existing programs have "
            "12 months to comply with Law and Regulation requirements."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Deberán implementar... en un plazo máximo de doce (12) meses' — mandatory deadline. "
            "Municipalities designated. Quantified 12-month timeline. DS 001-2022 added further "
            "180-day deadlines for program updates."
        ),
        "comments": (
            "Implementation deadline instrument. DS 001-2022-MINAM added complementary deadlines "
            "for program/sanctions updates (180/120 days)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "DS 001-2022-MINAM (amending Arts. 34, 38, 39): National Registry of Recyclers "
            "administered by MINAM DGGRS via SIGERSOL; municipalities report formalized "
            "organizations within 10 business days; enables municipal registry inscription and "
            "participation in Source Segregation and Selective Collection Programs nationally."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "MINAM administers national registry (Art. 13-B DL 1278 reg, cross-referenced). "
            "10-business-day reporting deadline (monitoring). Public access registry. Added by "
            "2022 amendment to regulation framework."
        ),
        "comments": (
            "2022 amendment instrument operating alongside municipal registry (Art. 38–39). Coded "
            "separately as distinct national-level governance instrument."
        ),
    },
]

COLUMNS = [
    ("A", "policy_name_es"),
    ("B", "policy_name_en"),
    ("C", "country"),
    ("D", "policy_url"),
    ("E", "policy_year"),
    ("F", "policy_objective"),
    ("G", "policy_target"),
    ("H", "policy_target_text"),
    ("I", "policy_type"),
    ("J", "policy_type_justification"),
    ("K", "policy_integration"),
    ("L", "policy_sectors_list"),
    ("M", "policy_circularity"),
    ("N", "policy_lifecycle_phases_list"),
    ("O", "policy_budget"),
    ("P", "policy_budget_text"),
    ("Q", "policy_score"),
    ("R", "instrument_type"),
    ("S", "instrument_lifecycle_stage"),
    ("T", "instrument_description"),
    ("U", "instrument_in_force"),
    ("V", "instrument_implementation"),
    ("W", "instrument_implementation_text"),
    ("X", "instrument_score"),
    ("Y", "comments"),
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
                POLICY["policy_name_es"],
                POLICY["policy_name_en"],
                POLICY["country"],
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
        "A": 34, "B": 34, "C": 10, "D": 38, "E": 8, "F": 44, "G": 8, "H": 44,
        "I": 8, "J": 38, "K": 10, "L": 36, "M": 10, "N": 30, "O": 10, "P": 40,
        "Q": 10, "R": 10, "S": 18, "T": 50, "U": 10, "V": 12, "W": 44, "X": 10, "Y": 38,
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
        "Approved: 2 June 2010 | Last amended: 9 January 2022 (DS Nº 001-2022-MINAM)\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Supreme Decree No. 005-2010-MINAM and its "
        "2022 amendment (DS 001-2022-MINAM) affecting Articles 34, 38, and 39. Spanish headings "
        "retained for reference."
    )

    sections = [
        (
            "Decree Articles 1–2 — Approval",
            "Artículo 1º: Apruébese el Reglamento de la Ley Nº 29419...",
            "Article 1: The Regulation of Law No. 29419, Law that Regulates the Activity of Recyclers, "
            "consisting of seven Titles, fifty-five Articles, three Final Complementary Provisions and "
            "two Annexes, is approved as an integral part of this Supreme Decree.",
        ),
        (
            "Regulation Article 1 — Objective",
            "El objetivo del presente Reglamento es regular lo establecido en la Ley Nº 29419...",
            "The objective of this Regulation is to regulate provisions of Law No. 29419 to contribute "
            "to protection, training, and social/labor development of recycling workers, promoting "
            "formalization and association and improving adequate management for reuse of solid waste.",
        ),
        (
            "Regulation Article 2 — Scope",
            "El presente Reglamento es de cumplimiento obligatorio a nivel nacional...",
            "This Regulation is of mandatory national compliance for persons engaged in selective "
            "management of non-hazardous solid waste and institutional actors under Law 29419.",
        ),
        (
            "Regulation Article 5 — Ministry of Environment",
            "El Ministerio del Ambiente es responsable de: promover la gestión integral...",
            "The Ministry of Environment is responsible for: (5.1) promoting integrated solid waste "
            "management and recycler formalization; (5.2) systematizing municipal information on the "
            "Formalization and Selective Collection Program via SINIA; (5.3) promoting environmental "
            "education with municipalities.",
        ),
        (
            "Regulation Article 7 — Local Governments (excerpt)",
            "Las Municipalidades... son las responsables de elaborar e implementar el Programa...",
            "District and provincial municipalities are responsible for: elaborating and implementing "
            "the Program of Formalization of Recyclers and Selective Collection; incorporating recycling "
            "studies and operational plans into PIGARS; annual reporting to MINAM; implementing source "
            "segregation incentives; fiscalizing segregation/collection/formalization; maintaining the "
            "Registry of Authorized Recycler Organizations.",
        ),
        (
            "Regulation Article 19 — Authorized waste types (plastics)",
            "c) Plástico: todos los plásticos.",
            "c) Plastic: all plastics. (Also authorizes paper, cardboard, ferrous/non-ferrous metals, "
            "glass, rubber, textiles, and organics. Hazardous waste is excluded from collection scope.)",
        ),
        (
            "Regulation Article 29 — Operational Plan and coverage target",
            "lograr cubrir progresivamente la totalidad de predios... el 100% de los predios...",
            "Municipalities must design an Operational Plan for Selective Collection to progressively "
            "cover all properties in their jurisdiction, including strategies to reach 100% property "
            "coverage, with annual physical and budget targets and biennial updates.",
        ),
        (
            "Regulation Article 36 — Mandatory training",
            "Todos los recicladores deberán participar obligatoriamente de un Programa de Capacitación...",
            "All recyclers must mandatorily participate in a Training Program with at least four "
            "modules of three hours each: integrated waste management; occupational safety; business "
            "management and recycling; social skills and personal development. SENATI and other "
            "institutions may deliver the program.",
        ),
        (
            "Regulation Article 47 — Household incentive bonus",
            "bono de incentivo dirigido a los vecinos... segregan sus residuos... al menos una vez por semana",
            "Municipalities progressively establish an incentive bonus for residents participating in "
            "the formalization program who segregate recoverable waste and deliver it to authorized "
            "recyclers at least once per week, with tariff discounts regulated by municipal ordinance.",
        ),
        (
            "Regulation Article 48 — Enterprise environmental incentive",
            "más de 500 kg de residuos sólidos reaprovechables al año",
            "Enterprises/institutions implementing source segregation and donating recoverables to "
            "recycler organizations at least twice weekly receive municipal recognition; donation "
            "certificates are issued for deliveries exceeding 500 kg of recoverable waste per year.",
        ),
        (
            "Regulation Article 51 — Sanctions",
            "bajo sanción pecuniaria e incautación de los vehículos y clausura de locales",
            "Municipalities must approve a Sanctions Regime including differentiated fines, vehicle "
            "seizure, and premises closure for informal recycling, unauthorized vehicles, collection "
            "by minors, hazardous waste collection, and other listed non-compliant conduct.",
        ),
        (
            "Regulation Article 54 — FONAM credit and guarantee fund",
            "El FONAM implementará y administrará un Fondo de Garantía...",
            "FONAM will promote a credit program and implement/administer a Guarantee Fund for "
            "recyclers (individuals and organizations), capturing international/national cooperation "
            "and public/private donations, with monitoring and honoring of guarantees.",
        ),
        (
            "2022 Amendment — Article 34 (formalization requirements)",
            "Las organizaciones de recicladores deben presentar... solicitud para su formalización",
            "Recycler organizations must submit formalization requests to municipalities with personal "
            "data, training certificates, vaccination cards, equipment/vehicle lists, and processing "
            "fees. Municipal evaluation within 15 business days; formalization certificate issued if "
            "requirements met (DS 001-2022-MINAM, replacing prior Art. 34 text).",
        ),
        (
            "2022 Amendment — Article 38 (Municipal Registry)",
            "La municipalidad administra el Registro Municipal... reportada al MINAM, a través del SIGERSOL",
            "Municipalities administer the Registry of Recycler Organizations authorizing selective "
            "collection, transport, and conditioning. Information must be reported to MINAM through "
            "SIGERSOL for inclusion in the National Registry of Recyclers.",
        ),
        (
            "2022 Amendment — National Registry (Art. 13-B, cross-reference)",
            "El Registro Nacional de Recicladores es administrado por la Dirección General de Gestión de Residuos Sólidos del MINAM",
            "The National Registry of Recyclers is administered by MINAM's DGGRS, is publicly accessible, "
            "and systematizes information on formalized recycler organizations and members. Municipalities "
            "report within 10 business days of issuing formalization certificates.",
        ),
        (
            "Final Complementary Provisions",
            "implementar el Programa... en un plazo máximo de doce (12) meses",
            "From publication, municipalities must implement the Formalization and Selective Collection "
            "Program within 12 months. Municipalities with existing programs have 12 months to comply "
            "with Law and Regulation requirements.",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: This regulation implements Law No. 29419. The 2022 amendment (DS 001-2022-MINAM) also "
        "modified the Regulation of the Solid Waste Management Law (DL 1278) and introduced the "
        "National Registry of Recyclers and updated formalization/registry procedures."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_DS_005_2010_MINAM_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_DS_005_2010_MINAM_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")


if __name__ == "__main__":
    main()
