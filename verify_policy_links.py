#!/usr/bin/env python3
"""Verify direct PDF links for Mozambique policy documents."""

import re
import sys

import requests
from pypdf import PdfReader
from io import BytesIO

from generate_mozambique_plastic_policies_bilingual_docx import POLICIES

PATTERNS = {
    "5/95": r"5/95|Pol[ií]tica.*Ambiente",
    "16/2015": r"16/2015|saco de pl[aá]stico",
    "79/2017": r"79/2017|Embalagens",
    "94/2014": r"94/2014|Res[ií]duos S[oó]lidos Urbanos",
    "83/2014": r"83/2014|Res[ií]duos Perigosos",
    "39/2017": r"39/2017|POLMAR",
    "51/2024": r"51/2024|Fiscaliza",
    "54/2015": r"54/2015|Impacto Ambiental",
    "45/2024": r"45/2024|Auditoria Ambiental",
    "63/2024": r"63/2024|Ordenamento do Espa[cç]o Mar",
    "53/2024": r"53/2024|Economia Azul|EDEA",
    "51/2022": r"51/2022|Recifes de Coral",
    "20/97": r"20/97|Lei do Ambiente",
    "20/2019": r"20/2019|Lei do Mar",
    "97/2020": r"97/2020|Zona Costeira",
    "7/2021": r"7/2021|Desenvolvimento Territorial",
    "19/2007": r"19/2007|Ordenamento do Territ",
    "23/2008": r"23/2008|Ordenamento do Territ",
    "45/2006": r"45/2006|Marinho e Costeiro",
    "18/2004": r"18/2004|Qualidade Ambiental",
    "13/2021": r"13/2021|Sa[uú]de",
}


def expected_pattern(title_pt: str) -> str | None:
    for key, pat in PATTERNS.items():
        if key.replace("/", "") in title_pt.replace("º", "").replace("n.", ""):
            return pat
    if "ValoRe" in title_pt:
        return None
    if "EGIZC" in title_pt or "Zonas Costeiras" in title_pt:
        return None
    return None


def verify_link(url: str, pattern: str | None) -> tuple[bool, str]:
    try:
        r = requests.get(url, timeout=45, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code != 200:
            return False, f"HTTP {r.status_code}"
        ct = r.headers.get("content-type", "").lower()
        if "pdf" not in ct and not url.lower().endswith(".pdf"):
            return False, f"Not PDF ({ct})"
        if not pattern:
            return True, "PDF OK"
        reader = PdfReader(BytesIO(r.content))
        text = ""
        for page in reader.pages[:6]:
            text += page.extract_text() or ""
        if re.search(pattern, text, re.I):
            return True, "PDF content verified"
        return False, f"PDF OK but pattern not found: {pattern}"
    except Exception as exc:
        return False, str(exc)


def main() -> int:
    failures = 0
    for policy in POLICIES:
        title = policy["title_pt"]
        link = policy.get("link")
        if not link:
            print(f"SKIP (no link): {title[:60]}")
            continue
        pat = expected_pattern(title)
        ok, msg = verify_link(link, pat)
        status = "OK" if ok else "FAIL"
        print(f"{status}: {title[:55]:55} | {msg}")
        if not ok:
            failures += 1
            print(f"       {link}")
    print(f"\n{failures} failure(s)")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
