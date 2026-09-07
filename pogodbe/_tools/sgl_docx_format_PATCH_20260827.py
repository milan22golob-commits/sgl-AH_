"""
sgl_docx_format_PATCH_20260827.py
==================================
POPRAVKI za sgl_docx_format.py (Milanov Windows lokalni izvod) po
lekcijah iz Elias v5 (27. 8. 2026).

Aplikacija: v Milanovi obstoječi sgl_docx_format.py (C:\\Users\\ladmin\\Moj disk\\KODE\\_tools\\)
    (1) NADOMESTI obstoječo funkcijo add_bullet z verzijo spodaj
    (2) DODAJ novo funkcijo set_metadata_milan (na koncu datoteke)
    (3) V vsakem generatorju pogodb dodaj po init_doc():
          set_metadata_milan(doc)

Standard v7 dopolnitve (predlog za §5.6.10 popravek in §5.6.12 novo):

§5.6.10 (popravek 2026-08-27, po Milanovi opombi o preglednosti):
    Alineje z ZAMIKOM:
      - levi 1,25 cm  (odmik od glavnega roba)
      - viseči −0,50 cm  (pomišljaj pri 0,75, besedilo pri 1,25)
      - wrap pod besedilom
    Prej (20. 8.): levi 0,75 / viseči −0,75 → alineje niso izstopale.

§5.6.12 (novo 2026-08-27):
    Metadata author = "Milan Golob" v vseh pogodbah in izjavah ŠGL
    (Word File → Info → Properties). Enotno preko helperja.
"""

from docx.shared import Pt, Cm
# Predpostavlja, da so v sgl_docx_format.py že definirane:
#   apply_no_spacing, set_run_calibri


# --------- NADOMESTI obstoječo add_bullet s tem: ---------

def add_bullet(doc, text):
    """Alineja po v7 §5.6.10 (popravek 2026-08-27 — Milan): pomišljaj + presledek + ZAMIK.

    Zamiki:
      - levi 1,25 cm
      - viseči −0,50 cm  → pomišljaj pri 0,75 cm, besedilo pri 1,25 cm
      - wrap pod besedilom
    """
    p = doc.add_paragraph()
    apply_no_spacing(p)
    pf = p.paragraph_format
    pf.left_indent = Cm(1.25)
    pf.first_line_indent = Cm(-0.50)
    run = p.add_run(f'– {text}')
    set_run_calibri(run, size=11)
    return p


# --------- DODAJ na konec sgl_docx_format.py: ---------

def set_metadata_milan(doc):
    """Nastavi .docx metadata author = 'Milan Golob' (v7 §5.6.12).

    Klic OBVEZEN v vseh generatorjih pogodb ŠGL takoj po init_doc().
    Rezultat: Word File → Info → Properties → Author = Milan Golob
    (ne privzeta prazna vrednost ali "Python-docx").
    """
    doc.core_properties.author = 'Milan Golob'
    doc.core_properties.last_modified_by = 'Milan Golob'
