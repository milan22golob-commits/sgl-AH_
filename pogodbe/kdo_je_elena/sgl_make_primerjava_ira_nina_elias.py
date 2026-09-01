# -*- coding: utf-8 -*-
"""
sgl_make_primerjava_ira_nina_elias.py
====================================
Sinoptična primerjava treh avtorskih pogodb ŠGL:
  - Ira Ratej (Šunder v dvorani, maj 2025) — režiserka + avtorica priredbe + glasba
  - Nina Šorak (Steklena menažerija, avg 2026) — režiserka + kostumografka
  - Elias Rudolf (Kdo je Elena?, avg 2026) — dramatik + režiser

Format: LEŽEČE (landscape), 5-stolpčna tabela, po ključnih delih pogodbe.
Namen: Milan vidi razlike vzporedno + komentar tretjih stolpec (dejansko peti).

Output:
  Šentjakobsko gledališče\\Pogodbe in računi\\Kdo je Elena\\
  ŠGL_PRIMERJAVA_Ira_Nina_Elias.docx
"""

import sys
from pathlib import Path

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT_DIR = Path(r'C:\Users\ladmin\Moj disk\Šentjakobsko gledališče\Pogodbe in računi\Kdo je Elena')
OUT_DOCX = OUT_DIR / 'ŠGL_PRIMERJAVA_Ira_Nina_Elias.docx'


def apply_no_spacing(paragraph):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE


def set_landscape(doc):
    section = doc.sections[0]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_width
    section.page_height = new_height
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)


def make_grid_table(doc, rows, cols, col_widths_cm):
    table = doc.add_table(rows=rows, cols=cols)
    tbl = table._element
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    # dodaj tanke sive obrobe
    tblBorders = OxmlElement('w:tblBorders')
    for border in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{border}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '808080')
        tblBorders.append(b)
    tblPr.append(tblBorders)
    for i, w in enumerate(col_widths_cm):
        for row in table.rows:
            row.cells[i].width = Cm(w)
    return table


def fill(cell, text, size=8.5, bold=False, italic=False, color=None):
    cell.text = ''
    if not text:
        return
    lines = text.split('\n') if isinstance(text, str) else [text]
    for i, line in enumerate(lines):
        if i == 0:
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        apply_no_spacing(p)
        run = p.add_run(line)
        run.font.name = 'Calibri'
        run.font.size = Pt(size)
        run.bold = bold
        run.italic = italic
        if color is not None:
            run.font.color.rgb = RGBColor.from_string(color)


def add_title(doc, text, size=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    apply_no_spacing(p)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.name = 'Calibri'


def add_para(doc, text, size=10, italic=False):
    p = doc.add_paragraph()
    apply_no_spacing(p)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.name = 'Calibri'
    r.italic = italic


# ============================================================
# PRIMERJALNI PODATKI — po ključnih delih pogodbe
# ============================================================
# Vsak zapis: (naslov_dela, ira, nina, elias, komentar)

PRIMERJAVA = [
    (
        'UVOD STRANK — izvajalka/izvajalec',
        'ga. Ira Ratej\nnaslov: __________\nEMŠO: __________\ndavčna: __________\n(v nadaljevanju izvajalka),\n\n+ pooblastilo ARAHNE d.o.o. (256. čl. OZ) — vmesnik izda račun z DDV.',
        'ga. Nina Šorak – samozaposlena v kulturi, dramaturginja in režiserka\nTrubarjeva 26, LJ\ndavčna: 81120559\nmatična: 2484714000\nDDV: NE\nTRR: SI56 6100...\nkot režiserka in kostumografinja (izvajalka)',
        'g. Elias Rudolf, odvetnik s pisarno v LJ, Obrežna steza 2\nmatična: 2861887000\ndavčna: 34928138\nDDV: NE (76.a ZDDV-1)\nTRR: SI56 0400...\nkot dramatik in režiser (izvajalec)',
        'Ira: dvostopenjski model (F.O. + Arahne vmesnik → DDV). Nina: SVK, sama fakturira brez DDV. Elias: samostojni odvetnik, sam fakturira brez DDV (76.a).\n\nEliasova varianta je najčistejša — brez vmesnika, brez DDV, direktno.',
    ),
    (
        'NAZIV POGODBE',
        'POGODBA O SODELOVANJU',
        'AVTORSKA POGODBA',
        'AVTORSKA POGODBA',
        'Ira: "sodelovanje" ker gre skozi Arahne (pogodba o poslovnem sodelovanju s pravno osebo). Nina, Elias: direktna avtorska pogodba (F.O./SVK).',
    ),
    (
        '1. čl. — identifikacija predmeta',
        'ŠUNDER V DVORANI, po motivih Audience M. Frayna\npremiera: 28. novembra 2025\n\n+ obsežen davčno-pravni argument: "samostojno delo... zgolj inspiracija po motivih... ne priredba v smislu 7. in 33. čl. ZASP..."',
        'STEKLENA MENAŽERIJA (The Glass Menagerie) T. Williams, prev. T. Mahkota\npremiera: 18. septembra 2026\n\nBrez posebnega davčnega argumenta — prevod obstaja, avtorske pravice prevoda ureja ločeno.',
        'KDO JE ELENA?\npremiera: predvidoma marca 2027\n\n+ trditev izvirnosti: "izvirno avtorsko delo... ustvaril samostojno... ne priredba".\n\n+ dodatek Elias-specifika: "izvajalec dela opravlja preko svoje registrirane pisarne".',
        'Ira: mora braniti pravno pozicijo "ni priredba" (Frayn kot vir).\nNina: nihče ne braniti — prevod je licenca.\nElias: NAJČIŠČE — izvirno delo, brez tuje predloge.\n\nEliasov dodatek o "registrirani pisarni" je NOV — pojasni dvojno identiteto (odvetnik + umetnik). Sklic na 16. čl. ZOdv v v5 izpuščen (odveč).',
    ),
    (
        '3. čl. — predmet pogodbe (vloge)',
        'TRI vloge:\n– avtorstvo dramskega dela ŠUNDER V DVORANI\n– režijo uprizoritve\n– avtorstvo glasbene opreme',
        'DVE vlogi (obe po 5. čl. ZASP):\n– režijo uprizoritve\n– kostumografijo uprizoritve\n\n+ opomba: brez izvirne glasbe, avtor glasbe ni angažiran.',
        'DVE vlogi (obe po 5. čl. ZASP):\n– avtorstvo izvirnega dramskega dela KDO JE ELENA?\n– režijo uprizoritve',
        'Ira: 3 vloge, dvojne ZASP navedbe (5. + 80. čl.) — v v7 REVIEW naznano kot NETOČNO (80. čl. je za prenos, ne za nastanek).\nNina: pravilno oba pod 5. čl.\nElias: sledi Nini — 5. čl. za obe. Popravljen glede v1_PREDLOG (kjer sem imel 80. čl. za režijo).',
    ),
    (
        '4. čl. — umetniška odgovornost',
        '"Izvajalka je odgovorna za umetniško vrednost... vodenje igralske ekipe in koordinacijo dela soustvarjalcev."\n\n"Pri izbiri soustvarjalcev je izvajalka DOLŽNA PREDNOSTNO angažirati pri naročniku redno zaposlene ali pogodbeno že angažirane sodelavce."',
        '(vključeno v 4. čl.) "odgovorna za umetniško vrednost... vodenje igralske ekipe in koordinacijo."\n\n(Klavzula o soustvarjalcih obravnavana v 26. čl. "ne bo zahtevala drugih soustvarjalcev; ekipa se spremeni le s soglasjem".)',
        '"odgovoren za umetniško vrednost... prevzame izvedbo idejne zasnove uprizoritve in KOORDINACIJO dela soustvarjalcev."\n\n"Pri izbiri soustvarjalcev in pri finančnih zahtevah izvajalec SODELUJE z direktorjem."',
        'Ira: najstrožja dikcija ("dolžna prednostno") — DP-pravnik svari (indic odvisnega razmerja).\nNina: mehčano (samo o spremembi ekipe).\nElias: MEHČANO po DP-priporočilu — "koordinacija" namesto "vodenje", "sodeluje" namesto "prednostno angažira". Za samostojnega odvetnika kot izvajalca zdaj primerno.',
    ),
    (
        '5. čl. — honorar / narava razmerja',
        '5.500,00 EUR BRUTO (znesek vključuje DDV)\nPlačilo do 12. maja 2026 na TRR Arahne d.o.o.\n\n(!) "bruto" pri B2B je terminološka napaka po v7 §5.2 — v Ira ostalo iz starega časa.',
        'PRODUKCIJSKI OKVIR 6.000,00 EUR — izvajalka samostojno razporeja med materialne stroške in svoj honorar.\nMaterialne stroške krije naročnik neposredno.\nZaključni obračun 15 dni po premieri; plačilo v 2 delih (30 + 60 dni).',
        '2.300,00 EUR skupaj, razdeljen:\n– besedilo: 600,00 EUR\n– režija: 1.700,00 EUR\n\nDDV se ne obračuna (76.a ZDDV-1).\nIzvajalec izda račun po premieri; plačilo v 30 dneh.',
        'TRI RAZLIČNI MODELI:\n– Ira: enoten "bruto" znesek preko vmesnika z DDV\n– Nina: produkcijski okvir (kostumografka potrebuje materiale)\n– Elias: fiksen honorar razdeljen po vlogah\n\nEliasov v5 uvaja razdelitev 600/1.700 po DP-priporočilu (davčna varovalka pri sporu o "primernem honorarju" ZASP 81.).',
    ),
    (
        '6. čl. — prenos MAP (obseg pravic)',
        'Splošna navedba: "vse svoje MAP (ZASP 21.–31. čl.) na vseh treh delih"\nnašteje 4 pravice opisno:\n– javna priobčitev\n– reproduciranje\n– distribuiranje\n– snemanje',
        'Naštete pravice OPISNO:\n– javna priobčitev\n– reproduciranje za tehn. posnetke\n– distribuiranje fizičnih nosilcev\n– snemanje\n– uporaba fotografij in posnetkov',
        'Naštete pravice s SKLICI NA ZASP ČLENE:\n– reproduciranje (23. čl.)\n– distribuiranje (24. čl.)\n– javna izvedba (26. čl.)\n– radiodifuzija (30. čl.)\n– dajanje na voljo (32.a čl. — splet)\n– uporaba foto/posnetkov',
        'Ira, Nina: opisno (zadošča praksi, a manj varno po 75. čl. ZASP).\nElias v5: SPECIFIČNI SKLICI po AP-priporočilu (zadošča 75. čl. — "kar ni izrecno preneseno, ostane pri avtorju"). Najstrožja tehnika.',
    ),
    (
        '7. čl. — trajanje pravic + reverzija',
        '"Prenos velja za čas, dokler je predstava na rednem repertoarju ŠGL, oziroma dokler UO predstave formalno ne umakne."\n\nBrez reverzije.',
        '"Pravice za žive izvedbe VELJAJO, dokler je na repertoarju... za dokumentiranje/arhiv in že nastale foto/posnetke se prenesejo TRAJNO IN NEIZKLJUČNO."\n\nBrez reverzije.',
        'Enako kot Nina (živo + trajno za arhiv), PLUS eksplicitna REVERZIJA:\n"Z umikom z repertoarja MAP za žive izvedbe v celoti preidejo NAZAJ na izvajalca, brez dodatnega pravnega dejanja."',
        'Ira: samo trajanje živih pravic.\nNina: doda trajni prenos za arhiv/promocijo (koristno).\nElias: doda tudi REVERZIJO (najmočnejša zaščita avtorja po ZASP 78./83.). Priporočilo AP-pravnika v1 REVIEW.',
    ),
    (
        '8. čl. — varovalka avtorja besedila',
        '— NE VSEBUJE (ni relevantno; Ira ni izvirni dramatik).',
        '— NE VSEBUJE (Nina ni avtorica besedila).',
        '"Izvajalec OSTANE avtor in imetnik avtorskih pravic na dramskem besedilu KDO JE ELENA? kot samostojnem literarnem delu. Prenos se nanaša izključno na uprizarjanje v produkciji naročnika in ne omejuje pravice izvajalca, da besedilo objavi, ponudi drugim gledališčem..."',
        'Elias-specifika (edinstvena med tremi). Ker je Elias AVTOR IZVIRNEGA BESEDILA, mora imeti eksplicitno zaščito, da lahko besedilo uporablja tudi drugod (objava, druga uprizoritev). Sicer bi 75. čl. ZASP interpretirali v korist ŠGL zaradi ozke razlage.\n\nODPRTO VPRAŠANJE: ekskluzivnost drugim SLO gledališčem v času repertoarja — v5 dovoljuje.',
    ),
    (
        '9. čl. — moralne pravice + navedba vloge',
        'Navedba: "Ira Ratej — avtorica, režiserka in avtorica glasbene opreme"\n\n+ trojni sklop:\n1) izvajalka obdrži pravice\n2) naročnik navede ime + vlogo\n3) naročnik se OBVEZUJE spoštovati moralne pravice (didaktična trojna struktura po v7 §5.5)',
        'Navedba: "Nina Šorak — režija in kostumografija"\n\n+ posebnost: "izvajalka soglaša z rednimi PRILAGODITVAMI kostumov (menjava zasedbe, vskok, obraba) — se ne štejejo za poseg v pravico do celovitosti."',
        'Navedba: "Elias Rudolf — dramatik in režiser"\n\nBrez posebnosti glede rednih prilagoditev (pri drami ni obrabe kostumov).',
        'Ira: "avtorica..." — dolga sintagma s tremi vlogami.\nNina: klasična "režija in kostumografija".\nElias: "dramatik in režiser" — utečena slovenska sintagma (namesto "avtor besedila", ki ni v ŠGL slovarju).\n\nNina ima specifiko za kostume (redne prilagoditve). Elias tega ne potrebuje.',
    ),
    (
        '10. čl. — fizični nosilci',
        '"Fizični in digitalni nosilci, ki jih je izvajalka IZROČILA naročniku za NAMEN IZVEDBE PREDSTAVE (rokopis, beleške, zvočne datoteke, video vaj) ostanejo v hrambi naročnika... Izvirniki ostanejo v lasti izvajalke."',
        '(dva ločena člena)\n18. čl.: "za dokumentacijski arhiv IZROČI KOPIJE osnutkov in skic (po lastni izbiri: fotografske, digitalne, fotokopije). Izvirniki ostanejo v njeni lasti."\n\n(Nina ima ločeno klavzulo o kostumih = last naročnika po umiku.)',
        'ELIAS v5 (popravljeno vs v4): "za dokumentacijski arhiv naročnika IZROČI KOPIJE dramskega besedila v pisni ali digitalni obliki. Po lastni izbiri lahko izroči kopije režijske knjige/beležk. Izvirniki ostanejo v lasti izvajalca; izročitev kopij ne posega v avtorske pravice."',
        'Ira: klasična dikcija (izvajalka je za IZVEDBO izročila materiale, ki so v hrambi).\nNina: ločeno — arhivske kopije (izvajalka odloči obliko).\n\nEliasov v4 je napačno prevzel Ira/Nina logiko izvedbenega izročanja. v5 popravljen po Nini MG 18. čl. — ARHIVSKA logika: naročnik prosi za kopije, ne zahteva izvirnikov za izvedbo. To je Milanova pravilna korekcija.',
    ),
    (
        '11. čl. — jamstva',
        '"je EDINA AVTORICA vseh treh avtorskih del iz 3. čl."\n"je dramsko delo samostojno... ne priredba Fraynovega besedila v smislu 7. in 33. čl. ZASP" (obramba pred Frayn)\n"razpolaga z vsemi pooblastili za prenos"',
        '"je AVTORICA del... da so izvirna in da niso kršene pravice tretjih"\n"na delih ni pravic tretjih oseb, ki bi omejevale prenos"\n+ "v uprizoritvi ne bo uporabljeno glasbeno delo tretjih..."',
        'ELIAS v5 (razdeljeno vs v4):\n– "je IZKLJUČNI AVTOR DRAMSKEGA BESEDILA... samostojno ustvarjeno, ne priredba"\n– "se ZAVEZUJE, da bo režijo v celoti opravil sam kot izključni avtor režije, BREZ SOREŽISERJEV"\n– razpolaga s pooblastili\n\n+ opomba: jamstva ne pokrivajo sorodnih pravic tretjih (igralci 118. ZASP, glasba, scenograf).',
        'Ira: enotno jamstvo za obe/vse tri vloge (že režirala, torej "je edina").\nNina: enotno + posebnost (brez glasbe tretjih).\nElias v5 (po Milanovi opombi): RAZDELJENO — besedilo JE (napisano), režija bo (zaveza za prihodnost). Bolj natančno pravno stanje.\n\nDodana opomba o sorodnih pravicah tretjih po AP-priporočilu (v1 REVIEW).',
    ),
    (
        '12. čl. — obveznosti izvajalca / razpored',
        '"Razpored vaj je določen s tedenskim programom gledališča, ki ga celotni ekipi projekta POSREDUJE ORGANIZATOR KULTURNEGA PROGRAMA naročnika."\n\n"Izvajalka izrecno izjavlja, da v času izvrševanja NE BO SPREJELA DRUGIH ANGAŽMAJEV, ki bi jo ovirali."',
        '"Razpored vaj DOLOČA naročnik v dogovoru z izvajalko."\n\nBrez omembe organizatorja. Brez klavzule o drugih angažmajih.',
        'ELIAS v5 (kombinirano):\n"Terminski plan in razpored vaj KOORDINIRA ORGANIZATOR KULTURNEGA PROGRAMA naročnika, SPORAZUMNO Z IZVAJALCEM, ob upoštevanju splošnega tedenskega programa in razpoložljivosti soustvarjalcev."',
        'Ira: najstrožja — enostransko določa naročnik + prepoved drugih angažmajev (indici delovnega razmerja).\nNina: v dogovoru z izvajalko (mehčano).\nElias v5: KOORDINIRA organizator SPORAZUMNO — vrnjena Milanova omemba organizatorja iz Ira, a mehčano. Manj podrejenostna dikcija, primernejša za samostojnega odvetnika.',
    ),
    (
        '16. čl. — obveščanje o odsotnosti',
        '"sporočila INŠPICIENTU PREDSTAVE ALI ORGANIZATORJU KULTURNEGA PROGRAMA naročnika TAKOJ, ko je to mogoče (tel/SMS/mail)"',
        'Brez posebnega člena (splošno v 27. čl. — "nemudoma obvesti naročnika").',
        'ELIAS v5: "sporoči NAROČNIKU (organizatorju kulturnega programa naročnika) TAKOJ, ko je to mogoče"',
        'Ira: "inšpicient TAKOJ" — element delovne discipline (DP-svari).\nNina: samo naročnik.\nElias v5: kompromis — naročnik primarno, organizator kot operativni kontakt. Manj pravno tvegana dikcija.',
    ),
    (
        'PODPISNI BLOK',
        'Datum: 30. maj 2025\nIzvajalka: | Naročnik:\nŠentjakobsko gled. LJ | (prazno)\nIra Ratej | Milan Golob\navtorica, režiserka... | direktor\nŽig: | (prazno)\n\n(Ira: 5 vrstic, samo LEVI žig — Arahne nima svojega, Ira sama nima)',
        'Datum: 5. avgust 2026\nIzvajalka: | Naročnik:\nŠGL Ljubljana - društvo\nNina Šorak | Milan Golob\nrežiserka + kostumografinja | direktor\nŽig: | (prazno)\n\n(Nina: F.O./SVK — leva prazna; le desni žig)',
        'ELIAS v5 po v7 §4.0:\nDatum: 1. september 2026\nIzvajalec: | Naročnik:\n(prazno) | Šentjakobsko gled. - društvo (BOLD)\nElias Rudolf (BOLD) | Milan Golob (BOLD)\ndramatik in režiser | direktor\nŽig: | Žig:\n(oba žiga — odvetnik ima žig)',
        'Ira: ročno grajen, imena NE BOLD, žig samo levo.\nNina: podobno.\n\nElias v5 uporabi v7 §4.0 postavitev (Milanov popravek 20. 8.):\n– vrstica 4: naziv pravne osebe (BOLD desno)\n– VRSTICA 5: IMENA BOLD (obe)\n– vrstica 6: funkcija\n– vrstica 8: ŽIG OBOJESTRANSKO (odvetnik ima žig!)\n\nStandardno pravilo, ki v Ira/Nina še ni bilo dosledno uveljavljeno.',
    ),
    (
        'IZVODI',
        '2 izvoda (1 naročnik + 1 izvajalka)',
        '3 izvoda (1 izvajalka + 2 naročnik)',
        '3 izvoda (1 izvajalec + 2 naročnik)',
        'Ira: staro pravilo 2 izvoda.\nNina, Elias: novejše 3 izvoda (verjetno po standardu, ker naročnik potrebuje 2 — arhiv in računovodstvo).',
    ),
]


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = Document()
    set_landscape(doc)

    # naslov
    add_title(doc, 'PRIMERJAVA POGODB: Ira Ratej / Nina Šorak / Elias Rudolf', size=14)
    add_para(doc, '(sinoptična analiza po ključnih delih pogodbe — komentar v petem stolpcu)', size=10, italic=True)
    add_para(doc, '', size=8)

    # legenda
    add_para(doc, 'Ira Ratej — POGODBA O SODELOVANJU, "Šunder v dvorani" (maj 2025), tri avtorske vloge, plačilo preko Arahne d.o.o.', size=9, italic=True)
    add_para(doc, 'Nina Šorak — AVTORSKA POGODBA, "Steklena menažerija" (avg 2026), režija + kostumografija, SVK samostojna izda račun', size=9, italic=True)
    add_para(doc, 'Elias Rudolf v5 — AVTORSKA POGODBA, "Kdo je Elena?" (avg 2026), dramatik + režiser, samostojni odvetnik izda račun', size=9, italic=True)
    add_para(doc, '', size=8)

    # tabela
    # ležeči A4: 29,7 × 21 cm; z robovi 1,5 cm × 2 = 26,7 cm širine
    # 5 stolpcev: 4,2 | 5,5 | 5,5 | 5,5 | 6,0 = 26,7
    cols_w = [4.2, 5.5, 5.5, 5.5, 6.0]
    n_rows = len(PRIMERJAVA) + 1
    table = make_grid_table(doc, n_rows, 5, cols_w)

    # header
    header = ['DEL POGODBE', 'IRA (Šunder v dvorani)', 'NINA (Steklena menažerija)', 'ELIAS v5 (Kdo je Elena?)', 'KOMENTAR']
    for i, h in enumerate(header):
        fill(table.rows[0].cells[i], h, size=10, bold=True)

    # data rows
    for row_idx, (naslov, ira, nina, elias, kom) in enumerate(PRIMERJAVA, start=1):
        fill(table.rows[row_idx].cells[0], naslov, size=9, bold=True)
        fill(table.rows[row_idx].cells[1], ira, size=8.5)
        fill(table.rows[row_idx].cells[2], nina, size=8.5)
        fill(table.rows[row_idx].cells[3], elias, size=8.5)
        fill(table.rows[row_idx].cells[4], kom, size=8.5, italic=True, color='2E7D32')

    # zaključek
    doc.add_paragraph()
    add_title(doc, 'POVZETEK RAZLIK', size=12)
    add_para(doc, '', size=8)

    povzetek = [
        ('Zakaj se Elias RAZLIKUJE od Nine (kljub isti strukturi):',
         '1. Elias je AVTOR IZVIRNEGA BESEDILA — Nina ni (režija + kostumi). Zato Elias potrebuje varovalko avtorja (8. čl.) in razdeljena jamstva (11. čl. besedilo/režija).\n'
         '2. Elias je samostojni ODVETNIK — Nina je SVK. Različen davčni/regulatorni okvir (ZOdv 16. čl. za Eliasa).\n'
         '3. Elias nima materialov za produkcijo — Nina ima kostume. Elias 10. čl. je ARHIVSKI (naročnik prosi kopije), Nina je PRODUKCIJSKI (naročnik hrani, dokler traja).\n'
         '4. Elias ima FIKSEN honorar razdeljen po vlogah; Nina ima PRODUKCIJSKI OKVIR (materiali + honorar).'),
        ('Zakaj se Elias RAZLIKUJE od Ire:',
         '1. Ira je preko VMESNIKA (Arahne d.o.o.) → DDV; Elias direktno kot samostojni odvetnik brez DDV (76.a).\n'
         '2. Ira je POGODBA O SODELOVANJU (B2B), Elias AVTORSKA POGODBA (F.O.).\n'
         '3. Ira ima 3 vloge (avtorstvo priredbe + režija + glasba); Elias 2 (avtorstvo izvirnika + režija).\n'
         '4. Ira mora braniti "ni priredba Frayna"; Elias trdi izvirnost brez tuje predloge.\n'
         '5. Ira nima varovalke avtorja — Elias ima (8. čl.), ker gre za izvirno besedilo, ki ga bo lahko uporabljal tudi zunaj ŠGL.\n'
         '6. Ira ima 2 izvoda pogodbe; Elias 3 (novejši standard).'),
        ('Skupne poteze:',
         '– Vse tri sledijo v7 9-poglavni strukturi.\n'
         '– Vse tri ohranijo moralne pravice avtorju (V. poglavje).\n'
         '– Vse tri dopuščajo neomejeno število ponovitev brez dodatnega honorarja.\n'
         '– Vse tri imajo klavzulo o razdrtju s strani naročnika (materialna škoda / ugled) in izvajalca (če predstava ne pride do premiere).'),
        ('KJE JE ELIAS v5 NAJDLJE PRED IRO/NINO:',
         '– Popravljeni ZASP sklici (Ira ima 80. čl. tudi za nastanek režije — netočno).\n'
         '– Specifični sklici na členi ZASP pri prenosu MAP (Ira, Nina navedeta opisno).\n'
         '– Eksplicitna reverzijska klavzula (ne pri Iri, ne pri Nini).\n'
         '– Varovalka avtorja besedila (Elias-specifika, ni pri drugih).\n'
         '– Podpisni blok po v7 §4.0 (Ira/Nina imata staro, nepravilno strukturo — imena niso bold).\n'
         '– Razdeljen honorar po vlogah (DP-priporočilo, ni pri drugih).\n'
         '– Mehčani delovno-pravni indici (Ira ima najstrožje, DP-tveganje rekvalifikacije).'),
    ]
    for naslov, telo in povzetek:
        p = doc.add_paragraph()
        apply_no_spacing(p)
        r = p.add_run(naslov)
        r.bold = True
        r.font.size = Pt(10)
        r.font.name = 'Calibri'
        r.font.color.rgb = RGBColor.from_string('1F4E79')
        p2 = doc.add_paragraph()
        apply_no_spacing(p2)
        r2 = p2.add_run(telo)
        r2.font.size = Pt(9.5)
        r2.font.name = 'Calibri'
        doc.add_paragraph()

    doc.save(str(OUT_DOCX))
    print(f'GENERIRANO: {OUT_DOCX}')
    print(f'Velikost: {OUT_DOCX.stat().st_size:,} B')
    print(f'Vrstic v tabeli: {len(PRIMERJAVA)}')


if __name__ == '__main__':
    main()
