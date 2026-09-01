# -*- coding: utf-8 -*-
"""
sgl_make_pogodba_elias_v5_LECA_CITATI.py
========================================
LEČNA verzija v5: za VSAK sklic na zakonski člen v pogodbi doda pod besedilo
VERBATIM CITAT tega člena. Namen: Milan vidi, ali argumentacija stoji ali pade
pri neposredni primerjavi s tekstom zakona.

Barvna semantika:
  - MODRA (1F4E79)  — verbatim citat zakona (pravna podlaga)
  - ORANŽNA (E67E22) — GLASNO OPOZORILO: napačen sklic ali dvomen sklic
  - VIJOLIČNA (8B008B) — pojasnilo, zakaj citat podpira (ali ne) klavzulo

Preverjeni citati (WebSearch, 27. 8. 2026):
  - ZASP 5. čl. — Varovana dela
  - ZASP 23. čl. — Pravica reproduciranja
  - ZASP 24. čl. — Pravica distribuiranja
  - ZASP 26. čl. — Pravica javnega izvajanja
  - ZASP 30. čl. — Pravica radiodifuznega oddajanja
  - ZASP 32.a čl. — Pravica dajanja na voljo javnosti
  - ZASP 118. čl. — Izvajalci (definicija — začetek sorodnih pravic)
  - ZOdv 21. čl. — Nezdružljivost (POPRAVLJENO — v v3/v4 napačno navedeno 16. čl.)

NEDOKAZANE / DVOMLJIVE reference (POTREBNO PREVERITI PRI PRAVNIKU):
  - ZDDV-1 76.a čl. — v v5 kot razlog za NE-obračun DDV, ampak 76.a se nanaša
    na REVERSE CHARGE (ne na oprostitev). Pravilnejši sklic je verjetno 94.
    čl. (mali davčni zavezanec) ali 44. čl. (oprostitev za avtorska dela).

Output:
  Šentjakobsko gledališče\\Pogodbe in računi\\Kdo je Elena\\
  ŠGL_avtorska pogodba_KDO JE ELENA_Elias Rudolf_v5_LECA-CITATI.docx
"""

import sys
import shutil
from pathlib import Path

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

sys.path.insert(0, r'C:\Users\ladmin\Moj disk\KODE\_tools')
from sgl_docx_format import (
    apply_no_spacing, add_blank_para, set_margins,
    add_doc_title, add_section, add_clen, add_para_runs, add_bullet,
    add_footer_pagenum,
    make_table_no_borders, fill_cell,
)

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

OUT_DIR = Path(r'C:\Users\ladmin\Moj disk\Šentjakobsko gledališče\Pogodbe in računi\Kdo je Elena')
REF_DOCX = Path(r'C:\Users\ladmin\Moj disk\Šentjakobsko gledališče\Pogodbe in računi\Šunder v dvorani\ŠGL_avtorska pogodba_ŠUNDER_Maša Milčinski_lektura_v3_CLEAN.docx')
OUT_DOCX = OUT_DIR / 'ŠGL_avtorska pogodba_KDO JE ELENA_Elias Rudolf_v5_LECA-CITATI.docx'

COL_CITAT = '1F4E79'  # modra — verbatim citat
COL_WARN = 'E67E22'   # oranžna — opozorilo
COL_ARG = '8B008B'    # vijolična — argumentacija

# Verbatim citati iz WebSearch (27. 8. 2026, viri: PISRS, zakonodaja.com, ZAPS)
CITATI = {
    'ZASP_5': (
        '5. člen ZASP — Varovana dela',
        '(1) Avtorska dela so individualne intelektualne stvaritve s področja književnosti, znanosti in umetnosti, ki so na kakršen koli način izražene, če ni s tem zakonom drugače določeno.\n'
        '(2) Za avtorska dela veljajo zlasti: govorjena dela, kot npr. govori, pridige, predavanja; pisana dela, kot npr. leposlovna dela, članki, priročniki, študije ter računalniški programi; glasbena dela z besedilom ali brez besedila; gledališka, gledališko-glasbena in lutkovna dela; koreografska in pantomimska dela; fotografska dela in dela, narejena po postopku, podobnem fotografiranju; avdiovizualna dela; likovna dela, kot npr. slike, grafike in kipi; arhitekturna dela, kot npr. skice, načrti ter izvedeni objekti s področja arhitekture, urbanizma in krajinske arhitekture; dela uporabne umetnosti in industrijskega oblikovanja; kartografska dela; predstavitve znanstvene, izobraževalne ali tehnične narave (tehnične risbe, načrti, skice, tabele, izvedenska mnenja, plastične predstavitve in druga dela enake narave).',
        'Vir: pisrs.si (ZASP-UPB3), zakonodaja.com (ZASP-NPB11)',
    ),
    'ZASP_23': (
        '23. člen ZASP — Pravica reproduciranja',
        '(1) Pravica reproduciranja je izključna pravica, da se delo fiksira na materialnem nosilcu ali drugem primerku, in sicer neposredno ali posredno, začasno ali trajno, delno ali v celoti ter s kakršnimkoli sredstvom ali v katerikoli obliki.\n'
        '(2) Delo se reproducira zlasti v obliki grafičnega razmnoževanja, tridimenzionalnega razmnoževanja, zvočnega ali vizualnega snemanja, gradnje oziroma izvedbe arhitekturnega objekta, shranitve v elektronski obliki.',
        'Vir: zakonodaja.com/zakon/zasp/23-clen-pravica-reproduciranja',
    ),
    'ZASP_24': (
        '24. člen ZASP — Pravica distribuiranja',
        'Pravica distribuiranja je izključna pravica, da se izvirnik ali primerki dela dajo v promet s prodajo ali drugačnim prenosom lastninske pravice ali s ponudbo v ta namen za javnost.',
        'Vir: zakonodaja.com/zakon/zasp',
    ),
    'ZASP_26': (
        '26. člen ZASP — Pravica javnega izvajanja',
        'Pravica javnega izvajanja obsega izključne pravice, da se: (i) delo književnosti prenaša javnosti z živo izvedbo (pravica javnega recitiranja); (ii) glasbeno delo prenaša javnosti z živo izvedbo (pravica javnega glasbenega izvajanja); (iii) delo prenaša javnosti z odrsko izvedbo (pravica javnega uprizarjanja).',
        'Vir: zakonodaja.com/zakon/zasp; ZAPS povzetek 22. čl. ZASP',
    ),
    'ZASP_30': (
        '30. člen ZASP — Pravica radiodifuznega oddajanja',
        'Pravica radiodifuznega oddajanja je izključna pravica, da se delo prenaša javnosti z radijskimi ali televizijskimi programskimi signali, namenjenimi javnosti, žično ali brezžično, vključno prek satelita ali po kablu ali podobnem sistemu.',
        'Vir: zakonodaja.com/zakon/zasp',
    ),
    'ZASP_32A': (
        '32.a člen ZASP — Pravica dajanja na voljo javnosti',
        'Pravica dajanja na voljo javnosti je izključna pravica, da se delo po žici ali brezžično naredi dostopno javnosti na način, ki omogoča posameznikom iz javnosti dostop do njega s kraja in v času, ki ju sami izberejo.',
        'Vir: zakonodaja.com/zakon/zasp/32a-clen-pravica-dajanja-na-voljo-javnosti',
    ),
    'ZASP_118': (
        '118. člen ZASP — Izvajalci (definicija, začetek 5. poglavja Sorodne pravice)',
        'Izvajalci so igralci, pevci, glasbeniki, plesalci in druge osebe, ki igrajo, pojejo, podajajo, deklamirajo, nastopajo, interpretirajo ali kako drugače izvajajo avtorska ali folklorna dela. Kot izvajalci se štejejo tudi režiserji gledaliških predstav, dirigenti orkestrov, vodje pevskih zborov, oblikovalci tona ter varietejski in cirkuški umetniki.',
        'Vir: zakonodaja.com/zakon/zasp/1-oddelek-pravice-izvajalcev',
    ),
    'ZOdv_21': (
        '21. člen ZOdv — Nezdružljivost',
        'Z opravljanjem odvetniškega poklica je nezdružljivo:\n'
        '1. opravljanje druge dejavnosti kot poklica, razen v znanstveni, pedagoški, umetniški ali publicistični dejavnosti;\n'
        '2. opravljanje plačane državne službe;\n'
        '3. opravljanje notariata;\n'
        '4. opravljanje vodstvene funkcije v gospodarskem subjektu ali drugi pravni osebi, ki ni zajeta v 35. členu tega zakona.',
        'Vir: pisrs.si (ZAKO265); povzetek na zakonodaja.com/zakon/zodv/21-clen',
    ),
}


def init_doc():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if not REF_DOCX.exists():
        return Document(), None
    tmp = OUT_DIR / '_tmp_elias_leca.docx'
    shutil.copy(str(REF_DOCX), str(tmp))
    doc = Document(str(tmp))
    body = doc.element.body
    sectPr = None
    for child in list(body):
        if child.tag == qn('w:sectPr'):
            sectPr = child
            continue
        body.remove(child)
    if sectPr is not None:
        body.append(sectPr)
    return doc, tmp


def citat_line(doc, key):
    """Modra vrstica z verbatim citatom zakona."""
    naslov, telo, vir = CITATI[key]
    p = doc.add_paragraph()
    apply_no_spacing(p)
    r1 = p.add_run(f'📖 CITAT — {naslov}: ')
    r1.bold = True
    r1.italic = True
    r1.font.size = Pt(8.5)
    r1.font.name = 'Calibri'
    r1.font.color.rgb = RGBColor.from_string(COL_CITAT)
    r2 = p.add_run(f'„{telo}"')
    r2.italic = True
    r2.font.size = Pt(8.5)
    r2.font.name = 'Calibri'
    r2.font.color.rgb = RGBColor.from_string(COL_CITAT)
    p2 = doc.add_paragraph()
    apply_no_spacing(p2)
    r3 = p2.add_run(f'    ({vir})')
    r3.italic = True
    r3.font.size = Pt(7.5)
    r3.font.name = 'Calibri'
    r3.font.color.rgb = RGBColor.from_string(COL_CITAT)


def arg_line(doc, besedilo):
    """Vijolična vrstica: pojasnilo, zakaj citat podpira klavzulo."""
    p = doc.add_paragraph()
    apply_no_spacing(p)
    r1 = p.add_run('▸ ARG: ')
    r1.bold = True
    r1.italic = True
    r1.font.size = Pt(8.5)
    r1.font.name = 'Calibri'
    r1.font.color.rgb = RGBColor.from_string(COL_ARG)
    r2 = p.add_run(besedilo)
    r2.italic = True
    r2.font.size = Pt(8.5)
    r2.font.name = 'Calibri'
    r2.font.color.rgb = RGBColor.from_string(COL_ARG)


def warn_block(doc, naslov, telo):
    """Velik oranžni blok — opozorilo."""
    doc.add_paragraph()
    p1 = doc.add_paragraph()
    apply_no_spacing(p1)
    r = p1.add_run(f'⚠⚠⚠  {naslov}  ⚠⚠⚠')
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.name = 'Calibri'
    r.font.color.rgb = RGBColor.from_string(COL_WARN)
    p2 = doc.add_paragraph()
    apply_no_spacing(p2)
    r2 = p2.add_run(telo)
    r2.font.size = Pt(9.5)
    r2.font.name = 'Calibri'
    r2.font.color.rgb = RGBColor.from_string(COL_WARN)
    doc.add_paragraph()


def build_signature_block_v7(doc, datum, izvajalec_naziv, izvajalec_ime, izvajalec_funkcija,
                              narocnik_naziv, narocnik_ime, narocnik_funkcija):
    doc.add_paragraph()
    doc.add_paragraph()
    table = make_table_no_borders(doc, rows=9, cols=2)

    def cell(row, col, text, bold=False):
        fill_cell(table.cell(row, col), text, bold=bold)

    cell(0, 0, f'Datum: {datum}')
    cell(0, 1, f'Datum: {datum}')
    cell(2, 0, 'Izvajalec:')
    cell(2, 1, 'Naročnik:')
    cell(4, 0, izvajalec_naziv or '', bold=bool(izvajalec_naziv))
    cell(4, 1, narocnik_naziv, bold=True)
    cell(5, 0, izvajalec_ime, bold=True)
    cell(5, 1, narocnik_ime, bold=True)
    cell(6, 0, izvajalec_funkcija)
    cell(6, 1, narocnik_funkcija)
    cell(8, 0, 'Žig:')
    cell(8, 1, 'Žig:')
    return table


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc, tmp = init_doc()

    set_margins(doc, top_cm=2.5, bottom_cm=2.5, left_cm=2.5, right_cm=2.5)
    add_footer_pagenum(doc)

    # ============================================================
    # ZGORNJA GLAVA
    # ============================================================
    p = doc.add_paragraph()
    apply_no_spacing(p)
    r = p.add_run('LEČNA VERZIJA v5 — verbatim citati zakonskih členov, na katere se sklicuje pogodba. NI za podpis.')
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Calibri'
    r.font.color.rgb = RGBColor.from_string(COL_WARN)

    for lab, opis in [
        ('CITAT', 'modra — verbatim tekst zakonskega člena (vir naveden pod citatom)'),
        ('ARG',   'vijolična — argumentacija, zakaj citat podpira (ali ne) klavzulo'),
        ('OPOZORILO', 'oranžna — napačen sklic ali dvomen sklic (velik blok)'),
    ]:
        pl = doc.add_paragraph()
        apply_no_spacing(pl)
        rl = pl.add_run(f'▸ [{lab}] = {opis}')
        rl.bold = True
        rl.font.size = Pt(8.5)
        rl.font.name = 'Calibri'
        col = {'CITAT': COL_CITAT, 'ARG': COL_ARG, 'OPOZORILO': COL_WARN}[lab]
        rl.font.color.rgb = RGBColor.from_string(col)
    doc.add_paragraph()

    # OPOZORILO — pri 76.a čl. ZDDV-1
    warn_block(
        doc,
        'DVOMEN SKLIC: 76.a čl. ZDDV-1 v uvodu strank in 5. čl. pogodbe',
        '76.a čl. ZDDV-1 ureja "obrnjeno davčno obveznost" (reverse charge — prejemnik blaga/storitve postane plačnik DDV), '
        'NE pa oprostitve za mali davčni zavezanec. Bizi.si sicer navaja "Ni plačnik DDV - 76.a člen ZDDV-1" kot '
        'klasifikacijsko oznako (Elias ima verjetno DDV ID za prejemanje storitev iz EU, pri domačih dobavah ne obračuna DDV), '
        'ampak PRAVNI RAZLOG za NE-obračun DDV pri domačih avtorskih storitvah je verjetno drug:\n'
        '  → 94. čl. ZDDV-1 (mali davčni zavezanec, prag 50.000 EUR letnega prometa), ali\n'
        '  → 44. čl. ZDDV-1 (oprostitve — točka 12: duhovne storitve pisateljev/skladateljev/izvajalcev), ali\n'
        '  → najpreprosteje brez sklica: "Izvajalec ni davčni zavezanec za DDV, zato se DDV ne obračuna."\n\n'
        'PRIPOROČILO: Elias naj (kot odvetnik) sam preveri, kateri člen ZDDV-1 pravilno pokriva njegov status pri '
        'AVTORSKI storitvi (ki je različno obravnavana od odvetniške storitve).'
    )

    # OPOZORILO — 16. čl. ZOdv v CHANGELOG-u v3/v4 je bil napačen
    warn_block(
        doc,
        'POPRAVEK NAPAKE: v v3/v4 sem navajal 16. čl. ZOdv, pravilno je 21. čl.',
        'V v3 in v4 sem v 1. členu (pojasnilo o registrirani pisarni odvetnika) sklical na 16. čl. Zakona o odvetništvu. '
        'V resnici 16. čl. ZOdv ureja STATUS SPECIALISTA ODVETNIKA (izvoljen v pedagoški naziv na pravni fakulteti) — '
        'z umetniško dejavnostjo NIMA NOBENE ZVEZE.\n\n'
        'PRAVILNI ČLEN je 21. čl. ZOdv (glej citat pri 1. členu spodaj). V v5 sem sicer sklic na ZOdv izpustil (klavzula '
        'skrajšana), ampak v CHANGELOG.md še vedno stoji napačna referenca. To bo popravljeno v v6.'
    )

    doc.add_paragraph()

    # ============================================================
    # UVOD STRANK
    # ============================================================
    add_para_runs(doc, [('Pogodbeni stranki:', False)])
    doc.add_paragraph()
    add_para_runs(doc, [
        ('Šentjakobsko gledališče Ljubljana - društvo', True),
        (', Krekov trg 2, 1000 Ljubljana, davčna številka: 31033008, matična številka: 5147689000, davčni zavezanec za DDV: NE, ki ga zastopa ', False),
        ('g. Milan Golob', True),
        (' kot direktor gledališča (v nadaljevanju naročnik),', False),
    ])
    doc.add_paragraph()
    add_para_runs(doc, [('in', False)])
    doc.add_paragraph()

    add_para_runs(doc, [
        ('g. Elias Rudolf, odvetnik', True),
        (' s pisarno v Ljubljani, Obrežna steza 2, 1000 Ljubljana,', False),
    ])
    add_para_runs(doc, [('matična številka: 2861887000', False)])
    add_para_runs(doc, [('davčna številka: 34928138', False)])
    add_para_runs(doc, [('davčni zavezanec za DDV: NE (76.a člen ZDDV-1)', False)])
    arg_line(doc, 'Sklic na 76.a je bizi klasifikacija; pravilni pravni razlog je verjetno drug (glej oranžno opozorilo zgoraj).')
    add_para_runs(doc, [('TRR: SI56 0400 0028 0838 318 (OTP banka d.d.)', False)])
    add_para_runs(doc, [('kot dramatik in režiser (v nadaljevanju izvajalec),', False)])
    doc.add_paragraph()

    add_para_runs(doc, [('skleneta naslednjo', False)])
    add_doc_title(doc, 'AVTORSKO POGODBO')

    # ============== I. UVODNE DOLOČBE ==============
    add_section(doc, 'UVODNE DOLOČBE')

    add_clen(doc)
    add_para_runs(doc, [
        ('Ta pogodba se nanaša na avtorsko delo izvajalca pri krstni uprizoritvi gledališke predstave ', False),
        ('KDO JE ELENA?', True),
        (', katere avtor dramskega besedila in režiser je izvajalec, in ki bo premierno uprizorjena ', False),
        ('predvidoma marca 2027', True),
        (' na Šentjakobskem odru Šentjakobskega gledališča Ljubljana - društvo (v nadaljevanju: uprizoritev).', False),
    ])
    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da je dramsko besedilo ', False),
        ('KDO JE ELENA?', True),
        (' izvirno avtorsko delo izvajalca v smislu 5. člena Zakona o avtorski in sorodnih pravicah (v nadaljevanju: ZASP), ki ga je izvajalec ustvaril samostojno in ne predstavlja priredbe ali predelave dela tretje osebe.', False),
    ])
    citat_line(doc, 'ZASP_5')
    arg_line(doc, 'Klavzula temelji na 5. čl. ZASP: dramsko besedilo je "individualna intelektualna stvaritev s področja književnosti" (izrecno našteto med primeri avtorskih del). Trditev izvirnosti postavlja Eliasovo besedilo v to kategorijo.')

    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da izvajalec pri tej pogodbi nastopa kot dramatik in režiser uprizoritve. Avtorstvo dramskega besedila in režijske izvedbe pripada izvajalcu kot fizični osebi v skladu z ZASP. Delo po tej pogodbi izvajalec opravlja preko svoje registrirane pisarne.', False),
    ])
    citat_line(doc, 'ZOdv_21')
    arg_line(doc, 'Klavzula o "opravljanju dela preko registrirane pisarne" je pravno mogoča zaradi 21. čl. ZOdv (1. točka — "razen v umetniški ali publicistični dejavnosti"). Sam sklic na 21. čl. v pogodbi ni potreben (Elias sam skrbi za skladnost), a je pravni temelj.')
    add_para_runs(doc, [('Pogodbeni stranki se strinjata, da so storitve, ki so predmet te pogodbe, avtorsko delo, skladno z ZASP.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Pogodbeni stranki soglašata, da se lahko datum premiere oziroma datumi ponovitev, vaj ali drugih dogodkov v zvezi z uprizoritvijo spremenijo iz razlogov višje sile ali drugih nepredvidenih okoliščin, brez odgovornosti katere koli pogodbene stranke.', False)])
    arg_line(doc, 'Splošna klavzula višje sile — brez specifičnega zakonskega sklica. Temelji na 153. čl. OZ (višja sila kot razbremenilni razlog).')

    # ============== II. PREDMET POGODBE ==============
    add_section(doc, 'PREDMET POGODBE')

    add_clen(doc)
    add_para_runs(doc, [('S to pogodbo izvajalec prevzema pri uprizoritvi dve avtorski vlogi:', False)])
    for b in [
        'avtorstvo izvirnega dramskega dela KDO JE ELENA? (avtorsko delo po 5. členu ZASP);',
        'režijo uprizoritve (avtorsko delo po 5. členu ZASP).',
    ]:
        add_bullet(doc, b)
    citat_line(doc, 'ZASP_5')
    arg_line(doc, 'Obe vlogi po 5. čl. ZASP: dramsko besedilo = "pisano/gledališko delo"; režija = "gledališko-glasbeno delo" (naštevanje v 2. odst. 5. čl.). Sklic na 80. čl. ZASP za režijo (kot v v1_PREDLOG in v Iri) je NAPAČEN — 80. čl. ureja PRENOS avtorske pogodbe, ne nastanek.')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec je odgovoren za umetniško vrednost uprizoritve in za usklajeno delovanje z direktorjem gledališča glede izvedbe sprejetega terminskega plana. Izvajalec prevzame izvedbo idejne zasnove uprizoritve in koordinacijo dela soustvarjalcev.', False)])
    add_para_runs(doc, [('Pri izbiri soustvarjalcev in pri finančnih zahtevah, povezanih z uprizoritvijo, izvajalec sodeluje z direktorjem gledališča ter upošteva finančne, tehnične in kadrovske zmožnosti naročnika.', False)])
    arg_line(doc, 'Mehčana dikcija ("koordinacija" namesto "vodenje", "sodeluje" namesto "prednostno angažira") — glede na DP-priporočilo iz REVIEW-a (14. 8.), NE na Milanovo izrecno zahtevo. Odprto vprašanje: obdržati mehčano ali vrniti Ira-dikcijo?')

    # ============== III. NARAVA RAZMERJA ==============
    add_section(doc, 'NARAVA RAZMERJA')

    add_clen(doc)
    add_para_runs(doc, [
        ('Pogodbeni stranki soglašata, da prejme izvajalec za delo po tej pogodbi, ki obsega obe avtorski vlogi iz 3. člena te pogodbe, avtorski honorar v skupni višini ', False),
        ('2.300,00 EUR', True),
        (', in sicer razdeljen po vlogah:', False),
    ])
    add_bullet(doc, 'za avtorstvo dramskega dela KDO JE ELENA?: 600,00 EUR;')
    add_bullet(doc, 'za režijo uprizoritve: 1.700,00 EUR.')
    add_para_runs(doc, [('Ker izvajalec ni davčni zavezanec za DDV (76.a člen ZDDV-1), se davek na dodano vrednost ne obračuna.', False)])
    arg_line(doc, 'Sklic na 76.a čl. ZDDV-1 je DVOMEN — glej veliko oranžno opozorilo na vrhu dokumenta. Pravilnejši razlog je verjetno 94. čl. (mali davčni zavezanec) ali 44. čl. (oprostitev za avtorska dela pisateljev), ali brez sklica.')
    add_para_runs(doc, [('S honorarjem iz prejšnjega odstavka so v celoti poravnane vse obveznosti naročnika do izvajalca iz naslova te pogodbe. V honorarju so všteti vsi materialni stroški, ki jih bo imel izvajalec v zvezi z opravljanjem avtorskega dela po tej pogodbi.', False)])
    add_para_runs(doc, [('Izvajalec izstavi naročniku račun po opravljeni premieri uprizoritve. Naročnik izplača honorar na transakcijski račun izvajalca najkasneje trideseti (30.) dan od prejema računa.', False)])

    # ============== IV. PRENOS MATERIALNIH AVTORSKIH PRAVIC ==============
    add_section(doc, 'PRENOS MATERIALNIH AVTORSKIH PRAVIC')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec odstopa naročniku za honorar, določen s to pogodbo, materialne avtorske pravice na obeh avtorskih delih iz 3. člena te pogodbe, in sicer:', False)])
    for b in [
        'na dramskem besedilu KDO JE ELENA? — za potrebe uprizarjanja te produkcije;',
        'na režijski izvedbi uprizoritve.',
    ]:
        add_bullet(doc, b)
    add_para_runs(doc, [('Prenos obsega zlasti:', False)])
    for b in [
        'pravico reproduciranja (23. člen ZASP) za tehnične posnetke, potrebne za uprizoritev;',
        'pravico distribuiranja (24. člen ZASP) fizičnih nosilcev, ki so del tehnične produkcije uprizoritve;',
        'pravico javne izvedbe in uprizoritve (26. člen ZASP);',
        'pravico radiodifuznega oddajanja (30. člen ZASP);',
        'pravico dajanja na voljo javnosti (32.a člen ZASP), zlasti za spletne prenose;',
        'pravico uporabe fotografij in posnetkov uprizoritve v informativno-promocijskih materialih.',
    ]:
        add_bullet(doc, b)
    citat_line(doc, 'ZASP_23')
    citat_line(doc, 'ZASP_24')
    citat_line(doc, 'ZASP_26')
    citat_line(doc, 'ZASP_30')
    citat_line(doc, 'ZASP_32A')
    arg_line(doc, 'Vseh 5 sklicev je konsistentnih z verbatim citati: reproduciranje = fiksiranje na materialnem nosilcu (tehnični posnetki), distribuiranje = prodaja/prenos primerkov (fizični nosilci produkcije), javna izvedba = javno uprizarjanje dela (predstave), radiodifuzija = RTV signali (TV prenos), dajanje na voljo = splet z izbiro kraja/časa (streaming). Klavzula pokriva vse relevantne rabe uprizoritve.')

    add_para_runs(doc, [('Izvajalec izrecno dovoljuje neomejeno število ponovitev brez dodatnih zahtev za plačilo honorarja:', False)])
    for b in ['na matičnem odru naročnika,', 'na gostovanjih doma in v tujini,', 'na festivalih doma in v tujini,', 'za televizijske in spletne prenose.']:
        add_bullet(doc, b)

    add_clen(doc)
    add_para_runs(doc, [('Pravice za žive izvedbe in ponovitve iz prejšnjega člena veljajo, dokler je uprizoritev na rednem repertoarju Šentjakobskega gledališča Ljubljana - društvo, oziroma dokler je upravni odbor gledališča formalno ne umakne z repertoarja. Pravice za dokumentiranje in arhiviranje ter uporabo že nastalih fotografij in posnetkov v informativno-promocijskih materialih se prenesejo trajno in neizključno.', False)])
    add_para_runs(doc, [('Z umikom uprizoritve z repertoarja prenesene materialne avtorske pravice za žive izvedbe in ponovitve v celoti preidejo nazaj na izvajalca, brez potrebe po dodatnem pravnem dejanju.', False)])
    arg_line(doc, 'Reverzijska klavzula: temelji na splošnem načelu ZASP 75.–78. čl. (ozka razlaga prenosa; kar ni izrecno preneseno, ostane pri avtorju) in 83. čl. (vračanje pravic ob neizvrševanju). Verbatim za 75.–78. in 83. čl. tukaj ni ekstrahiran; priporočilo — dodati v v6 lečno verzijo.')

    add_clen(doc)
    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da izvajalec ', False),
        ('ostane avtor in imetnik avtorskih pravic na dramskem besedilu', True),
        (' KDO JE ELENA? kot samostojnem literarnem delu. Prenos pravic iz te pogodbe se nanaša izključno na uprizarjanje tega besedila v produkciji naročnika in ne omejuje pravice izvajalca, da svoje dramsko besedilo objavi, ponudi v uprizoritev drugim gledališčem ali kako drugače uporablja zunaj te produkcije.', False),
    ])
    arg_line(doc, 'Varovalka avtorja: sledi 75. čl. ZASP (ozka razlaga prenosa) — izrecno ločuje avtorstvo besedila kot samostojnega literarnega dela od prenosa pravic za konkretno uprizoritev.')

    # ============== V. MORALNE AVTORSKE PRAVICE ==============
    add_section(doc, 'MORALNE AVTORSKE PRAVICE')

    add_clen(doc)
    add_para_runs(doc, [('Moralne avtorske pravice ostanejo izvajalcu v skladu z ZASP, zlasti pravica do priznanja avtorstva in pravica do celovitosti dela.', False)])
    add_para_runs(doc, [('Naročnik bo izvajalca v vseh informativno-promocijskih materialih, gledališkem listu in drugih javnih objavah, povezanih z uprizoritvijo, navedel s polnim imenom in priimkom — Elias Rudolf — z navedbo obeh avtorskih vlog: "dramatik in režiser".', False)])
    add_para_runs(doc, [('Izvajalec ima pravico, da se upre vsaki skazitvi, okrnitvi ali drugačni spremembi svojega dela ter pravico, da se upre vsaki izvedbi dela, ki bi žalila njegovo čast in ugled.', False)])
    add_para_runs(doc, [('Izvajalec je seznanjen, da bo njegovo delo predmet tržnega komuniciranja, ki ga za namene obveščanja javnosti izvaja naročnik.', False)])
    arg_line(doc, 'Moralne pravice: temelj v ZASP 16.–20. čl. (pravica prve objave 17., pravica priznanja avtorstva 18., pravica spoštovanja dela 19., pravica skesanja 20.). Verbatim citati za 18. in 19. čl. tukaj niso ekstrahirani — priporočilo dodati v v6.')

    # ============== VI. FIZIČNI NOSILCI IN LASTNINA ==============
    add_section(doc, 'FIZIČNI NOSILCI IN LASTNINA')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec za dokumentacijski arhiv naročnika izroči kopije dramskega besedila v pisni ali digitalni obliki. Po lastni izbiri lahko izroči tudi kopije režijske knjige, beležk ali drugih delovnih materialov, povezanih z uprizoritvijo.', False)])
    add_para_runs(doc, [('Izvirnik dramskega besedila in drugi izvirniki avtorskega dela ostanejo v lasti izvajalca. Izročitev kopij po prejšnjem odstavku ne posega v avtorske pravice izvajalca.', False)])
    arg_line(doc, 'Klavzula pravilno razlikuje LASTNIŠTVO NOSILCA (stvar — SPZ) od AVTORSKE PRAVICE (nemat. — ZASP 39. čl.). Izročitev nosilca ne prenaša pravice — pravni temelj brez posebnega sklica.')

    # ============== VII. JAMSTVA ==============
    add_section(doc, 'JAMSTVA')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec jamči naročniku, da:', False)])
    for b in [
        'je izključni avtor dramskega besedila KDO JE ELENA?, ki ga je ustvaril samostojno, in da to delo ni priredba ali predelava dela tretje osebe ter ne krši avtorskih ali drugih pravic tretjih oseb;',
        'se zavezuje, da bo režijo uprizoritve v celoti opravil sam kot izključni avtor režije, brez sorežiserjev;',
        'razpolaga z vsemi potrebnimi pooblastili za prenos materialnih avtorskih pravic po tej pogodbi.',
    ]:
        add_bullet(doc, b)
    add_para_runs(doc, [('Jamstva iz prejšnjega odstavka se ne raztezajo na sorodne pravice drugih ustvarjalcev in izvajalcev uprizoritve (igralci, skladatelj glasbe, scenograf, kostumograf, lektor), katerih pravice naročnik uredi z ločenimi pogodbami.', False)])
    citat_line(doc, 'ZASP_118')
    arg_line(doc, 'Opomba o sorodnih pravicah: 118. čl. ZASP definira "izvajalce" (igralci, pevci, glasbeniki, plesalci, režiserji drugih gledaliških predstav — tu Elias, ampak tudi drugi soustvarjalci) kot imetnike sorodnih pravic. Eliasovo jamstvo pokriva SAMO njegova dela (besedilo + režija), ne pravic drugih. Naročnik jih ureja z ločenimi pogodbami.')
    add_para_runs(doc, [('V primeru kakršnih koli zahtevkov tretjih oseb iz naslova avtorskih ali sorodnih pravic na delih iz 3. člena te pogodbe prevzema izvajalec polno odgovornost.', False)])

    # ============== VIII. OBVEZNOSTI IZVAJALCA ==============
    add_section(doc, 'OBVEZNOSTI IZVAJALCA')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec opravi delo vestno in v rokih, ki jih opredeljuje terminski plan, sprejet sporazumno z direktorjem gledališča.', False)])
    add_para_runs(doc, [('Terminski plan in razpored vaj koordinira organizator kulturnega programa naročnika, sporazumno z izvajalcem, ob upoštevanju splošnega tedenskega programa gledališča in razpoložljivosti soustvarjalcev.', False)])
    arg_line(doc, 'Mehčana dikcija po DP-priporočilu (ne po Milanovi izrecni zahtevi). Ira ima trše: "razpored VAJ JE DOLOČEN s tedenskim programom gledališča, ki ga celotni ekipi projekta POSREDUJE organizator kulturnega programa naročnika."')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec se v zvezi s prevzetim delom obvezuje ravnati glede rabe sredstev (materialnih, finančnih in kadrovskih) kot skrben gospodar in pri tem upoštevati finančne, tehnične in kadrovske zmožnosti naročnika ter splošna pravila, ki veljajo v prostorih naročnika.', False)])
    arg_line(doc, 'Mehčano: brez izrecnega "hišnega reda" (Ira ima izrecno). "Skrben gospodar" je standardni izraz OZ (6. čl. OZ — načelo dobre vere in poštenja).')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec sodeluje pri tiskovnih konferencah, predstavitvah in drugih promocijskih dejavnostih, povezanih z uprizoritvijo, v obsegu, ki je običajen za tovrstne produkcije, ter prispeva podatke za gledališki list.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Če izvajalec zaradi bolezni ali druge upravičene odsotnosti začasno ne more opravljati dela, o tem nemudoma obvesti naročnika. Pogodbeni stranki sporazumno določita nadomestno rešitev ali prilagodita terminski plan.', False)])
    add_para_runs(doc, [('Če izvajalec prevzetega dela iz objektivnih razlogov (višja sila, huda bolezen ipd.) ne more nadaljevati, naročnik nedokončano delo dokonča z nadomestnim ustvarjalcem, ki ga stranki določita sporazumno, ob varovanju moralnih avtorskih pravic izvajalca. V tem primeru ima izvajalec pravico do honorarja v sorazmerju z že opravljenim delom; naročnik zadrži pravico uporabe že ustvarjenih idejnih zasnov in dramskega besedila za to produkcijo, za katere je bil sorazmerni del honorarja izplačan.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Morebitno odsotnost, ki jo lahko povzroči bolezen ali višja sila in ki lahko povzroči kakršenkoli odlog načrtovanega terminskega plana, izvajalec sporoči naročniku (organizatorju kulturnega programa naročnika) takoj, ko je to mogoče (telefonsko, s SMS ali po elektronski pošti).', False)])

    # ============== IX. KONČNE DOLOČBE ==============
    add_section(doc, 'KONČNE DOLOČBE')

    add_clen(doc)
    add_para_runs(doc, [('Naročnik lahko odstopi od te pogodbe, če izvajalec svojih obveznosti ne izpolnjuje in s tem naročniku povzroča materialno škodo ali škoduje njegovemu ugledu.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec lahko odstopi od te pogodbe, če naročnik brez njegove krivde ne izpelje uprizoritve do premiere. V tem primeru sme izvajalec zadržati že izplačani honorar kot nadomestilo za opravljeno delo do trenutka razdrtja, pravice na dramskem besedilu KDO JE ELENA? pa v celoti ostanejo izvajalcu.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Izrazi za osebe, ki so v tej pogodbi zapisani v moški slovnični obliki, se uporabljajo nevtralno in veljajo za vse spole.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Pogodbeni stranki soglašata, da dogovorjeni znesek avtorskega honorarja ni javen podatek in ga v javnosti ne bosta uporabljali, razen kadar to določa zakon.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Naročnik se obvezuje, da bo osebne podatke izvajalca varoval in obdeloval izključno za namene izvajanja te pogodbe, skladno s Splošno uredbo o varstvu podatkov (GDPR) in veljavnim Zakonom o varstvu osebnih podatkov (ZVOP-2).', False)])
    arg_line(doc, 'GDPR = Uredba (EU) 2016/679 (splošno). ZVOP-2 = slovenski Zakon o varstvu osebnih podatkov, sprejet 15. 12. 2022. Verbatim ni ekstrahirano; klavzula je standardna in ne potrebuje specifičnega člena.')

    add_clen(doc)
    add_para_runs(doc, [('Vse morebitne spremembe in dopolnitve te pogodbe veljajo le v primeru sklenitve pisnega aneksa, ki ga podpišeta obe pogodbeni stranki.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Morebitne spore iz te pogodbe bosta pogodbeni stranki reševali sporazumno. Če to ni mogoče, je za reševanje sporov pristojno stvarno pristojno sodišče v Ljubljani.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Ta pogodba je sestavljena v treh (3) enakih izvodih, od katerih prejme izvajalec en (1) izvod, naročnik pa dva (2). Pogodba začne veljati z dnem podpisa obeh pogodbenih strank.', False)])

    # === PODPISNI BLOK ===
    build_signature_block_v7(
        doc,
        datum='1. september 2026',
        izvajalec_naziv=None,
        izvajalec_ime='Elias Rudolf',
        izvajalec_funkcija='dramatik in režiser',
        narocnik_naziv='Šentjakobsko gledališče Ljubljana - društvo',
        narocnik_ime='Milan Golob',
        narocnik_funkcija='direktor',
    )

    # === POVZETEK LEČE ===
    doc.add_paragraph()
    add_section(doc, 'POVZETEK LEČE — verbatim citati in ugotovitve')

    povzetek = [
        ('✓ POTRJENO (verbatim najden, argumentacija stoji):',
         '5. čl. ZASP · 23. čl. ZASP · 24. čl. ZASP · 26. čl. ZASP · 30. čl. ZASP · 32.a čl. ZASP · 118. čl. ZASP · 21. čl. ZOdv'),
        ('⚠ POPRAVLJENO (napačen sklic v prejšnjih verzijah):',
         '16. čl. ZOdv (v v3/v4) → pravilno 21. čl. ZOdv. Napaka izvirala iz moje predpostavke; verbatim iz PIS potrdil, da 16. čl. govori o statusu specialista, ne o umetniški dejavnosti.'),
        ('⚠ DVOMLJIVO (potrebna dodatna preverba, verjetno napačen sklic):',
         '76.a čl. ZDDV-1 (uvod strank in 5. čl.) — se nanaša na reverse charge, ne na oprostitev/malega davčnega zavezanca. Pravilnejši sklic je verjetno 94. čl. (mali) ali 44. čl. (oprostitev). Elias naj kot odvetnik sam potrdi ali predlaga popravek.'),
        ('◯ NI EKSTRAHIRANO (v tej lečni verziji še ni verbatim citata — priporočilo v6):',
         '75.–78. čl. ZASP (ozka razlaga prenosa, trajanje) · 83. čl. ZASP (vračanje pravic) · 16.–20. čl. ZASP (moralne pravice) · 39. čl. ZASP (razmerje lastnine nosilca vs. pravice) · GDPR + ZVOP-2 (splošni sklic zadošča) · OZ (splošna pravila obligacij)'),
        ('OSTAJA ODPRTO (subjektivna odločitev):',
         '(1) Ekskluzivnost drugim gledališčem (8. čl. — trenutno neekskluzivno). (2) Mehčanje delovno-pravnih indicij (12.–16. čl. — trenutno mehčano po DP-priporočilu, ne po Milanovi izrecni zahtevi). (3) Ali obdržati klavzulo o registrirani pisarni (1. čl., 3. odst.) — pravna podlaga je 21. čl. ZOdv (potrjeno), a pravno ni obveznost do ŠGL.'),
    ]
    for naslov, telo in povzetek:
        p = doc.add_paragraph()
        apply_no_spacing(p)
        r = p.add_run(naslov)
        r.bold = True
        r.font.size = Pt(10)
        r.font.name = 'Calibri'
        r.font.color.rgb = RGBColor.from_string(COL_CITAT)
        p2 = doc.add_paragraph()
        apply_no_spacing(p2)
        r2 = p2.add_run(telo)
        r2.font.size = Pt(9.5)
        r2.font.name = 'Calibri'
        doc.add_paragraph()

    doc.save(str(OUT_DOCX))
    if tmp is not None and tmp.exists():
        tmp.unlink()
    print(f'GENERIRANO: {OUT_DOCX}')
    print(f'Velikost: {OUT_DOCX.stat().st_size:,} B')
    print(f'Verbatim citiranih členov: {len(CITATI)}')


if __name__ == '__main__':
    main()
