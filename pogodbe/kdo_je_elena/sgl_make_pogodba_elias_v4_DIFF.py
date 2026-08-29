# -*- coding: utf-8 -*-
"""
sgl_make_pogodba_elias_v4_DIFF.py
================================
DIFF verzija Elias v4 — enak vsebinski del kot v4, s VIDNIMI oznakami
sprememb v3→v4 in GLASNIMI opozorili o standardu v7.

Namen: Milanu omogoči, da v enem dokumentu vidi:
  - kaj se je spremenilo iz v3 v v4 (vijolične inline oznake)
  - kje se je v3 tepel s standardom v7 (oranžni bloki, glasno)
  - kje so odprta vprašanja glede lokalne sgl_docx_format.py verzije

Barvna semantika:
  - VIJOLIČNA (8B008B) — dif marker: kaj se je spremenilo iz v3 v v4
  - ORANŽNA (E67E22)   — GLASNO OPOZORILO: konflikt s standardom
  - RDEČA (C00000)     — kritični bloki na vrhu dokumenta

Output:
  Šentjakobsko gledališče\\Pogodbe in računi\\Kdo je Elena\\
  ŠGL_avtorska pogodba_KDO JE ELENA_Elias Rudolf_v4_DIFF-vs-v3.docx
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
OUT_DOCX = OUT_DIR / 'ŠGL_avtorska pogodba_KDO JE ELENA_Elias Rudolf_v4_DIFF-vs-v3.docx'

COL_DIFF = '8B008B'   # vijolična — dif marker
COL_WARN = 'E67E22'   # oranžna — opozorilo
COL_CRIT = 'C00000'   # rdeča — kritični blok


def init_doc():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if not REF_DOCX.exists():
        return Document(), None
    tmp = OUT_DIR / '_tmp_elias_v4_diff.docx'
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


def diff_line(doc, besedilo):
    """Vijoličen inline marker sprememb v3→v4 pod členom."""
    p = doc.add_paragraph()
    apply_no_spacing(p)
    r1 = p.add_run('▸ [v3→v4] ')
    r1.bold = True
    r1.italic = True
    r1.font.size = Pt(8.5)
    r1.font.name = 'Calibri'
    r1.font.color.rgb = RGBColor.from_string(COL_DIFF)
    r2 = p.add_run(besedilo)
    r2.italic = True
    r2.font.size = Pt(8.5)
    r2.font.name = 'Calibri'
    r2.font.color.rgb = RGBColor.from_string(COL_DIFF)


def warn_block(doc, naslov, telo, color=COL_WARN):
    """Velik obarvan blok — GLASNO OPOZORILO."""
    doc.add_paragraph()
    # naslov v celoti
    p1 = doc.add_paragraph()
    apply_no_spacing(p1)
    r = p1.add_run(f'⚠⚠⚠  {naslov}  ⚠⚠⚠')
    r.bold = True
    r.font.size = Pt(11)
    r.font.name = 'Calibri'
    r.font.color.rgb = RGBColor.from_string(color)
    # telo
    p2 = doc.add_paragraph()
    apply_no_spacing(p2)
    r2 = p2.add_run(telo)
    r2.font.size = Pt(10)
    r2.font.name = 'Calibri'
    r2.font.color.rgb = RGBColor.from_string(color)
    doc.add_paragraph()


def head_block(doc, naslov, telo):
    """Zgornji blok — glava opozoril."""
    p1 = doc.add_paragraph()
    apply_no_spacing(p1)
    r = p1.add_run(naslov)
    r.bold = True
    r.font.size = Pt(12)
    r.font.name = 'Calibri'
    r.font.color.rgb = RGBColor.from_string(COL_CRIT)
    p2 = doc.add_paragraph()
    apply_no_spacing(p2)
    r2 = p2.add_run(telo)
    r2.font.size = Pt(10)
    r2.font.name = 'Calibri'
    r2.font.color.rgb = RGBColor.from_string(COL_CRIT)


def build_signature_block_v7(doc, datum, izvajalec_naziv, izvajalec_ime, izvajalec_funkcija,
                              narocnik_naziv, narocnik_ime, narocnik_funkcija):
    """Podpisni blok po v7 §4.0 (grajen eksplicitno)."""
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
    cell(4, 1, narocnik_naziv, bold=True)  # B.1 #1: naziv ŠGL bold
    cell(5, 0, izvajalec_ime, bold=True)   # v7 §4.0: vrstica 5 IME BOLD
    cell(5, 1, narocnik_ime, bold=True)    # v7 §4.0: vrstica 5 IME BOLD
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
    # ZGORNJI BLOK: KRITIČNA OPOZORILA
    # ============================================================
    head_block(
        doc,
        'DIFF VERZIJA — Elias Rudolf v4 vs v3 · NI ZA PODPIS',
        'V tem dokumentu so označene VSE spremembe v3→v4 in KONFLIKTI, ki jih je '
        'v3 imel s standardom v7 (ŠG_STANDARDI_pogodbe_v7_changelog.md, potrjeno '
        '20. 8. 2026). Namen: primerjava. Za podpis uporabi ločeno čisto verzijo v4.'
    )
    doc.add_paragraph()

    # opozorilo #1 — podpisni blok NAPAKA v v3
    warn_block(
        doc,
        'KONFLIKT S STANDARDOM v7 §4.0 — PODPISNI BLOK V v3 STRUKTURNO NAPAČEN',
        'v7 §4.0 (Milanov popravek 20. 8. 2026) izrecno predpisuje razporeditev '
        'tabele 9×2:  vrstica 0 datum · 2 oznaka · 4 naziv pravne osebe · '
        'VRSTICA 5 IME+PRIIMEK BOLD · 6 funkcija · 8 žig. '
        'V mojem v3 podpisnem bloku so bili vsi ti podatki pomaknjeni za 1 vrstico '
        'dol (imena v vrstici 6/7, funkcije v vrstici 7/8, IN brez bolda). '
        'To je natanko napaka, pred katero standard v7 §4.0 IZRECNO SVARI. '
        'Popravek v v4: tabela zdaj sledi točno §4.0; imena so bold; nazivi '
        'pravnih oseb bold po registru B.1 #1; žig na obeh straneh '
        '(odvetnik ima žig; ŠGL ima žig).'
    )

    # opozorilo #2 — sgl_docx_format.py verzija
    warn_block(
        doc,
        'ODVISNOST OD LOKALNE sgl_docx_format.py (potrebna verzija ≥ 2026-08-20)',
        'Standard v7 §5.6.2 (20. 8.) je spremenil zamike poglavij: 0,63 / −1,27 cm → '
        '0,75 / −0,75 cm (rimska številka poravnana z levim robom telesa). '
        'Standard v7 §5.6.10 (20. 8.) je dodal: alineje z označevalcem "–" (en-dash) '
        '+ presledek. '
        'Če tvoja lokalna sgl_docx_format.py NI POSODOBLJENA na verzijo od '
        '2026-08-20, poglavja bodo imela stare zamike in alineje bodo BREZ '
        'pomišljaja. Kopija sgl_docx_format.py na Google Drive je iz 7. 5. 2026 '
        '(še stara). Preveri lokalno pot na Windowsu.'
    )

    # opozorilo #3 — pravni scenarij odvetnik = avtor = izvajalec
    warn_block(
        doc,
        'INFORMATIVNO — VRSTNI RED STRANK v7 §3.2 NE POKRIVA ELIASA IZRECNO',
        'v7 §3.2 predpisuje tristranski vrstni red (Naročnik / Avtor F.O. / Izvajalec '
        'pravna oseba) SAMO za "sodelavski Potodom". Pri Eliasu ni tretje pravne '
        'osebe — samostojni odvetnik po ZOdv posluje kot F.O. z registrirano '
        'poklicno dejavnostjo, ne kot pravna oseba (d.o.o./s.p./zavod). '
        'Zato je pri Eliasu DVOSTRANSKA pogodba (dvojina glagolov: "stranki soglašata"). '
        'To ni konflikt s standardom, je le razjasnitev; v4 je DVOSTRANSKA. '
        'Odstavek v 1. členu, ki pojasni dvojno identiteto (avtor = izvajalec = ista '
        'F.O., ki dela preko svoje pisarne po 16. čl. ZOdv), NI predpisan v v7 — '
        'je Elias-specifika (novost, ki bo šla v v8).'
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
    diff_line(doc, 'DODAN naziv "g." pred imenom — v7 §3.1 (g./ga. samo v uvodu strank; §3.3 pravi "pred imenom, ne pred funkcijo"). v3 je imel samo "Elias Rudolf, odvetnik".')

    add_para_runs(doc, [('matična številka: 2861887000', False)])
    add_para_runs(doc, [('davčna številka: 34928138', False)])
    add_para_runs(doc, [('davčni zavezanec za DDV: NE (76.a člen ZDDV-1)', False)])
    add_para_runs(doc, [('TRR: SI56 0400 0028 0838 318 (OTP banka d.d.)', False)])
    add_para_runs(doc, [('kot avtor in režiser (v nadaljevanju izvajalec),', False)])
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
        ('marca 2027', True),
        (' na Šentjakobskem odru Šentjakobskega gledališča Ljubljana - društvo (v nadaljevanju: uprizoritev).', False),
    ])
    diff_line(doc, 'DODAN BOLD na "marca 2027" — datum premiere po bold registru B.1 #6. v3 je imel "v marcu 2027" plain.')

    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da je dramsko besedilo ', False),
        ('KDO JE ELENA?', True),
        (' izvirno avtorsko delo izvajalca v smislu 5. člena Zakona o avtorski in sorodnih pravicah (v nadaljevanju: ZASP), ki ga je izvajalec ustvaril samostojno in ne predstavlja priredbe ali predelave dela tretje osebe.', False),
    ])
    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da izvajalec pri tej pogodbi nastopa kot avtor in režiser uprizoritve. Avtorstvo dramskega besedila in režijske izvedbe pripada izvajalcu kot fizični osebi v skladu z ZASP. Delo po tej pogodbi izvajalec opravlja preko svoje registrirane pisarne, kar je skladno s 16. členom Zakona o odvetništvu (dovoljena umetniška dejavnost).', False),
    ])
    diff_line(doc, 'NESPREMENJENO iz v3 — Elias-specifika, ni predpisano v v7 (predlagam za v8).')
    add_para_runs(doc, [('Pogodbeni stranki se strinjata, da so storitve, ki so predmet te pogodbe, avtorsko delo, skladno z ZASP.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Pogodbeni stranki soglašata, da se lahko datum premiere oziroma datumi ponovitev, vaj ali drugih dogodkov v zvezi z uprizoritvijo spremenijo iz razlogov višje sile ali drugih nepredvidenih okoliščin, brez odgovornosti katere koli pogodbene stranke.', False)])

    # ============== II. PREDMET POGODBE ==============
    add_section(doc, 'PREDMET POGODBE')

    add_clen(doc)
    add_para_runs(doc, [('S to pogodbo izvajalec prevzema pri uprizoritvi dve avtorski vlogi:', False)])
    for b in [
        'avtorstvo izvirnega dramskega dela KDO JE ELENA? (avtorsko delo po 5. členu ZASP);',
        'režijo uprizoritve (avtorsko delo po 5. členu ZASP).',
    ]:
        add_bullet(doc, b)

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec je odgovoren za umetniško vrednost uprizoritve in za usklajeno delovanje z direktorjem gledališča glede izvedbe sprejetega terminskega plana. Izvajalec prevzame izvedbo idejne zasnove uprizoritve in koordinacijo dela soustvarjalcev.', False)])
    add_para_runs(doc, [('Pri izbiri soustvarjalcev in pri finančnih zahtevah, povezanih z uprizoritvijo, izvajalec sodeluje z direktorjem gledališča ter upošteva finančne, tehnične in kadrovske zmožnosti naročnika.', False)])

    # ============== III. NARAVA RAZMERJA ==============
    # OPOZORILO PRED III. — GLASNO
    warn_block(
        doc,
        'POPRAVEK NASLOVA POGLAVJA III. PO v7 §1.1 IN §8 CHANGELOG #2',
        'v3 je imel naslov "AVTORSKI HONORAR" (staro ime iz v6). v7 predpisuje '
        '"NARAVA RAZMERJA" kot nadrejen termin, ki pokriva vse tipe pogodb '
        '(honorarna, pro bono, B2B, koprodukcijska). Popravek v v4 spodaj.'
    )
    add_section(doc, 'NARAVA RAZMERJA')
    diff_line(doc, 'PREIMENOVANO: "AVTORSKI HONORAR" (v3) → "NARAVA RAZMERJA" (v4, po v7 §1.1)')

    add_clen(doc)
    add_para_runs(doc, [
        ('Pogodbeni stranki soglašata, da prejme izvajalec za delo po tej pogodbi, ki obsega obe avtorski vlogi iz 3. člena te pogodbe (avtorstvo dramskega dela in režijo), avtorski honorar v skupni višini ', False),
        ('2.300,00 EUR', True),
        ('. Ker izvajalec ni davčni zavezanec za DDV (76.a člen ZDDV-1), se davek na dodano vrednost ne obračuna.', False),
    ])
    diff_line(doc, 'NESPREMENJENO iz v3 glede zneska; besedilo skladno z v7 §5.2 (brez "bruto" pri B2B ✓).')
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
    add_para_runs(doc, [('Izvajalec izrecno dovoljuje neomejeno število ponovitev brez dodatnih zahtev za plačilo honorarja:', False)])
    for b in ['na matičnem odru naročnika,', 'na gostovanjih doma in v tujini,', 'na festivalih doma in v tujini,', 'za televizijske in spletne prenose.']:
        add_bullet(doc, b)

    add_clen(doc)
    add_para_runs(doc, [('Pravice za žive izvedbe in ponovitve iz prejšnjega člena veljajo, dokler je uprizoritev na rednem repertoarju Šentjakobskega gledališča Ljubljana - društvo, oziroma dokler je upravni odbor gledališča formalno ne umakne z repertoarja. Pravice za dokumentiranje in arhiviranje ter uporabo že nastalih fotografij in posnetkov v informativno-promocijskih materialih se prenesejo trajno in neizključno.', False)])
    add_para_runs(doc, [('Z umikom uprizoritve z repertoarja prenesene materialne avtorske pravice za žive izvedbe in ponovitve v celoti preidejo nazaj na izvajalca, brez potrebe po dodatnem pravnem dejanju.', False)])

    add_clen(doc)
    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da izvajalec ', False),
        ('ostane avtor in imetnik avtorskih pravic na dramskem besedilu', True),
        (' KDO JE ELENA? kot samostojnem literarnem delu. Prenos pravic iz te pogodbe se nanaša izključno na uprizarjanje tega besedila v produkciji naročnika in ne omejuje pravice izvajalca, da svoje dramsko besedilo objavi, ponudi v uprizoritev drugim gledališčem ali kako drugače uporablja zunaj te produkcije.', False),
    ])

    # ============== V. MORALNE AVTORSKE PRAVICE ==============
    add_section(doc, 'MORALNE AVTORSKE PRAVICE')

    add_clen(doc)
    add_para_runs(doc, [('Moralne avtorske pravice ostanejo izvajalcu v skladu z ZASP, zlasti pravica do priznanja avtorstva in pravica do celovitosti dela.', False)])
    add_para_runs(doc, [('Naročnik bo izvajalca v vseh informativno-promocijskih materialih, gledališkem listu in drugih javnih objavah, povezanih z uprizoritvijo, navedel s polnim imenom in priimkom — Elias Rudolf — z navedbo obeh avtorskih vlog: "avtor besedila in režiser".', False)])
    add_para_runs(doc, [('Izvajalec ima pravico, da se upre vsaki skazitvi, okrnitvi ali drugačni spremembi svojega dela ter pravico, da se upre vsaki izvedbi dela, ki bi žalila njegovo čast in ugled.', False)])
    add_para_runs(doc, [('Izvajalec je seznanjen, da bo njegovo delo predmet tržnega komuniciranja, ki ga za namene obveščanja javnosti izvaja naročnik.', False)])

    # ============== VI. FIZIČNI NOSILCI IN LASTNINA ==============
    add_section(doc, 'FIZIČNI NOSILCI IN LASTNINA')

    add_clen(doc)
    add_para_runs(doc, [('Fizični in digitalni nosilci, ki jih izvajalec izroči naročniku za namen izvedbe uprizoritve (rokopis dramskega besedila, beleške, režijska knjiga, delovni materiali), ostanejo v hrambi naročnika za čas, dokler je uprizoritev na repertoarju, izključno za namen uprizarjanja, arhiviranja in promocije uprizoritve.', False)])
    add_para_runs(doc, [('Izvirnik dramskega besedila in drugi izvirniki avtorskega dela ostanejo v lasti izvajalca.', False)])

    # ============== VII. JAMSTVA ==============
    add_section(doc, 'JAMSTVA')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec jamči naročniku, da:', False)])
    for b in [
        'je avtor izvirnega dramskega dela KDO JE ELENA? in da to delo ni priredba ali predelava dela tretje osebe ter ne krši avtorskih ali drugih pravic tretjih oseb;',
        'je edini avtor obeh del iz 3. člena te pogodbe;',
        'razpolaga z vsemi potrebnimi pooblastili za prenos materialnih avtorskih pravic po tej pogodbi.',
    ]:
        add_bullet(doc, b)
    add_para_runs(doc, [('Jamstva iz prejšnjega odstavka se ne raztezajo na sorodne pravice drugih ustvarjalcev in izvajalcev uprizoritve (igralci, skladatelj glasbe, scenograf, kostumograf, lektor), katerih pravice naročnik uredi z ločenimi pogodbami.', False)])
    add_para_runs(doc, [('V primeru kakršnih koli zahtevkov tretjih oseb iz naslova avtorskih ali sorodnih pravic na delih iz 3. člena te pogodbe prevzema izvajalec polno odgovornost.', False)])

    # ============== VIII. OBVEZNOSTI IZVAJALCA ==============
    add_section(doc, 'OBVEZNOSTI IZVAJALCA')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec opravi delo vestno in v rokih, ki jih opredeljuje terminski plan, sprejet sporazumno z direktorjem gledališča.', False)])
    add_para_runs(doc, [('Razpored vaj določita pogodbeni stranki sporazumno, ob upoštevanju splošnega tedenskega programa gledališča in razpoložljivosti soustvarjalcev.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec se v zvezi s prevzetim delom obvezuje ravnati glede rabe sredstev (materialnih, finančnih in kadrovskih) kot skrben gospodar in pri tem upoštevati finančne, tehnične in kadrovske zmožnosti naročnika ter splošna pravila, ki veljajo v prostorih naročnika.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec sodeluje pri tiskovnih konferencah, predstavitvah in drugih promocijskih dejavnostih, povezanih z uprizoritvijo, v obsegu, ki je običajen za tovrstne produkcije, ter prispeva podatke za gledališki list.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Če izvajalec zaradi bolezni ali druge upravičene odsotnosti začasno ne more opravljati dela, o tem nemudoma obvesti naročnika. Pogodbeni stranki sporazumno določita nadomestno rešitev ali prilagodita terminski plan.', False)])
    add_para_runs(doc, [('Če izvajalec prevzetega dela iz objektivnih razlogov (višja sila, huda bolezen ipd.) ne more nadaljevati, naročnik nedokončano delo dokonča z nadomestnim ustvarjalcem, ki ga stranki določita sporazumno, ob varovanju moralnih avtorskih pravic izvajalca. V tem primeru ima izvajalec pravico do honorarja v sorazmerju z že opravljenim delom; naročnik zadrži pravico uporabe že ustvarjenih idejnih zasnov in dramskega besedila za to produkcijo, za katere je bil sorazmerni del honorarja izplačan.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Morebitno odsotnost, ki jo lahko povzroči bolezen ali višja sila in ki lahko povzroči kakršenkoli odlog načrtovanega terminskega plana, izvajalec sporoči naročniku takoj, ko je to mogoče (telefonsko, s SMS ali po elektronski pošti).', False)])

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

    add_clen(doc)
    add_para_runs(doc, [('Vse morebitne spremembe in dopolnitve te pogodbe veljajo le v primeru sklenitve pisnega aneksa, ki ga podpišeta obe pogodbeni stranki.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Morebitne spore iz te pogodbe bosta pogodbeni stranki reševali sporazumno. Če to ni mogoče, je za reševanje sporov pristojno stvarno pristojno sodišče v Ljubljani.', False)])

    add_clen(doc)
    add_para_runs(doc, [('Ta pogodba je sestavljena v treh (3) enakih izvodih, od katerih prejme izvajalec en (1) izvod, naročnik pa dva (2). Pogodba začne veljati z dnem podpisa obeh pogodbenih strank.', False)])

    # === PODPISNI BLOK — z opozorilom ===
    warn_block(
        doc,
        'PODPISNI BLOK v4 — sledi v7 §4.0 razporeditvi (napaka v3 popravljena)',
        'Vrstica 4 = naziv pravne osebe (levo prazna pri samostojnem odvetniku; '
        'desno "Šentjakobsko gledališče..." BOLD po registru B.1 #1). '
        'Vrstica 5 = IME+PRIIMEK BOLD (levo Elias Rudolf, desno Milan Golob). '
        'Vrstica 6 = funkcija. Vrstica 8 = ŽIG obojestransko (odvetnik ima žig).'
    )
    build_signature_block_v7(
        doc,
        datum='1. september 2026',
        izvajalec_naziv=None,
        izvajalec_ime='Elias Rudolf',
        izvajalec_funkcija='odvetnik, avtor besedila in režiser',
        narocnik_naziv='Šentjakobsko gledališče Ljubljana - društvo',
        narocnik_ime='Milan Golob',
        narocnik_funkcija='direktor',
    )

    # === POVZETEK POPRAVKOV NA KONCU ===
    doc.add_paragraph()
    add_section(doc, 'POVZETEK POPRAVKOV v3 → v4')

    povzetek = [
        ('P1 · uvod strank', 'dodan "g." pred imenom Elias (v7 §3.1)'),
        ('P2 · III. poglavje', '"AVTORSKI HONORAR" → "NARAVA RAZMERJA" (v7 §1.1, §8)'),
        ('P3 · 1. člen', '"marca 2027" BOLD (bold register B.1 #6)'),
        ('P4 · podpisni blok', 'popolna prenova po v7 §4.0 — vrstice, bold, žig obojestransko'),
        ('P5 · dvostranska struktura', 'jasno DVOSTRANSKA (avtor = izvajalec = ista F.O.)'),
    ]
    for kaj, opis in povzetek:
        p = doc.add_paragraph()
        apply_no_spacing(p)
        r1 = p.add_run(f'• {kaj}: ')
        r1.bold = True
        r1.font.size = Pt(10)
        r1.font.name = 'Calibri'
        r1.font.color.rgb = RGBColor.from_string(COL_DIFF)
        r2 = p.add_run(opis)
        r2.font.size = Pt(10)
        r2.font.name = 'Calibri'
        r2.font.color.rgb = RGBColor.from_string(COL_DIFF)

    doc.add_paragraph()
    warn_block(
        doc,
        'PRIPOROČILO ZA NASLEDNJI KORAK',
        'Preveri lokalno sgl_docx_format.py — če je pred 2026-08-20, posodobi:  '
        '(1) add_section() zamiki 0,75 / −0,75;  '
        '(2) add_bullet() dodaj pomišljaj "–" + presledek;  '
        '(3) dodaj add_signature_block() helper (ta skripta ga NE uporablja — '
        'gradi blok eksplicitno, tako da deluje neodvisno od verzije helperja). '
        'Za konsolidirano referenco (v8 standard) povej — pripravim sinteza dokument.'
    )

    doc.save(str(OUT_DOCX))
    if tmp is not None and tmp.exists():
        tmp.unlink()
    print(f'GENERIRANO: {OUT_DOCX}')
    print(f'Velikost: {OUT_DOCX.stat().st_size:,} B')


if __name__ == '__main__':
    main()
