"""
sgl_make_pogodba_elias_rudolf_v4.py
==================================
v4: aplikacija manjkajočih pravil iz standarda v7 na v3.

Popravki v4 vs v3 (potrjeno iz ŠG_STANDARDI_pogodbe_v7_changelog.md, 20. 8. 2026):

  P1. Uvod strank: dodan naziv "g." pred imenom (v7 §3.1 — g./ga. samo v uvodu)
  P2. III. poglavje ime: "AVTORSKI HONORAR" → "NARAVA RAZMERJA"
      (v7 §1.1 tabela; §8 changelog #2)
  P3. Datum premiere v 1. členu bold (v7 oblikovalska taksonomija B.1 #6)
  P4. Podpisni blok popolna prenova po v7 §4.0 (dodan 2026-08-20 Milan):
      - vrstice: 0 datum · 2 oznaka · 4 naziv pravne osebe · 5 IME+PRIIMEK BOLD
        · 6 funkcija · 8 žig (obojestransko)
      - NAPAKA v3: imena so bila v vrstici 6/7 (ne 5) in NIVE bila bold
      - naziv "Šentjakobsko gledališče" tudi BOLD (v7 taksonomija B.1 #1)
  P5. Dvostranska pogodba pravilno: pri Eliasu se AVTOR (F.O.) in IZVAJALEC
      (samostojni odvetnik) zlivata v eno pravno osebo — DVE stranki, dvojina
      glagolov ("stranki soglašata"), ne tristranska.

============================================================
OPOZORILA GLEDE ODVISNOSTI OD sgl_docx_format.py
============================================================

Ta skripta predpostavlja, da Milanova lokalna sgl_docx_format.py vsebuje
POPRAVKE OD 2026-08-20:

  U1. add_section(): zamiki poglavij 0,75 / −0,75 cm (§5.6.2 popravek 20. 8.)
      → prej 0,63 / −1,27; rimska številka je zdaj poravnana z levim robom
        telesa besedila.
  U2. add_bullet(): alineja doda POMIŠLJAJ "–" + presledek (§5.6.10, 20. 8.)
      → prej brez marker-ja.
  U3. add_signature_block(): enotni helper za podpisni blok (§4.0, 20. 8.)
      → prej ročno grajen (napaka pomika vrstic).

Ta skripta:
  - U3: ne kliče add_signature_block (kompatibilnost s starimi verzijami);
        podpisni blok gradim EKSPLICITNO po v7 §4.0 postavitvi, tako da je
        rezultat pravilen ne glede na verzijo helperja.
  - U2: kliče add_bullet — ČE imaš staro verzijo BREZ pomišljaja, alineje
        bodo brez marker-ja. Priporočilo: posodobi lokalno sgl_docx_format.py
        na verzijo od 2026-08-20.
  - U1: kliče add_section — ČE imaš staro verzijo, zamiki bodo drugačni.

============================================================
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

OUT_DIR = Path(r'C:\Users\ladmin\Moj disk\Šentjakobsko gledališče\Pogodbe in računi\Kdo je Elena')
REF_DOCX = Path(r'C:\Users\ladmin\Moj disk\Šentjakobsko gledališče\Pogodbe in računi\Šunder v dvorani\ŠGL_avtorska pogodba_ŠUNDER_Maša Milčinski_lektura_v3_CLEAN.docx')
OUT_DOCX = OUT_DIR / 'ŠGL_avtorska pogodba_KDO JE ELENA_Elias Rudolf_avtor-režiser_v4.docx'


def init_doc():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if not REF_DOCX.exists():
        return Document(), None
    tmp = OUT_DIR / '_tmp_elias_v4.docx'
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


def build_signature_block_v7(doc, datum, izvajalec_naziv, izvajalec_ime, izvajalec_funkcija,
                              narocnik_naziv, narocnik_ime, narocnik_funkcija):
    """Podpisni blok po v7 §4.0 postavitvi (grajen eksplicitno, brez odvisnosti
    od add_signature_block v sgl_docx_format).

    Vrstice tabele:
      0: Datum:                        Datum:
      1: (prazno)
      2: Izvajalec:                    Naročnik:
      3: (prazno)
      4: <naziv prav. os. ali prazno>  <naziv prav. os. — Šentjakobsko gled.>
      5: <IME+PRIIMEK BOLD>            <IME+PRIIMEK BOLD>
      6: <funkcija>                    <funkcija>
      7: (prazno)
      8: Žig:                          Žig:

    Nazivi pravnih oseb (vrstica 4) so BOLD po Milanovem bold registru B.1 #1.
    Ime + priimek (vrstica 5) je vedno BOLD po v7 §4.0 + bold register B.1 #3.
    Pri fizični osebi / samostojnem odvetniku / SVK leva celica 4 ostane
    prazna.
    """
    doc.add_paragraph()
    doc.add_paragraph()
    table = make_table_no_borders(doc, rows=9, cols=2)

    def cell(row, col, text, bold=False):
        fill_cell(table.cell(row, col), text, bold=bold)

    cell(0, 0, f'Datum: {datum}')
    cell(0, 1, f'Datum: {datum}')

    cell(2, 0, 'Izvajalec:')
    cell(2, 1, 'Naročnik:')

    # vrstica 4: naziv pravne osebe (BOLD po B.1 #1); pri F.O./samostojnem odvetniku
    # leva prazna po v7 §4.0
    cell(4, 0, izvajalec_naziv or '', bold=bool(izvajalec_naziv))
    cell(4, 1, narocnik_naziv, bold=True)

    # vrstica 5: IME+PRIIMEK — vedno BOLD
    cell(5, 0, izvajalec_ime, bold=True)
    cell(5, 1, narocnik_ime, bold=True)

    # vrstica 6: funkcija
    cell(6, 0, izvajalec_funkcija)
    cell(6, 1, narocnik_funkcija)

    # vrstica 8: žig obojestransko (odvetnik ima žig; ŠGL ima žig)
    cell(8, 0, 'Žig:')
    cell(8, 1, 'Žig:')

    return table


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc, tmp = init_doc()

    set_margins(doc, top_cm=2.5, bottom_cm=2.5, left_cm=2.5, right_cm=2.5)
    add_footer_pagenum(doc)

    # === UVOD STRANK ===
    add_para_runs(doc, [('Pogodbeni stranki:', False)])
    doc.add_paragraph()

    # P1 v7 §3.1: "g." v uvodu, pred imenom (ne pred funkcijo — §3.3)
    add_para_runs(doc, [
        ('Šentjakobsko gledališče Ljubljana - društvo', True),
        (', Krekov trg 2, 1000 Ljubljana, davčna številka: 31033008, matična številka: 5147689000, davčni zavezanec za DDV: NE, ki ga zastopa ', False),
        ('g. Milan Golob', True),
        (' kot direktor gledališča (v nadaljevanju naročnik),', False),
    ])
    doc.add_paragraph()
    add_para_runs(doc, [('in', False)])
    doc.add_paragraph()

    # P1 v7 §3.1: "g. Elias Rudolf" (g. pred imenom, funkcija za vejico)
    add_para_runs(doc, [
        ('g. Elias Rudolf, odvetnik', True),
        (' s pisarno v Ljubljani, Obrežna steza 2, 1000 Ljubljana,', False),
    ])
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

    # 1. člen — P3: "marca 2027" BOLD (bold register B.1 #6)
    add_clen(doc)
    add_para_runs(doc, [
        ('Ta pogodba se nanaša na avtorsko delo izvajalca pri krstni uprizoritvi gledališke predstave ', False),
        ('KDO JE ELENA?', True),
        (', katere avtor dramskega besedila in režiser je izvajalec, in ki bo premierno uprizorjena ', False),
        ('marca 2027', True),  # P3: BOLD
        (' na Šentjakobskem odru Šentjakobskega gledališča Ljubljana - društvo (v nadaljevanju: uprizoritev).', False),
    ])
    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da je dramsko besedilo ', False),
        ('KDO JE ELENA?', True),
        (' izvirno avtorsko delo izvajalca v smislu 5. člena Zakona o avtorski in sorodnih pravicah (v nadaljevanju: ZASP), ki ga je izvajalec ustvaril samostojno in ne predstavlja priredbe ali predelave dela tretje osebe.', False),
    ])
    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da izvajalec pri tej pogodbi nastopa kot avtor in režiser uprizoritve. Avtorstvo dramskega besedila in režijske izvedbe pripada izvajalcu kot fizični osebi v skladu z ZASP. Delo po tej pogodbi izvajalec opravlja preko svoje registrirane pisarne, kar je skladno s 16. členom Zakona o odvetništvu (dovoljena umetniška dejavnost).', False),
    ])
    add_para_runs(doc, [('Pogodbeni stranki se strinjata, da so storitve, ki so predmet te pogodbe, avtorsko delo, skladno z ZASP.', False)])

    # 2. člen
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

    # ============== III. NARAVA RAZMERJA (P2: preimenovano) ==============
    add_section(doc, 'NARAVA RAZMERJA')

    add_clen(doc)
    add_para_runs(doc, [
        ('Pogodbeni stranki soglašata, da prejme izvajalec za delo po tej pogodbi, ki obsega obe avtorski vlogi iz 3. člena te pogodbe (avtorstvo dramskega dela in režijo), avtorski honorar v skupni višini ', False),
        ('2.300,00 EUR', True),
        ('. Ker izvajalec ni davčni zavezanec za DDV (76.a člen ZDDV-1), se davek na dodano vrednost ne obračuna.', False),
    ])
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

    # === P4: PODPISNI BLOK po v7 §4.0 ===
    build_signature_block_v7(
        doc,
        datum='1. september 2026',
        izvajalec_naziv=None,  # F.O.-samostojni odvetnik, ni pravna oseba → leva prazna
        izvajalec_ime='Elias Rudolf',
        izvajalec_funkcija='odvetnik, avtor besedila in režiser',
        narocnik_naziv='Šentjakobsko gledališče Ljubljana - društvo',
        narocnik_ime='Milan Golob',
        narocnik_funkcija='direktor',
    )

    doc.save(str(OUT_DOCX))
    if tmp is not None and tmp.exists():
        tmp.unlink()
    print(f'GENERIRANO: {OUT_DOCX}')
    print(f'Velikost: {OUT_DOCX.stat().st_size:,} B')


if __name__ == '__main__':
    main()
