# -*- coding: utf-8 -*-
"""
sgl_make_pogodba_elias_REVIEW_v3.py
==================================
Regenerira Eliasovo pogodbo v3 (podatki iz Bizija, izvajalec izda račun) z
VGRAJENIMI pravnimi komentarji treh perspektiv, kot barvni inline tekst za
vsakim členom.

Perspektive:
- DELOVNO-PRAVNO (modra, 1F4E79) — ZDR-1, ZPIZ-2, davki, tveganje rekvalifikacije
- AVTORSKO-PRAVNO (zelena, 2E7D32) — ZASP, obseg prenosa, moralne pravice
- ZGODOVINSKO-PRAVNO (rdeča, C00000) — primerjava z Nina MG, Šunder v2, Prah,
  Revizor, TKV in vpliv sprememb v3 vs v1_PREDLOG

Output:
  Šentjakobsko gledališče\\Pogodbe in računi\\Kdo je Elena\\
  ŠGL_avtorska pogodba_KDO JE ELENA_Elias Rudolf_v3_REVIEW-3pravniki.docx
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
OUT_DOCX = OUT_DIR / 'ŠGL_avtorska pogodba_KDO JE ELENA_Elias Rudolf_v3_REVIEW-3pravniki.docx'

COL = {'DP': '1F4E79', 'AP': '2E7D32', 'ZP': 'C00000'}
LBL = {'DP': 'DELOVNO-PRAVNO', 'AP': 'AVTORSKO-PRAVNO', 'ZP': 'ZGODOVINSKO'}

# Komentarji treh perspektiv, prilagojeni v3 (izvajalec = samostojni odvetnik,
# izda račun, hibridna izplačilna klavzula odpade, popravki iz v1 REVIEW-a
# vgrajeni v besedilo).
COMMENTS = {
    1: {
        'AP': 'Sklic na 5. čl. ZASP za dramsko besedilo je pravilen. Trditev "ni priredba" ostane pravno močna (individualnost po 5. čl.).',
        'DP': 'v3 DODATNO ZAVAROVANO: pojasnilo o statusu izvajalca (odvetnik z registrirano dejavnostjo, delo po 16. čl. ZOdv) močno zmanjša tveganje rekvalifikacije v delovno razmerje — izvajalec je poslovni subjekt s samostojno organizacijo dela.',
        'ZP': 'NOVOST v3 — odstavek o odvetniški dejavnosti in 16. čl. ZOdv nima analogije v ŠGL korpusu (niti Nina MG, niti Šunder, niti Prah/Likar). Specifično za Eliasa kot odvetnika-avtorja.',
    },
    2: {
        'ZP': 'Identičen Nina MG (2. čl.) in Šunder v2. Novejši standard "višja sila", brez zastarelega sklica na COVID-19 (Revizor 2023).',
    },
    3: {
        'AP': 'v3 POPRAVLJEN glede v1: režija zdaj sklic na 5. čl. ZASP (ne 80. čl.). Pravilno — 80. čl. ureja PRENOS pravice (avtorska pogodba), režija je avtorsko delo po 5. čl.',
        'DP': 'Razčlenitev honorarja med vlogama (X za besedilo, Y za režijo) v pogodbi ni izvedena — priporočilo iz v1 REVIEW-a ostane odprto. Ob morebitnem sporu je notranja razčlemba (npr. v prilogi ali interni evidenci naročnika) koristna dokazna varovalka.',
        'ZP': 'Sprememba v3 = neposredna implementacija AP-priporočila iz v1 REVIEW-a. Zdaj sledi Nini MG (obe vlogi pod 5. čl. ZASP).',
    },
    4: {
        'DP': 'v3 OMILJENO: "vodenje igralske ekipe" → "koordinacija dela soustvarjalcev"; "prednostno redno zaposlene" → "sodeluje z direktorjem in upošteva zmožnosti naročnika". Odstranjeni najbolj izpostavljeni indici odvisnega razmerja (ZDR-1 13. čl.). Za samostojnega odvetnika kot izvajalca je to zdaj v mejah normalnega umetniškega vodenja.',
        'AP': 'Brez pripomb (vsebinske obveznosti, ne prenos pravic).',
        'ZP': 'Nina MG dikcija ("koordinacija", ne "vodenje").',
    },
    5: {
        'DP': 'KLJUČNA IZBOLJŠAVA v3: hibridna izplačilna klavzula odpade. Izvajalec izda račun kot samostojni odvetnik po 76.a čl. ZDDV-1 (brez DDV). Odstranjena dvoumnost "bruto vs bruto-bruto" iz v1/v2 — 2.300,00 EUR je zdaj enoznačno "znesek na računu brez DDV". Priporočilo: če se strankama zdi smiselno, ločen dodatek/priloga z notranjo razčlenitvijo po vlogah.',
        'AP': 'Honorar/davki so izven ZASP. Edina opomba: enoten honorar za obe vlogi je pri sporu o "primernem honorarju" (ZASP 81.) manj obranljiv — okvirna razdelitev besedilo/režija pomaga.',
        'ZP': 'v3 = model "Nina MG" (samostojna izdaja računa), NE več Prah/Likar (F.O. brez statusa) ali Šunder/TKV (Arahne d.o.o. z DDV). Priloga A (izjava ZPIZ) NI potrebna — Elias sam plačuje PIZ/ZZ prispevke iz svoje dejavnosti.',
    },
    6: {
        'AP': 'v3 POPRAVLJEN: seznam prenesenih pravic zdaj vsebuje sklice na KONKRETNE ZASP člene (23. reproduciranje, 24. distribuiranje, 26. javna izvedba, 30. radiodifuzija, 32.a dajanje na voljo javnosti za splet). Skladno z zahtevo 75. čl. ZASP po izrecni navedbi. Ostala kar-nenavedeno pravilo (76. čl.) tako manj tvegano.',
        'ZP': 'Bolj podrobno kot Nina MG (ki našteje pravice opisno, brez sklicev na člene). Neposredna implementacija AP-priporočila iz v1 REVIEW-a.',
    },
    7: {
        'AP': 'v3 DOPOLNJEN: drugi odstavek uvaja REVERZIJO — z umikom uprizoritve z repertoarja se pravice samodejno vrnejo izvajalcu. Rešitev pravne praznine (ZASP 78./83.); dobra zaščita avtorja. Naročnik ohrani trajno pravico za dokumentacijo in že nastale fotografije/posnetke.',
        'ZP': 'Nina MG (trajanje + trajni prenos za dokumentacijo) + reverzija (AP-priporočilo v1 REVIEW-a). Boljša v3 kombinacija.',
    },
    8: {
        'AP': 'KLJUČNA VAROVALKA — nespremenjena iz v1/v2. Izvajalec ostane imetnik pravic na besedilu kot samostojnem literarnem delu; prenos je ozko vezan na uprizoritev naročnika. Skladno z ZASP 75./76. (ozka razlaga). Najmočnejši člen pogodbe.',
        'ZP': 'Elias-specifika — nima analogije v Nini MG (kostumografija) ne v Šunderju (Ratej ni avtorica besedila, samo priredovalka). Ekvivalent Likarjine varovalke (prevajalka obdrži pravice na prevodu).',
    },
    9: {
        'AP': 'Pravilno — moralne pravice (ZASP 16.-20.) so neodtujljive. Navedba avtorja (18. čl.), upor skazitvi (19. čl.).',
        'ZP': 'Identično Nini MG (9. čl.) in Šunder v2. Standardna dikcija.',
    },
    10: {
        'AP': 'Pravilno razlikovanje med lastništvom NOSILCA (rokopis = stvar) in avtorsko pravico (ZASP 39.). Konsistentno z 8. čl.',
        'ZP': 'Prevzeto iz Nina MG.',
    },
    11: {
        'AP': 'v3 DOPOLNJEN: dodan drugi odstavek o SORODNIH pravicah tretjih (igralci 118. čl. ZASP, skladatelj, scenograf, kostumograf, lektor). Pojasni, da jamstvo "edini avtor" pokriva samo dela izvajalca, ne dela drugih soustvarjalcev, katere ureja naročnik z ločenimi pogodbami. Priporočilo iz v1 REVIEW-a implementirano.',
        'DP': 'Brez neposredne DP-opombe.',
        'ZP': 'Jamstvo izvirnosti brez navezave na tuji vir (za razliko od Šunder v2, kjer je bilo treba braniti "po motivih Frayna, ne priredba"). Klavzula o sorodnih pravicah = novost v v3, konsistentna s praksami profesionalnih gledališč.',
    },
    12: {
        'DP': 'KLJUČNI POPRAVEK v3: "razpored vaj SPORAZUMNO" (ne več "določen s tedenskim programom gledališča"). Odstrani vezanost na razpored naročnika = odstrani enega najmočnejših indicij delovnega razmerja (podrejenost razporeju). Za samostojnega odvetnika-izvajalca je zdaj model dvostranskega dogovora, ne enostranskega naročnikovega ukaza.',
        'ZP': 'Nina MG dikcija, ne več Šunder v2 / Revizor.',
    },
    13: {
        'DP': 'v3 OMILJENO: brez izrecnega sklica na "hišni red", nadomestno "splošna pravila, ki veljajo v prostorih naročnika". Manj vezano na režim zaposlenca; primerno za poslovnega izvajalca. Kombinacija s 4. čl. (koordinacija) in 5. čl. (izdaja računa) občutno zmanjša kumulativni profil odvisnega razmerja.',
        'ZP': 'Odklon od Šunder v2/Nina MG ("hišni red"); v3 je bolj rafinirano.',
    },
    14: {
        'DP': 'Že v v2 mehčano ("v obsegu, ki je običajen"). Sodelovanje pri promociji je projektno vezano, ne splošna delovna dolžnost.',
        'AP': 'Navedba v gledališkem listu izpolnjuje obveznost priznanja avtorstva (18. čl. ZASP).',
        'ZP': 'Nina MG dikcija.',
    },
    15: {
        'DP': 'KLJUČNI POPRAVEK v3: pri odsotnosti se "sporazumno določita nadomestna rešitev" (ne več enostransko naročnikova določitev). Pri objektivni nezmožnosti izvajalca dokončati delo se nadomestni ustvarjalec določi "sporazumno" — spet dvostransko. Odstrani indice pretirane vezanosti na naročnika.',
        'AP': 'Pravica do sorazmernega honorarja ohranjena; naročnik zadrži pravico uporabe že ustvarjenih zasnov in besedila za sorazmerni izplačani del. Konsistentno z 8. čl. (pravice na neizplačanem delu ostanejo avtorju).',
        'ZP': 'Nina MG dikcija ("sporazumno določita"), ne Šunder v2 dikcija (enostransko).',
    },
    16: {
        'DP': 'POPRAVEK v3: "sporočil NAROČNIKU" (ne več "inšpicientu predstave ali organizatorju kulturnega programa NAROČNIKA TAKOJ"). Odpravi element delovne discipline (poročanje o prisotnosti operativnemu vodji), ohrani pa obveznost pravočasnega obvestila naročniku o zamiku plana. Delovno-pravno bistveno čistejše.',
        'ZP': 'Odklon od Šunder v2/Revizor. Bolj primerno za profesionalno pogodbeno razmerje.',
    },
    17: {
        'AP': 'Brez pripomb.',
        'ZP': 'Standardna dikcija (Nina MG, Šunder).',
    },
    18: {
        'AP': 'Pomembno in pravilno: če do uprizoritve ne pride, pravice na besedilu v celoti ostanejo izvajalcu. Konsistentno z 8. čl. in z ZASP 83. (vrnitev pravic ob neizvrševanju).',
        'ZP': 'Nina MG dikcija + dodatek o pravicah na besedilu (Elias-specifika, konsistentno z 8. čl.).',
    },
    19: {
        'ZP': 'Standardna klavzula nevtralnega spola (v vseh ŠGL v7 pogodbah).',
    },
    20: {
        'DP': 'Zaupnost zneska je nevtralna in koristna (zmanjša primerljivost z drugimi izplačili).',
        'ZP': 'Standardno v vseh pogodbah v7.',
    },
    21: {
        'ZP': 'Novejši standard z izrecno navedbo GDPR + ZVOP-2 (kot Nina MG, Likar).',
    },
    22: {
        'AP': 'Pisnost aneksov ustreza zahtevi po pisni obliki avtorske pogodbe (ZASP 80. čl.).',
        'ZP': 'Standardno v vseh pogodbah.',
    },
    23: {
        'DP': 'OPOZORILO: ob morebitnem sporu o naravi razmerja (če bi bila uveljavljena domneva delovnega razmerja) lahko stvarno pristojnost prevzame DELOVNO sodišče — ta klavzula tega ne izključi. To ni napaka pogodbe; je splošna omejitev kolizijskih pravil.',
        'ZP': 'Standardno v vseh pogodbah.',
    },
    24: {
        'ZP': 'Trije izvodi (1 izvajalec, 2 naročnik) — identično Nina MG in Šunder v2. Priloga A NI potrebna (Elias sam plačuje prispevke iz registrirane dejavnosti).',
    },
}

SPLOSNO = {
    'DP': (
        'v3 je glede delovno-pravnega tveganja bistveno izboljšana proti v1_PREDLOG: '
        '(1) izvajalec je registrirani samostojni odvetnik, ki dela opravlja preko svoje pisarne — poslovni subjekt, ne fizična oseba brez statusa; '
        '(2) izda račun brez DDV (76.a čl. ZDDV-1) — čista poslovna transakcija; '
        '(3) razpored vaj SPORAZUMNO (12. čl.), koordinacija (ne vodenje) v 4. čl., brez izrecnega "hišnega reda" (13. čl.), obveščanje naročnika (ne inšpicienta takoj) v 16. čl. '
        'Kumulativni profil odvisnega razmerja (ZDR-1 13. čl.) je zdaj NIZEK. Odprto priporočilo: notranja razčlemba honorarja po vlogah (X besedilo, Y režija) za lažjo obrambo v primeru spora.'
    ),
    'AP': (
        'v3 je zdaj nadpovprečno dobra glede ključne avtorsko-pravne osi: '
        '(1) 3. čl. pravilno sklicuje 5. čl. ZASP za obe vlogi (popravek napake v1); '
        '(2) 6. čl. našteva prenesene pravice z izrecnimi sklici na ZASP člene (23., 24., 26., 30., 32.a) — zadošča zahtevi 75. čl. po izrecnosti; '
        '(3) 7. čl. uvaja reverzijsko klavzulo (pravice se vrnejo ob umiku); '
        '(4) 8. čl. varovalka avtorja ostane največja moč (avtor obdrži pravice na besedilu kot samostojnem literarnem delu); '
        '(5) 11. čl. jamstvo dopolnjeno z opombo o sorodnih pravicah tretjih. '
        'Vsi popravki iz v1 REVIEW-a implementirani. Ni več znanih AP-pomanjkljivosti.'
    ),
    'ZP': (
        'v3 je hibrid, ki nima direktne predhodnice v korpusu: '
        '(1) delovno-pravna in izplačilna struktura sledi Nini MG (samostojna, izda račun, sporazumni razpored); '
        '(2) avtorsko-pravna struktura razširjena z Eliasovo specifiko (varovalka 8. čl., specifikacija ZASP členov 6. čl., reverzija 7. čl.); '
        '(3) 1. čl. z odvetniško statusno klavzulo (16. čl. ZOdv) je popolna novost. '
        'Klasična linija "Šunder v2 + Prah + Nina MG" je s tem obogatena. Priloga A NE — Elias ima registrirano dejavnost. Priporočljivo obdržati v3 kot novo referenčno pogodbo za "avtor izvirnega dramskega dela + režiser, samostojni izvajalec z lastno dejavnostjo".'
    ),
}


def init_doc():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    tmp = OUT_DIR / '_tmp_elias_rev_v3.docx'
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


def add_comment_line(doc, key, besedilo):
    """Barvni inline komentar-paragraf."""
    p = doc.add_paragraph()
    apply_no_spacing(p)
    r1 = p.add_run(f'▸ [{LBL[key]}] ')
    r1.bold = True
    r1.italic = True
    r1.font.size = Pt(8.5)
    r1.font.name = 'Calibri'
    r1.font.color.rgb = RGBColor.from_string(COL[key])
    r2 = p.add_run(besedilo)
    r2.italic = True
    r2.font.size = Pt(8.5)
    r2.font.name = 'Calibri'
    r2.font.color.rgb = RGBColor.from_string(COL[key])
    return p


_clen_counter = [0]


def comments_after_clen(doc):
    """Po add_clen + telesu člena vstavi komentarje za tekoči člen."""
    _clen_counter[0] += 1
    n = _clen_counter[0]
    c = COMMENTS.get(n, {})
    if not c:
        return
    for key in ('DP', 'AP', 'ZP'):
        if c.get(key):
            add_comment_line(doc, key, c[key])


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc, tmp = init_doc()

    set_margins(doc, top_cm=2.5, bottom_cm=2.5, left_cm=2.5, right_cm=2.5)
    add_footer_pagenum(doc)

    # === OPOZORILNA GLAVA ===
    p = doc.add_paragraph()
    apply_no_spacing(p)
    r = p.add_run('PREGLEDNA VERZIJA v3 — pravni komentarji treh perspektiv vgrajeni v besedilo. NI za podpis.')
    r.bold = True
    r.font.size = Pt(9)
    r.font.name = 'Calibri'
    r.font.color.rgb = RGBColor.from_string('808080')
    for key, opis in [
        ('DP', 'delovno-pravni (ZDR-1, ZPIZ-2, davki, ZOdv)'),
        ('AP', 'avtorsko-pravni (ZASP)'),
        ('ZP', 'zgodovinsko-pravni (primerjava z Nina MG, Šunder v2, Prah, in v1 Elias)'),
    ]:
        pl = doc.add_paragraph()
        apply_no_spacing(pl)
        rl = pl.add_run(f'▸ [{LBL[key]}] = {opis}')
        rl.bold = True
        rl.font.size = Pt(8.5)
        rl.font.name = 'Calibri'
        rl.font.color.rgb = RGBColor.from_string(COL[key])
    doc.add_paragraph()

    # === UVOD STRANK ===
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
        ('Elias Rudolf, odvetnik', True),
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

    add_clen(doc)
    add_para_runs(doc, [
        ('Ta pogodba se nanaša na avtorsko delo izvajalca pri krstni uprizoritvi gledališke predstave ', False),
        ('KDO JE ELENA?', True),
        (', katere avtor dramskega besedila in režiser je izvajalec, in ki bo premierno uprizorjena v marcu 2027 na Šentjakobskem odru Šentjakobskega gledališča Ljubljana - društvo (v nadaljevanju: uprizoritev).', False),
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
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Pogodbeni stranki soglašata, da se lahko datum premiere oziroma datumi ponovitev, vaj ali drugih dogodkov v zvezi z uprizoritvijo spremenijo iz razlogov višje sile ali drugih nepredvidenih okoliščin, brez odgovornosti katere koli pogodbene stranke.', False)])
    comments_after_clen(doc)

    # ============== II. PREDMET POGODBE ==============
    add_section(doc, 'PREDMET POGODBE')

    add_clen(doc)
    add_para_runs(doc, [('S to pogodbo izvajalec prevzema pri uprizoritvi dve avtorski vlogi:', False)])
    for b in [
        'avtorstvo izvirnega dramskega dela KDO JE ELENA? (avtorsko delo po 5. členu ZASP);',
        'režijo uprizoritve (avtorsko delo po 5. členu ZASP).',
    ]:
        add_bullet(doc, b)
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec je odgovoren za umetniško vrednost uprizoritve in za usklajeno delovanje z direktorjem gledališča glede izvedbe sprejetega terminskega plana. Izvajalec prevzame izvedbo idejne zasnove uprizoritve in koordinacijo dela soustvarjalcev.', False)])
    add_para_runs(doc, [('Pri izbiri soustvarjalcev in pri finančnih zahtevah, povezanih z uprizoritvijo, izvajalec sodeluje z direktorjem gledališča ter upošteva finančne, tehnične in kadrovske zmožnosti naročnika.', False)])
    comments_after_clen(doc)

    # ============== III. AVTORSKI HONORAR ==============
    add_section(doc, 'AVTORSKI HONORAR')

    add_clen(doc)
    add_para_runs(doc, [
        ('Pogodbeni stranki soglašata, da prejme izvajalec za delo po tej pogodbi, ki obsega obe avtorski vlogi iz 3. člena te pogodbe (avtorstvo dramskega dela in režijo), avtorski honorar v skupni višini ', False),
        ('2.300,00 EUR', True),
        ('. Ker izvajalec ni davčni zavezanec za DDV (76.a člen ZDDV-1), se davek na dodano vrednost ne obračuna.', False),
    ])
    add_para_runs(doc, [('S honorarjem iz prejšnjega odstavka so v celoti poravnane vse obveznosti naročnika do izvajalca iz naslova te pogodbe. V honorarju so všteti vsi materialni stroški, ki jih bo imel izvajalec v zvezi z opravljanjem avtorskega dela po tej pogodbi.', False)])
    add_para_runs(doc, [('Izvajalec izstavi naročniku račun po opravljeni premieri uprizoritve. Naročnik izplača honorar na transakcijski račun izvajalca najkasneje trideseti (30.) dan od prejema računa.', False)])
    comments_after_clen(doc)

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
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Pravice za žive izvedbe in ponovitve iz prejšnjega člena veljajo, dokler je uprizoritev na rednem repertoarju Šentjakobskega gledališča Ljubljana - društvo, oziroma dokler je upravni odbor gledališča formalno ne umakne z repertoarja. Pravice za dokumentiranje in arhiviranje ter uporabo že nastalih fotografij in posnetkov v informativno-promocijskih materialih se prenesejo trajno in neizključno.', False)])
    add_para_runs(doc, [('Z umikom uprizoritve z repertoarja prenesene materialne avtorske pravice za žive izvedbe in ponovitve v celoti preidejo nazaj na izvajalca, brez potrebe po dodatnem pravnem dejanju.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [
        ('Pogodbeni stranki izrecno ugotavljata, da izvajalec ', False),
        ('ostane avtor in imetnik avtorskih pravic na dramskem besedilu', True),
        (' KDO JE ELENA? kot samostojnem literarnem delu. Prenos pravic iz te pogodbe se nanaša izključno na uprizarjanje tega besedila v produkciji naročnika in ne omejuje pravice izvajalca, da svoje dramsko besedilo objavi, ponudi v uprizoritev drugim gledališčem ali kako drugače uporablja zunaj te produkcije.', False),
    ])
    comments_after_clen(doc)

    # ============== V. MORALNE AVTORSKE PRAVICE ==============
    add_section(doc, 'MORALNE AVTORSKE PRAVICE')

    add_clen(doc)
    add_para_runs(doc, [('Moralne avtorske pravice ostanejo izvajalcu v skladu z ZASP, zlasti pravica do priznanja avtorstva in pravica do celovitosti dela.', False)])
    add_para_runs(doc, [('Naročnik bo izvajalca v vseh informativno-promocijskih materialih, gledališkem listu in drugih javnih objavah, povezanih z uprizoritvijo, navedel s polnim imenom in priimkom — Elias Rudolf — z navedbo obeh avtorskih vlog: "avtor besedila in režiser".', False)])
    add_para_runs(doc, [('Izvajalec ima pravico, da se upre vsaki skazitvi, okrnitvi ali drugačni spremembi svojega dela ter pravico, da se upre vsaki izvedbi dela, ki bi žalila njegovo čast in ugled.', False)])
    add_para_runs(doc, [('Izvajalec je seznanjen, da bo njegovo delo predmet tržnega komuniciranja, ki ga za namene obveščanja javnosti izvaja naročnik.', False)])
    comments_after_clen(doc)

    # ============== VI. FIZIČNI NOSILCI IN LASTNINA ==============
    add_section(doc, 'FIZIČNI NOSILCI IN LASTNINA')

    add_clen(doc)
    add_para_runs(doc, [('Fizični in digitalni nosilci, ki jih izvajalec izroči naročniku za namen izvedbe uprizoritve (rokopis dramskega besedila, beleške, režijska knjiga, delovni materiali), ostanejo v hrambi naročnika za čas, dokler je uprizoritev na repertoarju, izključno za namen uprizarjanja, arhiviranja in promocije uprizoritve.', False)])
    add_para_runs(doc, [('Izvirnik dramskega besedila in drugi izvirniki avtorskega dela ostanejo v lasti izvajalca.', False)])
    comments_after_clen(doc)

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
    comments_after_clen(doc)

    # ============== VIII. OBVEZNOSTI IZVAJALCA ==============
    add_section(doc, 'OBVEZNOSTI IZVAJALCA')

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec opravi delo vestno in v rokih, ki jih opredeljuje terminski plan, sprejet sporazumno z direktorjem gledališča.', False)])
    add_para_runs(doc, [('Razpored vaj določita pogodbeni stranki sporazumno, ob upoštevanju splošnega tedenskega programa gledališča in razpoložljivosti soustvarjalcev.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec se v zvezi s prevzetim delom obvezuje ravnati glede rabe sredstev (materialnih, finančnih in kadrovskih) kot skrben gospodar in pri tem upoštevati finančne, tehnične in kadrovske zmožnosti naročnika ter splošna pravila, ki veljajo v prostorih naročnika.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec sodeluje pri tiskovnih konferencah, predstavitvah in drugih promocijskih dejavnostih, povezanih z uprizoritvijo, v obsegu, ki je običajen za tovrstne produkcije, ter prispeva podatke za gledališki list.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Če izvajalec zaradi bolezni ali druge upravičene odsotnosti začasno ne more opravljati dela, o tem nemudoma obvesti naročnika. Pogodbeni stranki sporazumno določita nadomestno rešitev ali prilagodita terminski plan.', False)])
    add_para_runs(doc, [('Če izvajalec prevzetega dela iz objektivnih razlogov (višja sila, huda bolezen ipd.) ne more nadaljevati, naročnik nedokončano delo dokonča z nadomestnim ustvarjalcem, ki ga stranki določita sporazumno, ob varovanju moralnih avtorskih pravic izvajalca. V tem primeru ima izvajalec pravico do honorarja v sorazmerju z že opravljenim delom; naročnik zadrži pravico uporabe že ustvarjenih idejnih zasnov in dramskega besedila za to produkcijo, za katere je bil sorazmerni del honorarja izplačan.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Morebitno odsotnost, ki jo lahko povzroči bolezen ali višja sila in ki lahko povzroči kakršenkoli odlog načrtovanega terminskega plana, izvajalec sporoči naročniku takoj, ko je to mogoče (telefonsko, s SMS ali po elektronski pošti).', False)])
    comments_after_clen(doc)

    # ============== IX. KONČNE DOLOČBE ==============
    add_section(doc, 'KONČNE DOLOČBE')

    add_clen(doc)
    add_para_runs(doc, [('Naročnik lahko odstopi od te pogodbe, če izvajalec svojih obveznosti ne izpolnjuje in s tem naročniku povzroča materialno škodo ali škoduje njegovemu ugledu.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Izvajalec lahko odstopi od te pogodbe, če naročnik brez njegove krivde ne izpelje uprizoritve do premiere. V tem primeru sme izvajalec zadržati že izplačani honorar kot nadomestilo za opravljeno delo do trenutka razdrtja, pravice na dramskem besedilu KDO JE ELENA? pa v celoti ostanejo izvajalcu.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Izrazi za osebe, ki so v tej pogodbi zapisani v moški slovnični obliki, se uporabljajo nevtralno in veljajo za vse spole.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Pogodbeni stranki soglašata, da dogovorjeni znesek avtorskega honorarja ni javen podatek in ga v javnosti ne bosta uporabljali, razen kadar to določa zakon.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Naročnik se obvezuje, da bo osebne podatke izvajalca varoval in obdeloval izključno za namene izvajanja te pogodbe, skladno s Splošno uredbo o varstvu podatkov (GDPR) in veljavnim Zakonom o varstvu osebnih podatkov (ZVOP-2).', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Vse morebitne spremembe in dopolnitve te pogodbe veljajo le v primeru sklenitve pisnega aneksa, ki ga podpišeta obe pogodbeni stranki.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Morebitne spore iz te pogodbe bosta pogodbeni stranki reševali sporazumno. Če to ni mogoče, je za reševanje sporov pristojno stvarno pristojno sodišče v Ljubljani.', False)])
    comments_after_clen(doc)

    add_clen(doc)
    add_para_runs(doc, [('Ta pogodba je sestavljena v treh (3) enakih izvodih, od katerih prejme izvajalec en (1) izvod, naročnik pa dva (2). Pogodba začne veljati z dnem podpisa obeh pogodbenih strank.', False)])
    comments_after_clen(doc)

    # === SPLOŠNI KOMENTARJI (povzetek) ===
    doc.add_paragraph()
    add_section(doc, 'SPLOŠNE UGOTOVITVE PRAVNIKOV (povzetek)')
    for key in ('DP', 'AP', 'ZP'):
        add_comment_line(doc, key, SPLOSNO[key])
    doc.add_paragraph()

    # === PODPISNI BLOK ===
    doc.add_paragraph()
    table = make_table_no_borders(doc, rows=9, cols=2)
    rows_data = [
        ('Datum: 1. september 2026', 'Datum: 1. september 2026'),
        ('', ''),
        ('Izvajalec:', 'Naročnik:'),
        ('', ''),
        ('', 'Šentjakobsko gledališče Ljubljana - društvo'),
        ('', ''),
        ('Elias Rudolf, odvetnik', 'Milan Golob'),
        ('avtor besedila in režiser', 'direktor'),
        ('', 'Žig:'),
    ]
    for i, (l, r) in enumerate(rows_data):
        fill_cell(table.cell(i, 0), l)
        fill_cell(table.cell(i, 1), r)

    doc.save(str(OUT_DOCX))
    if tmp.exists():
        tmp.unlink()
    print(f'GENERIRANO: {OUT_DOCX}')
    print(f'Velikost: {OUT_DOCX.stat().st_size:,} B')
    print(f'Komentiranih členov: {_clen_counter[0]}')


if __name__ == '__main__':
    main()
