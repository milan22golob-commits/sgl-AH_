# Kdo je Elena? — Elias Rudolf, avtorska pogodba

Kronologija v tem repozitoriju:

## v5 — 2026-08-27 (Milanovi vsebinski popravki + sinoptična primerjava)

Vir: Milanov pregled v4, 27. 8. 2026 (osem vprašanj po vrsti).

### Milanovi popravki v5 vs v4

| # | Kje | Sprememba | Razlog |
|---|---|---|---|
| Q1 | 1. člen | Odstavek o 16. čl. ZOdv **skrajšan** — samo omemba "preko svoje registrirane pisarne" | Pravno ni obveznost do ŠGL; sklic na ZOdv lahko zbudi nepotrebno pozornost |
| Q2 | 1. člen | "marca 2027" → **"predvidoma marca 2027"** | Točnega datuma še ni |
| Q3 | 5. člen | Enoten 2.300 EUR → **razdeljen 600 (besedilo) + 1.700 (režija)** | DP-pravnik v REVIEW-u priporočil kot davčno varovalko (ZASP 81. — "primeren honorar") |
| Q4 | 10. člen | Izročanje za izvedbo → **izročanje ARHIVSKIH kopij** (naročnik prosi, ne zahteva) | Elias nima produkcijskih materialov; logika po Nini MG 18. čl. |
| Q5 | 9. člen, podpis | "avtor besedila in režiser" → **"dramatik in režiser"** | Utečena slovenska sintagma; "avtor besedila" ni v ŠGL slovarju |
| Q6 | 11. člen | "edini avtor obeh del" → **razdeljeno** (a) izključni avtor besedila, (b) zaveza za izključno režijo | Besedilo že obstaja, režija bo nastala — jamstvo se natančno navezuje |
| Q7 | 12. člen | Sporazumen razpored → **koordinira organizator kulturnega programa** sporazumno | Vrnjena Milanova omemba iz Ira/Šunder, a mehčano |
| Q8 | 8. člen | Neekskluzivno **ostaja** | Fair za mladega avtorja; ekskluzivnost bi zahtevala višji honorar |

### Odprta vprašanja pred podpisom

1. **Ekskluzivnost** — če ŠGL hoče, da Elias v času repertoarja NE ponudi drugim slovenskim gledališčem, se doda ločena klavzula. Trenutno neekskluzivno.
2. **Milanovi ročni popravki** — Milan je omenil popravke v dokumentu iz zadnjega datuma. Čakajo posredovanja (upload / povezava / opis).

### Nova datoteka: sinoptična primerjava

`sgl_make_primerjava_ira_nina_elias.py` generira ležečo tabelo (5 stolpcev, 16 vrstic):

| DEL | IRA (Šunder) | NINA (Menažerija) | ELIAS v5 | KOMENTAR |

Pokaže po ključnih delih pogodbe (uvod strank, naslov, 1.–24. člen, podpis, izvodi), kako je isti del napisan pri vseh treh, s komentarjem v zelenem v petem stolpcu. Zaključek: 4 razdelki o razlikah (Elias vs Nina, Elias vs Ira, skupne poteze, kje je Elias najdlje).

## v4 — 2026-08-27 (aplikacija manjkajočih pravil iz standarda v7)

Vir: `ŠG_STANDARDI_pogodbe_v7_changelog.md` (potrjeno 20. 8. 2026),  
`LECNI_STANDARD_pregled_pogodb.md`, `OBLIKOVALSKA_TAKSONOMIJA_pogodb.md`.

### Glavno spoznanje

v3 je bil zgrajen po **primerih** (Nina MG, Šunder), **ne po standardu** —  
zato je marsikatera Milanova utrjena norma iz v7 obšla. v4 aplicira te norme.

### Kaj se je popravilo (v3 → v4)

| # | Kje | v3 | v4 | Podlaga |
|---|---|---|---|---|
| P1 | Uvod strank | "Elias Rudolf, odvetnik" | **"g. Elias Rudolf, odvetnik"** | v7 §3.1 (g./ga. samo v uvodu strank); §3.3 (naziv pred imenom) |
| P2 | III. poglavje ime | "AVTORSKI HONORAR" | **"NARAVA RAZMERJA"** | v7 §1.1 tabela; §8 changelog #2 |
| P3 | 1. člen datum premiere | "v marcu 2027" plain | **"marca 2027" BOLD** | Oblikovalska taksonomija B.1 #6 |
| P4 | Podpisni blok | ročno grajen, imena v vrstici 6–7, brez bolda, žig samo desno | **po v7 §4.0**: vrstica 4 naziv, **vrstica 5 IME BOLD**, vrstica 6 funkcija, **vrstica 8 žig OBOJESTRANSKO** | v7 §4.0 (Milanov popravek 20. 8. — izrecno svari pred to napako); bold register B.1 #1, #3 |
| P5 | Struktura | (nejasno) | **jasno DVOSTRANSKA** (dvojina glagolov; avtor = izvajalec = ista F.O. z registrirano dejavnostjo) | ZOdv 16. čl.; ne §3.2, ker Elias nima ločenega pravnega subjekta |

### GLASNA OPOZORILA (v DIFF .docx)

**O1 — Odvisnost od lokalne `sgl_docx_format.py`**  
Standard v7 je 20. 8. 2026 popravil dve stvari v helperju:
- **§5.6.2**: zamiki poglavij 0,63 / −1,27 → **0,75 / −0,75** cm (rimska št. poravnana z levim robom telesa).
- **§5.6.10**: alineje zdaj z označevalcem **"–" (en-dash) + presledek** (prej brez marker-ja).

Kopija `sgl_docx_format.py` na Drive je iz **7. 5. 2026** — stara. Če Milanova lokalna kopija ni novejša, poglavja in alineje bodo videti drugače od predpisa.

**O2 — Odvetnik ≠ pravna oseba**  
v7 §3.2 predpisuje tristranski vrstni red samo za "sodelavski Potodom" (kjer vmesnik je d.o.o./s.p./zavod). Elias kot samostojni odvetnik po ZOdv posluje kot F.O. z registrirano poklicno dejavnostjo — **NI ločena pravna oseba**. Zato dvostranska pogodba. Odstavek v 1. členu, ki pojasnjuje dvojno identiteto (avtor F.O. + izvajalec preko pisarne), je Elias-specifika — ni predpisan v v7, predlog za v8.

**O3 — Podpisni blok v v3 je bil STRUKTURNO NAPAČEN**  
v7 §4.0 (Milan, 20. 8. — dobesedni citat): *"vrstice tabele: 0 datum · 2 oznaka · 4 naziv prav. osebe · **5 ime+priimek (bold)** · 6 funkcija · 8 žig"*. Moj v3 je imel imena v vrstici 6/7 in **brez bolda**. To je natanko napaka, ki jo standard izrecno omenja kot ponavljajočo. **Popravljeno v v4.**

### Datoteke v4

- `sgl_make_pogodba_elias_rudolf_v4.py` — čista produkcijska verzija za podpis
- `sgl_make_pogodba_elias_v4_DIFF.py` — diff verzija z vidnimi popravki + oranžnimi opozorili (NI za podpis; za primerjavo)

## v3 — 2026-08-27 (po pregledu Bizija)

Vir podatkov: Bizi.si / AJPES PRS.

### Podatki izvajalca (izpolnjeni; prazna polja iz v2 zaprta)

| | |
|---|---|
| Naziv | ELIAS RUDOLF - ODVETNIK |
| Pravna oblika | samostojni odvetnik (SKIS S.14100 – samozaposleni delodajalec) |
| Sedež | Obrežna steza 2, 1000 Ljubljana |
| Matična št. | 2861887000 |
| Davčna št. | 34928138 |
| DDV | NE — 76.a čl. ZDDV-1 (atipični davčni zavezanec) |
| TRR (uporabljen v pogodbi) | SI56 0400 0028 0838 318 (OTP banka d.d.) |
| Drugi TRR | SI56 0400 0028 0838 221 (OTP banka d.d.) — po želji zamenljiv |
| Datum vpisa | 13. 11. 2023 |
| Dejavnost | Odvetništvo |

### Ključne spremembe v3 vs v2

| # | Kje | Sprememba | Podlaga |
|---|---|---|---|
| 1 | Uvod strank | Izvajalec zdaj **samostojni odvetnik** s polnim naslovom + matično + davčno + TRR (ne več prazna polja z "davčnim uradom" in "statusom") | Bizi.si |
| 2 | 1. člen | **Dodan odstavek**: pojasnilo, da izvajalec nastopa kot fizična oseba-avtor in delo opravlja preko svoje pisarne (16. čl. Zakona o odvetništvu — dovoljena umetniška dejavnost) | Zakon o odvetništvu, ZASP |
| 3 | 5. člen — **honorar** | Odpade hibridna izplačilna klavzula; ostane samo **izvajalec izda račun**. Rok: 30 dni od prejema računa (namesto 30 dni po premieri). "2.300,00 EUR" brez sufiksa "bruto" (ker gre za znesek na računu brez DDV). | Bizi.si — 76.a čl. ZDDV-1 |
| 4 | 24. člen, podpis | "Elias Rudolf, odvetnik" namesto "Elias Rudolf" | Uradni naziv poklicne oznake |

### Kar OSTAJA iz v2

- Varovalka avtorja (8. čl.) — pravice na besedilu kot samostojnem literarnem delu.
- Vse popravke iz REVIEW-a (5. čl. ZASP za režijo, specifični členi prenosa MAP, reverzija, opomba o sorodnih pravicah tretjih, mehčanje delovno-pravnih indicij).

### Priloga A (ZPIZ) NI potrebna

- Elias kot registrirani samozaposleni delodajalec (SKIS S.14100) sam plačuje  
  PIZ in ZZ prispevke; ne potrebuje ločene izjave zavarovanca po 18. čl. ZPIZ-2.

## v2 — 2026-08-27
- Osnova: **Nina Šorak MG** (20. 8. 2026) kot trenutni gold standard.
- Vgrajeni popravki iz **REVIEW-a treh pravnikov (DP / AP / ZP)** k v1_PREDLOG.

### Ključne spremembe v2 vs v1_PREDLOG

| # | Člen | Sprememba | Podlaga |
|---|---|---|---|
| 1 | 3. člen | Režija sklic na **5. člen ZASP** (ne 80.) | AP-review (80. čl. je za PRENOS, ne za nastanek režije) |
| 2 | 4. člen | "Vodenje igralske ekipe" → "koordinacija"; brez "prednostno redno zaposlene" | DP-review (mehčanje indicij delovnega razmerja) |
| 3 | 5. člen | Ohranjena hibridna izplačilna klavzula (račun ali obračun naročnika) | DP-review |
| 4 | 6. člen | Prenos MAP s **specifičnimi ZASP členi**: 23., 24., 26., 30., 32.a | AP-review (75. čl. zahteva izrecno navedbo) |
| 5 | 7. člen | Trajno neizključno za dokumentacijo/arhiv + izrecna **REVERZIJA** ob umiku | AP-review (ZASP 78./83.); Nina MG standard |
| 6 | 8. člen | Varovalka avtorja ohranjena (nespremenjena) | Elias-specifika (izvirno delo) |
| 7 | 11. člen | Dodana **opomba o sorodnih pravicah tretjih** (igralci, glasba, scenografija — ločene pogodbe) | AP-review (igralci = 118. čl. ZASP) |
| 8 | 12. člen | "Razpored vaj določen s tedenskim programom" → **sporazumno** | DP-review (odsotnost podrejenosti) |
| 9 | 13. člen | Brez izrecnega "hišnega reda"; splošna gospodarnost + splošna pravila v prostorih | DP-review |
| 10 | 15. člen | Odsotnost/nadomestni ustvarjalec — **sporazumno** (Nina MG dikcija) | Nina MG standard |
| 11 | 16. člen | Obveščanje **naročnika** (ne "inšpicientu takoj") | DP-review |

### Kar OSTAJA iz v1_PREDLOG (in ne pride od Nine)

- **8. člen — varovalka avtorja** (obdrži pravice na besedilu kot samostojnem literarnem delu).  
  Nina te klavzule nima, ker njena vloga ni izvirno avtorstvo. Za Eliasa je bistvena.
- **11. člen — jamstvo izvirnosti** brez navezave na tuji vir.  
  Za razliko od Šunder v2 (Ira Ratej), ki se je pravno branil kot "po motivih Frayna, ne priredba".
- **Honorar 2.300,00 EUR bruto skupaj za obe vlogi**, DDV se ne obračuna (F.O., ne d.o.o. vmesnika).

### Kar iz Nine MG **ni** vključeno (zavestno)

- **Produkcijski okvir** in samostojno razporejanje materialov: pri Eliasu ni  
  relevantno, ker avtor besedila + režiser ne kupuje kostumov / rekvizitov.  
- **Kostumska specifika**: predelava, izposoja od tretjih, prevzem po umiku.
- **Izplačilo v dveh delih (30 + 60 dni)** vezano na zaključni obračun: pri  
  Eliasu enkratno izplačilo 30 dni po premieri (enostavnejši scenarij).

### Odprta vprašanja pred podpisom

1. **Priloga A (izjava ZPIZ po 18. čl. ZPIZ-2)** — če je Elias F.O. brez  
   statusa (npr. brez s.p. ali samozaposlitve v kulturi), dodaj Prilogo A po  
   vzoru Likar (Prah, avgust 2026).
2. **Notranja razčlemba honorarja** (X za besedilo, Y za režijo) — DP-review  
   priporoča kot davčno/pravno varovalko. Trenutno je znesek enoten.
3. **Bruto vs. bruto-bruto** — v tekstu je izrecno "2.300,00 EUR bruto". Pri  
   F.O. brez statusa naročnik nosi bruto-bruto strošek; pojasnilo je smiselno  
   dodati v obrazložitev za izplačilo.

## v1_PREDLOG — 2026-08-14
- Generator: `sgl_make_pogodba_elias_rudolf.py` (Google Drive)
- Vzporedno: `sgl_make_pogodba_elias_REVIEW.py` — vgrajeni pravni komentarji  
  treh perspektiv (DELOVNO-PRAVNO, AVTORSKO-PRAVNO, ZGODOVINSKO-PRAVNO).
