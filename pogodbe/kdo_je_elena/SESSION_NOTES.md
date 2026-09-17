# SESSION_NOTES — Elias Rudolf, "Kdo je Elena?"

Kompresiran povzetek Claude Code Remote seje z Milanom (27. 8. 2026).
Za novo sejo, ki bo prevzela to delo — beri to + CHANGELOG.md, ne
celotnega pogovora (bil je ~500 K tokenov).

## Kaj je "Kdo je Elena?"

- **Predstava**: krstna uprizoritev, ŠGL, **predvidoma marca 2027**
  (točnega datuma še ni, Elias je popravil datum pogodbe na 7. 9. 2026).
- **Avtor + režiser** je ista oseba: **Elias Rudolf**, samostojni odvetnik
  (registrirana pisarna), hkrati AVTOR IZVIRNEGA DRAMSKEGA BESEDILA in
  REŽISER. Poleg tega tudi igralec — po Milanu "najboljši igralec srednje
  generacije".
- **Honorar**: 2.300,00 EUR skupaj, razdeljen po vlogah:
  600,00 EUR za besedilo + 1.700,00 EUR za režijo.

## Elias — poslovni profil (Bizi, 27. 8. 2026)

| | |
|---|---|
| Naziv | ELIAS RUDOLF - ODVETNIK |
| Sedež | Obrežna steza 2, 1000 Ljubljana |
| Matična št. | 2861887000 |
| Davčna št. | 34928138 |
| DDV | NE (klasifikacija "76.a čl. ZDDV-1" — glej OPOZORILO spodaj) |
| TRR (uporabljen v v5) | SI56 0400 0028 0838 318 (OTP d.d.) |
| Drugi TRR | SI56 0400 0028 0838 221 (OTP d.d.) |
| Datum vpisa | 13. 11. 2023 |
| Dejavnost | Odvetništvo |

## Trenutno stanje: v5

Aktivna verzija je **v5** (`sgl_make_pogodba_elias_rudolf_v5.py`).
Datum v pogodbi: **7. 9. 2026** (Milan popravil od 1. 9.).

Datoteke v mapi:
- `sgl_make_pogodba_elias_rudolf_v5.py` — **čista produkcijska**
- `sgl_make_pogodba_elias_v5_LECA_CITATI.py` — leča z verbatim citati
- `sgl_make_pogodba_elias_v5_DIFF.py` (v4 DIFF) — vidni popravki v3→v4
- `sgl_make_pogodba_elias_REVIEW_v3.py` — trije pravniki (DP/AP/ZP komentarji)
- `sgl_make_primerjava_ira_nina_elias.py` — sinoptična primerjava, 16 vrstic
- CHANGELOG.md — polna evolucija v1 → v5

## Odprta vprašanja pred podpisom

1. **Ekskluzivnost drugim slovenskim gledališčem** (8. čl.). Trenutno
   **NEEKSKLUZIVNO** — Elias lahko besedilo ponudi tudi drugod med tem, ko
   je pri ŠGL na repertoarju. Če ŠGL hoče ekskluzivo → dodati klavzulo +
   verjetno višji honorar.

2. **Sklic na 76.a čl. ZDDV-1** (uvod strank + 5. čl.). Leča razkrila,
   da 76.a se nanaša na **reverse charge**, ne na oprostitev. Bizi ga
   uporablja kot klasifikacijo. Pravilnejši pravni razlog za NE-obračun
   DDV pri avtorski storitvi je verjetno:
   - **94. čl. ZDDV-1** (mali davčni zavezanec, prag 50.000 EUR), ali
   - **44. čl. ZDDV-1** (oprostitev — točka 12 duhovne storitve
     pisateljev/skladateljev), ali
   - brez sklica: "Izvajalec ni davčni zavezanec za DDV, zato se DDV ne
     obračuna."
   Elias naj kot odvetnik sam preveri in potrdi.

3. **Klavzula o registrirani pisarni** (1. čl., 3. odst.). Pravna podlaga
   je **21. čl. ZOdv** (točka 1: "razen v znanstveni, pedagoški, umetniški
   ali publicistični dejavnosti"). Klavzula ni pravna obveznost do ŠGL —
   lahko se izpusti. Trenutno je kot skrajšana omemba brez sklica.

4. **"Mehčanje" delovno-pravnih indicij** (12., 13., 15., 16. čl.).
   **NI Milanova izrecna zahteva** — prišlo iz REVIEW DP-pravnika
   (14. 8., ZDR-1 13.a čl. tveganje rekvalifikacije). Nina MG uporablja
   mehčano dikcijo, Ira/Šunder uporablja trše. Če Milan hoče trše
   Ira-dikcije, se v v6 vrne.

5. **Ročni popravki v prejšnjem dokumentu** — Milan je omenjal popravke,
   nato rekel "morda pa nisem spreminjal". Zaenkrat brez akcije.

## Moje napake, razkrite z lečo (POPRAVITI v v6)

| # | V | Napačno | Pravilno |
|---|---|---|---|
| 1 | v3 CHANGELOG, v4 DIFF | 16. čl. ZOdv | **21. čl. ZOdv** (16. je o statusu specialista, ne o umetniški dejavnosti) |
| 2 | v3/v4/v5 uvod + 5. čl. | 76.a čl. ZDDV-1 kot razlog za NE-DDV | **94. čl.** ali **44. čl.** ali brez sklica (glej odprto vprašanje 2) |

## Popravki za standard v7 (predlagani v tej seji, čakajo Milan-potrditve)

Patch v `pogodbe/_tools/sgl_docx_format_PATCH_20260827.py`:

- **§5.6.10 (popravek)**: alineje z zamikom 1,25 / −0,50 cm (prej 0,75 / −0,75
  — Milanova opomba, da je prejšnja postavitev nepregledna).
- **§5.6.12 (novo)**: `set_metadata_milan(doc)` → author = "Milan Golob".

Aplikacija: nadomesti `add_bullet()` in dodaj `set_metadata_milan()` v
lokalni `C:\Users\ladmin\Moj disk\KODE\_tools\sgl_docx_format.py`.

## Elias vs Ira vs Nina — kje se Elias razlikuje

Glej sinoptično tabelo v `sgl_make_primerjava_ira_nina_elias.py` (16
vrstic) + zaključno primerjavo v CHANGELOG.md.

Kratek povzetek:
- **Ira** (Šunder v dvorani) = 3 vloge (avtor priredbe + režija + glasba),
  preko Arahne d.o.o. vmesnika z DDV.
- **Nina** (Steklena menažerija) = 2 vlogi (režija + kostumografija),
  SVK samostojna z produkcijskim okvirom 6.000 EUR.
- **Elias** (Kdo je Elena?) = 2 vlogi (avtor besedila + režija),
  samostojni odvetnik z DDV NE, fiksen honorar 2.300 EUR razdeljen po
  vlogah. **Edinstven po varovalki avtorja besedila (8. čl.)**, ker je
  edini imetnik izvirnega besedila v vseh treh primerih.

## Naslednji koraki

**Če se pogovor konsolidira v novo sejo, ki že dela pogodbe**:

1. Milan naj Eliasu pošlje **v5** (`ŠGL_avtorska pogodba_KDO JE
   ELENA_Elias Rudolf_dramatik-režiser_v5.docx`) po e-pošti (osnutek
   maila je v prejšnji seji, Milan ga ima).
2. Počakaj na Eliasovo mnenje.
3. Če Elias potrdi ali predlaga popravke — v6:
   - Vključi Eliasove pripombe
   - Popravi 76.a → 94/44 (glej odprto 2)
   - Odloči ekskluzivnost (glej odprto 1)
   - Odloči mehčanje (glej odprto 4)
   - Popravi CHANGELOG (16. → 21. čl. ZOdv)

## Kaj je konsolidacija z drugimi pogodbami

Ta seja je delala SAMO Elias/Kdo je Elena. Druge pogodbe (Nina, Ira, itd)
so bile prebrane kot referenca, ampak ne obdelane. Če nova seja hkrati
dela druge pogodbe, naj:
- Elias v5 obravnava kot samostojen artefakt (mapa `kdo_je_elena/`)
- Uporabi popravke standarda v7 (patch v `pogodbe/_tools/`) za VSE
  produkcije, ki jih obdeluje
- Sinoptična primerjava (`sgl_make_primerjava_ira_nina_elias.py`) je
  koristna referenca ne glede na produkcijo — Elias vzorec (izvirni
  avtor + režiser) je nov v ŠGL korpusu.
