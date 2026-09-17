# CLAUDE.md — repo sgl-AH_

Kratek kontekst za Claude Code seje, ki delajo s tem repozitorijem.

## Kaj je repo

`sgl-AH_` = veja delovnih artefaktov za Šentjakobsko gledališče Ljubljana –
društvo (ŠGL). Konkretno pogodbeni artefakti (avtorske pogodbe, pogodbe o
sodelovanju, izjave) — generatorske skripte in dokumentacija.

**Direktor ŠGL:** Milan Golob (`milan22golob@gmail.com`).

## Arhitektura pogodbenega dela

```
pogodbe/
├── _tools/                     ← patch datoteke za lokalni sgl_docx_format.py
│   └── sgl_docx_format_PATCH_20260827.py
├── kdo_je_elena/               ← Elias Rudolf, "Kdo je Elena?"
│   ├── CHANGELOG.md            ← evolucija v1 → v5
│   ├── SESSION_NOTES.md        ← kompresiran povzetek prve seje (27. 8. 2026)
│   └── sgl_make_pogodba_elias_*.py   ← generatorji
└── (druge produkcije, npr. steklena_menazerija/, sunder_v_dvorani/ …)
```

## Ključne konvencije

- **Standard v7** (`ŠG_STANDARDI_pogodbe_v7_changelog.md` na Google Drive,
  6. 5. 2026 + popravki 20. 8. 2026). 9-poglavna struktura:
  I. UVODNE · II. PREDMET · III. **NARAVA RAZMERJA** (ne "HONORAR") ·
  IV. PRENOS MAT. PRAVIC · V. MORALNE · VI. FIZIČNI NOSILCI · VII. JAMSTVA
  · VIII. OBVEZNOSTI · IX. KONČNE.
- **Tipografija**: Calibri 11 (telo), Arial 12 bold (naslov pogodbe),
  robovi 2,5 cm, alineje z zamikom 1,25 / −0,50 cm + pomišljaj "–".
- **Podpisni blok** (§4.0, 20. 8.): 9 vrstic, vrstica 4 naziv pravne osebe
  (bold), **vrstica 5 IME+PRIIMEK BOLD**, vrstica 6 funkcija, vrstica 8 žig.
- **Metadata**: `set_metadata_milan(doc)` — author = "Milan Golob".
- **g./ga.** samo v uvodu strank (§3.1); ne v telesu, ne v podpisu.
- **Pogodbeni stranki** dvojina (2 stranki) · **Pogodbene stranke** množina
  (3+ stranki).

## Poti na Windowsu (Milanov disk)

- Windows Google Drive Stream: `C:\Users\ladmin\Moj disk\`
- Helper: `C:\Users\ladmin\Moj disk\KODE\_tools\sgl_docx_format.py`
- Pogodbe: `C:\Users\ladmin\Moj disk\Šentjakobsko gledališče\Pogodbe in računi\`
- Za vsako produkcijo je podmapa (npr. `Kdo je Elena\`, `Steklena menažerija\`,
  `Šunder v dvorani\`).

Generatorji imajo poti hardkodirane; te NE spreminjaj — Milan jih zaganja
lokalno na Windowsu.

## Pomembna referenčna dokumenta (Google Drive)

- `ŠG_STANDARDI_pogodbe_v7_changelog.md` (glavni standard)
- `LECNI_STANDARD_pregled_pogodb.md` (7 leč — vrstni red pregleda)
- `OBLIKOVALSKA_TAKSONOMIJA_pogodb.md` (predlog, čaka odločitev)
- `SPEC_widget_sestavljalnik_pogodb_v0.1.md` (osnutek widgeta)

## Delovni tok pri novi/spremenjeni pogodbi

1. Preveri, katera produkcija (imenska mapa).
2. Beri CHANGELOG.md — tam je vsa evolucija tega primera.
3. Aplikacija: 7 leč po vrstnem redu (pravna → davčna → operativna →
   precedenčna → lektorska → stilistična → oblikovalska). Ne skoči na
   oblikovanje pred kontrolo pravnih napak.
4. Vsak sklic na zakonski člen preveri VERBATIM (WebSearch) — moje
   predpostavke so bile v tej seji dvakrat narobe (16. → 21. čl. ZOdv,
   76.a → 94/44. čl. ZDDV-1).
5. Pri Nini/Iri/Eliasu je vzorec različen (glej sinoptično primerjavo v
   `pogodbe/kdo_je_elena/sgl_make_primerjava_ira_nina_elias.py`).

## Avtor sporočil in commit-ov

Ime: `Milan Golob`, email: `milan22golob@gmail.com`. Commit trailer:
```
Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>
Claude-Session: https://claude.ai/code/session_0111P5TPge36xxxCB9q9TGEt
```
