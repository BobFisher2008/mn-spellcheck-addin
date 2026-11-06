# mn_MN_2018 — dict-mn aligned to the 2018 orthographic standard

This is the dictionary the extension and add-ins ship by default. It is
**generated**; do not edit `mn_MN.dic`/`mn_MN.aff` by hand.

- **Base:** dict-mn 2026.09.21 by Batmunkh Dorjgotov (MPL-2.0), which follows
  Ts. Damdinsüren & B. Osor, «Монгол үсгийн дүрмийн толь» (1983).
- **Target standard:** «Монгол хэлний зөв бичих дүрмийн журамласан толь» (2018,
  ISBN 978-99978-3-172-9), issued by the Presidential National Council on Language Policy.
  Its foreword says it keeps Damdinsüren's rules and standardizes specific words
  (principles 1.2–1.16, 2.1–2.4); those are what change here.
- **Regenerate:** `python tools/eval/extract_pdf_corpus.py && npm run dict:2018`
  (needs the 2018 PDF in the project root).

## What changed (see CHANGES.tsv, one line per change)

| Change | Count | Evidence |
|---|---:|---|
| 2018 headwords added (missing from dict-mn) | 1 094 | headword list of the 2018 book; inflection flags copied from a dict-mn word with the same 2018 paradigm code and ending; 208 checked against every inflected form the book prints |
| Re-derived inflection (ethnonyms, rule 1.12) | 3 | `ойрадын`, `зүрчиний`, `төвөдийн` |
| Superseded root spellings forbidden | ~170 | (a) words the book itself marks wrong with **×** (e.g. `× товъёг`, `× эмхтгэл`, `× эрвээхий`); (b) pairs matching a numbered principle (1.2 билиг, 1.4 `-гана`/`-жигна`, 1.5 `-ээхэй`, 1.7 `будил-`, 1.8 `авлига`, 1.11 `ь`, 1.14 `эмхэтгэл`); (c) curated `rules-2018.tsv` |
| Derived words of those roots forbidden | ~3 420 | dict-mn lists derivatives as stems (`авилгажигдуул…`) |
| 2018 counterparts of those derivatives added | 3 339 | root swapped, same flags (`пижгэнэгд` → `пижигнэгд`) |
| Exact forms forbidden | 13 + 1 | forms marked × in the book (`гэрлүү`, `нэгбүр`, …) and `ойрдын` |
| Suggestion entries (REP) added to the .aff | 177 | superseded spelling → 2018 form, so the right word is suggested first |

Safeguards: a spelling is never forbidden if the 2018 book uses it anywhere
(headwords, definitions, appendices — checked through Hunspell stemming, so
`будуулах` protects `будуул`), and a ×-mark on a word that is itself a 2018
headword (a "do not confuse" note, e.g. `бэлгэ`) is ignored. 170 look-alike
pairs without such evidence are **not** applied; they are listed in
`REVIEW-2018.tsv` for a linguist (add `forbid` or `keep` lines to `rules-2018.tsv`).

## Results on the 2018 book (tools/eval/compare_dicts.ts)

| | dict-mn | this build |
|---|---:|---:|
| 2018 headwords accepted | 97.03 % | 100 % |
| Loanword appendix forms | 73.75 % | 96.32 % |
| Inflection-table forms | 92.37 % | 97.81 % |
| Flags on 130 k words of 2018-standard definitions | 1.21 % | 0.79 % |
| Spellings marked × in the book rejected | 53 / 93 | 77 / 93 (the other 16 are valid words in another sense) |

## Licence and provenance — read before redistributing

The files remain **MPL-2.0** (they are a modification of dict-mn; source form
is the file itself). The additions are derived from the 2018 dictionary, whose
redistribution terms are **not established**. Confirm with the publisher (the
Council: www.toli.gov.mn) before public distribution; internal use within an
organization is a lower-risk first step.
