# Third-party components

The project's own code is MIT-licensed (see LICENSE). The components below keep
their own licences; MIT is compatible with all of them (MPL is file-level
copyleft: it applies to the dictionary and Hunspell files only, not to our code).

| Component | Version | License | Where it is used | Source |
|-----------|---------|---------|------------------|--------|
| dict-mn (mn_MN.aff, mn_MN.dic) | 2026.09.21 | MPL-2.0 (versions up to 2026.09.07 were LPPL 1.3c) | Dictionary data, shipped unmodified in `dict/` | https://github.com/bataak/dict-mn |
| Hunspell | 1.7.3 | Tri-license MPL 1.1 / GPL 2.0 / LGPL 2.1 (used under MPL 1.1) | Spellchecking engine, compiled to WebAssembly, unmodified | https://github.com/hunspell/hunspell |
| @farscrl/hunspell-wasm | 1.0.1 | MIT | JS/WASM bindings for Hunspell (embedded WASM) | https://github.com/farscrl/hunspell-wasm |

Notes

- MPL-2.0 (dict-mn) and MPL-1.1 (Hunspell) are file-level copyleft. Both are
  distributed unmodified; their full source is available at the URLs above.
  If you modify `mn_MN.aff`/`mn_MN.dic`, the modified files must stay under
  MPL-2.0 and be made available in source form (they already are source form).
- Project additions to the dictionary live in a separate file
  (`dictionaries/mn_MN/supplement.dic`, MPL-2.0) and are merged at runtime.
- The 2018 dictionary PDF «Монгол хэлний зөв бичих дүрмийн журамласан толь»
  (ISBN 978-99978-3-172-9) is **not** redistributed. However, the shipped
  dictionary `dictionaries/mn_MN_2018` contains words and spelling decisions
  **derived from it** (≈1 100 headwords, their inflections, and ≈200 superseded
  spellings it marks). Its redistribution terms are not established: confirm with
  the publisher (Хэлний бодлогын үндэсний зөвлөл, www.toli.gov.mn) before public
  distribution. `DICT=mn_MN npm run build` produces a build without this data.
