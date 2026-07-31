# बसुका ग्राम वंशावली — Basuka Village Vanshaavali

A digital, searchable family tree for **बसुका (Basuka)** village, transcribed from the printed
**सकरवार वंशावली (Sakarwar Vanshaavali)**, pages 18–63.

**Live site:** https://ar2623441-oss.github.io/my-village-BASUKA-vanshaavali/

## Contents

- `index.html` — the whole website (single file: HTML + CSS + JS)
- `data/basuka_vanshaavali.json` — all 1,583 people, their parent/children links, generation, and branch
- `Basuka_Vanshaavali_MASTER.xlsx` — the same data as a spreadsheet, with a corrections log and open questions

## Editing

- To fix or add a person: edit `data/basuka_vanshaavali.json` (see the "How to Edit" tab on the
  site itself for the exact format).
- To preview locally: open this folder in VS Code, install the **Live Server** extension, and
  right-click `index.html` → **Open with Live Server**. (Double-clicking the file directly won't
  load the data, since browsers block local `fetch()` calls.)

## Publishing

Push to the `main` branch, then in the repo go to **Settings → Pages → Source: Deploy from a
branch → main → / (root)**. The site publishes automatically within a minute or two of every push.
