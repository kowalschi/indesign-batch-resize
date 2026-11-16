# InDesign Batch Resize

An ExtendScript (JavaScript) tool for Adobe InDesign that turns one key visual into any number of
formats, driven by an Excel file.

**Case study with diagrams: [resize.magic.re](https://resize.magic.re/)**

I wrote the first version in **2020** at Hogarth, on my own initiative — nobody had asked for it, and
nobody on the team had automated this work before. My colleagues adopted it and **the team used it in
production for over four years**. I have kept developing it ever since; this repository keeps every
version, from `v01` to `v23`.

In advertising, a campaign's key visual has to be delivered in dozens or hundreds of sizes (posters,
POS materials, press ads, banners). Doing this by hand takes from 30 minutes to several hours per
format. With the script, a format needs about **5 minutes of human attention**, mostly for a final check.

## How it works

1. A **master InDesign document** holds the key visual laid out on several pages with different
   aspect ratios (portrait, square, landscape, extreme formats).
2. An **Excel file** lists the formats: name, visible size, total size, units, bleed, which master to use.
3. For each row the script picks the master page with the closest aspect ratio, creates a new
   document at the requested size and repositions every element according to its **script label**.
4. The result is saved as a new `.indd` in a `_result` folder, named from the Excel data.

### Label-based rules

Each object in the master carries a label that tells the script how it behaves when the page changes size:

- background → always covers the page, including bleed
- logo → fixed corner, size and safe space computed from the format (e.g. a percentage of the diagonal)
- legal / health-warning band → bottom of the page, fixed percentage of the height
- badges, headlines → anchored to a corner, an edge or the centre of the visible area
- "expand to bleed" on any combination of sides

`Add_Labels_for_resizing` is a companion palette that sets these labels with buttons laid out like
an InDesign page, so designers do not have to type them.

## History

| Year | Versions | What changed |
|---|---|---|
| 2020 | v01 – v09b | First working version: reads a tab-delimited text file, picks the master page by ratio, resizes and repositions. Written by hand, learning from the InDesign scripting reference and forums. |
| 2023 | v10a – v11 | Brand-specific rules for legal bands and logos. |
| 2024 | v11 – v12 multi-masters | Several master documents in one run, reads Excel directly, runs on Windows and macOS. |
| 2025 | v14 – v15a | Units (mm, cm, px, inch), composed file names, error report. |
| 2026 | v16 – v23 | Settings dialog that remembers naming options, configurable Excel columns, new label types, labels palette. From v17 on I refactor and extend the code with the help of an AI assistant; the logic and the rules come from six years of using it in production. |

The early versions are kept on purpose: they show how the code grew from one long procedural file
into functions with a configuration block, a dialog and error handling.

## Files

- `2020/` … `2026/` — one folder per year, one file per version
- `2026/Redimensionari_multiMasters_v23.jsx` — current version
- `2026/Add_Labels_for_resizing v4.jsx` — labels palette
- `2026/Redimensionari_template.xlsx` — Excel template with the expected header

## Usage

1. Copy the current script and the labels palette into the InDesign Scripts Panel folder.
2. Prepare the master document(s) and label the objects.
3. Fill in the Excel template and save it next to the masters.
4. Open a master and run the script.

## License

Published for portfolio and demonstration purposes. Please ask before reusing the code.

© Dan Ichimescu
