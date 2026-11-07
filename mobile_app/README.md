# H2R Nigeria – mobile app

An installable, offline-capable mobile web app (PWA) for browsing the
Hard to Reach (H2R) assessment results produced by the R scripts in this repo.

## What it does

| Tab | Content |
| --- | --- |
| **Overview** | Settlements assessed per LGA, sector shortcuts, latest key findings |
| **Indicators** | Pick an LGA and a sector to see the share of settlements reporting each response. "Top 3" questions highlight the three most reported answers. |
| **Compare** | Heat table of every indicator across all LGAs |
| **Findings** | The written key findings, searchable and filterable by sector |
| **About** | Methodology, dataset details and install instructions |

There are no dependencies or build step: just plain HTML, CSS and JS.
After the first visit, a service worker caches the app and its data so it
works without a connection. That matters for field teams with patchy coverage.

## Run it locally

```bash
cd mobile_app
python3 -m http.server 8000
# open http://localhost:8000 (on a phone: http://<your-computer-ip>:8000)
```

## Put it on phones

Host the `mobile_app/` folder on any static HTTPS host, for example GitHub
Pages, Netlify or an internal server. Service workers need HTTPS (or
`localhost`). Then, on the phone:

- **Android (Chrome):** menu ⋮ → *Install app* / *Add to Home screen*
- **iPhone (Safari):** Share → *Add to Home Screen*

## Update the data for a new round

The app reads `data/h2r_data.json`. To regenerate it from the R pipeline's
outputs:

```bash
pip install openpyxl
python3 mobile_app/scripts/build_data.py \
  --summaries input/Statement_Summaries_All_Clusters.xlsx \
  --names input/INDICATORS_TO_RANK.xlsx \
  --round "May 2019"
```

The script expects these sheets in the summaries workbook:

- `for_factsheet`: question, indicator name, type (`rank`/`direct`), sector, then one column per LGA
- `statements`: key findings (row 1 is the methodology note)
- `H2R_LGA_Results_*`: used to count settlements per LGA

Readable option labels come from the `name_change` sheet of
`INDICATORS_TO_RANK.xlsx`. When an indicator is missing there, the script
derives a label from the indicator name.

After you change any app file (not the data), bump `VERSION` in `sw.js` so
installed copies pick up the update.

### Known data quirks

The script works around a few issues in the `for_factsheet` sheet: duplicated
rows, a `#N/A` question text, livelihood indicators tagged as "shelter", and
sub-questions (non-displaced / returnees / IDPs) listed under one question
text. Some mortality rows at the end of the health section are still
interleaved in the source sheet. Fixing them in the Excel file and re-running
the script will clean them up.
