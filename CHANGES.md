# "How close are you?" readiness page: repo update (Oct 2026)

Copy these into the aquapip-survey repo at the same paths. No new dependencies.

## New
- `readiness.py`: the Streamlit page. It sits between "What are you already monitoring?" and the MEL Framework overview.
- `readiness_logic.py`: scoring, suggestions and readout rows (no Streamlit).
- `indicator_art.py`: loads `assets/indicators/*.svg` for the page and draws them in PDFs without extra packages.
- `assets/indicators/`: 33 pictures, one per sector x indicator. Each file name matches the 'Picture' column of the Readiness sheet.

## Changed
- `app.py`:
  - new `readiness` view and route; the monitoring page's Next button goes there;
  - Start over clears the page's `rd_*` answers;
  - "Edit answers" straight back to results drops the readout, so it can't go stale.
- `report.py`:
  - when the farmer has visited the page, the readout sits at the top of page 1 of the summary PDF;
  - new `build_readout_pdf()` (two pages) feeds the page's download button.
- `data/MEL_Indicator_Reference.xlsx`:
  - new 'Readiness' sheet;
  - 'Farm records' rows added to Monitoring_Approaches;
  - README notes for both.

## Editing
- Wording, parts, questions, credits and weights live in the 'Readiness' sheet. It's read by column header, and changes show on the next app load.
- To change a picture, replace its SVG and keep the same file name.
- The page uses `st.pills` on Streamlit 1.40 and later, and falls back to a select box and radio buttons on older versions.

## To check
- 'Sea-lice and disease checks on stock' (Farm records) is a proposed screener item for the finfish net-pen indicator.
- The workbook was built from the copy in the Claude project. If the repo copy has changed since, carry the new sheet and rows across.
- If "Edit answers" jumps straight back to results, this page's adds and removes aren't re-applied until the farmer passes through the page again.
