# TNC Filtration Service Estimator

A farmer-forward companion to **AquaPIP**. A grower picks a bivalve species, enters
water temperature and one or more size classes (shell height and/or dry weight,
plus how many animals), and gets the estimated **maximum volume of water their
stock can clear** — tied to MEL **Objective 1.3** (farmed biomass improves light
penetration / reduces hypoxia).

It lives in the **same GitHub repo as AquaPIP** and is deployed as its **own**
Streamlit Community Cloud app (own URL), so it can be iterated without touching or
rebooting AquaPIP.

---

## Files

```
filtration_app.py                 # Streamlit entrypoint (deploy THIS file)
filtration_logic.py               # loading + equation translation + safe eval (no Streamlit)
filtration_report.py              # references export (CSV + styled PDF)
data/
  Clearance_rate_estimation_tool_TNC.xlsx   # THE workbook — everything is read from here
  protocols/                      # drop dry-weight protocols here (see its README)
```

No new dependencies beyond what AquaPIP already uses: `streamlit`, `openpyxl`,
`reportlab`. (`pandas` ships with Streamlit.)

---

## Deploy (Streamlit Community Cloud)

1. Commit these files to the AquaPIP repo.
2. On share.streamlit.io → **New app** → same repo/branch → **Main file path** =
   `filtration_app.py`.
3. Deploy. (Optional) put the AquaPIP URL into `AQUAPIP_URL` near the top of
   `filtration_app.py` to show a "Back to AquaPIP" button, and add the new app's
   URL into AquaPIP the same way to link the two tools together.

---

## How the tool stays live as the science improves

**Everything** the tool shows — species list, clearance-rate (CR) equations,
references, length-weight conversions, the *How to use* text and the caveats — is
read **live** from the single workbook `data/Clearance_rate_estimation_tool_TNC.xlsx`
by **sheet + header name** (never by cell position). Philine keeps editing that
file exactly the way she does today. To publish a change: overwrite the file in
the repo (GitHub → the file → pencil/edit → *Upload*/replace, or commit the new
version) and the app reloads it on its next reboot.

### Adding / changing a species equation

In the **`Final_equations`** sheet, the CR column holds a normal Excel formula that
points at the calculator's input cells, e.g.

```
=16.34*EXP(-0.0152*(Filtration_calculator!E2-28.6)^2)*Filtration_calculator!D2^0.608
```

The tool auto-translates that into a safe, evaluable expression using named
variables:

| Calculator cell | Meaning | Variable |
|---|---|---|
| `Filtration_calculator!E2` | water temperature (°C) | `T` |
| `Filtration_calculator!C2` | mean shell height (mm) | `L` |
| `Filtration_calculator!D2` | dry tissue weight (g) | `W` |

`EXP()` → `exp()`, `^` → `**`, absolute `$` refs are stripped. Evaluation uses a
**whitelisted parser** (only `+ - * / ** %`, `exp/log/log10/sqrt/abs/pow/min/max`,
and the variables above) — no Python builtins, no arbitrary code.

- A row whose CR cell is **not a formula** (e.g. `No CR formula available`,
  `CR may become available`) is shown as a species that exists but is **greyed /
  not yet calculable**, using the sheet's own wording.
- Which inputs a species needs is detected from its equation: the one species
  whose equation uses shell height directly (*Mytilus edulis*) automatically shows
  **no dry-weight input**; all dry-weight-based species show the dry-weight options.

**Optional clean override:** if you ever add a column named exactly **`Equation (py)`**
to `Final_equations`, its text is used verbatim (same variables `T`/`L`/`W`) in
preference to translating the Excel formula — handy for anything awkward to express
as a spreadsheet formula.

### Dry-weight (DTW) options shown to the farmer

For dry-weight-based species the tool offers three sources, in accuracy order:

1. **Use one dry weight for all** — a single DTW value (default `1 g`, editable) for
   a single group. This is the placeholder/fallback.
2. **I have my own dry weights by size class (more accurate)** — a row per size
   class with its **measured** dry weight and count; rows are summed.
3. **Estimate dry weight from my shell height** — a row per size class by shell
   height; DTW is derived from a **length-weight conversion** and shown in the
   results. Enabled only for species that have a conversion in the library
   (currently *Crassostrea rhizophorae*); otherwise it's shown but disabled with a
   note.

*Mytilus edulis* (length-based) skips DTW entirely and takes shell height + count
per size class.

### Length-weight conversions

Add rows to the **`length-weight_conversions`** sheet (species, location, season,
the `equation` text, the live conversion formula, and a reference). Each becomes a
selectable conversion in the *Estimate from shell height* option and appears in the
**Length-weight conversions** library table. If several conversions exist for one
species, the farmer picks which to use.

### References

Add rows to the **`References`** sheet (short citation, full reference, link — a
plain URL, a `=HYPERLINK(...)` formula, or a bare DOI all work). The per-species
**"Reference & notes"** button matches a species' CR reference string (which may
list several sources joined by "and") to the full references, and the
**"Download all references"** buttons export the whole list as PDF and CSV.

### How-to & caveats text

Edit the **`How to use`** sheet. Text before *Instructions* becomes the intro; the
steps under *Instructions* go into the "How to use this tool" expander; text under
*Important caveats* becomes the always-visible amber caveats box.

---

## A note on temperature range

Several equations are fitted around a species' measured temperature range and can
return a negative rate far outside it. The tool **caps negative estimates at 0** and
flags that temperature was outside the fitted range, so a farmer never sees a
nonsensical negative number.
