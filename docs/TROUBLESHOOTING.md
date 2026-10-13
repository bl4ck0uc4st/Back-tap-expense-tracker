# Troubleshooting

Every problem in this list was hit during the build. Most are Apps Script or Sheets behaviours that are not documented anywhere obvious.

---

## The request never reaches the script

### Some parameters silently vanish

**Symptom:** `amount` and `type` arrive, `category` onwards is missing. `?echo=1` shows the later keys absent entirely rather than empty.

**Cause:** Google reserves several single-letter query parameters for its own routing. Using `c` causes the request to be intercepted before your script runs. The failure comes from Google Drive, not Apps Script, so there is nothing in your execution log.

**Fix:** use two-letter aliases. This project uses `am`, `ty`, `ct`, `ds`, `dt`.

**How to isolate it:** add parameters one at a time in a browser and find the first that breaks.

```
?am=41
?am=41&ty=Expense
?am=41&ty=Expense&ct=Food      ← if this is the first failure, it is the name
```

### "Sorry, unable to open the file at this time"

A Drive error page rather than your script's output. Either a reserved parameter as above, or the deployment does not exist at that URL.

### The tail of the URL is dropped

**Symptom:** the URL looks complete in the Shortcuts editor, but the last parameters arrive empty. Replacing the variables with literal text changes nothing.

**Cause:** long URLs entered directly into **Get Contents of URL** can be truncated.

**Fix:** build the URL in a **Text** action and point Get Contents of URL at that variable.

### "The network connection was lost"

Usually not your network. Apps Script redirects every request through an internal Google domain, and `POST` with a JSON body does not always survive that redirect from Shortcuts.

**Fix:** use `GET` with query parameters.

---

## The script runs but nothing changes

### Edits to the script have no effect

Saving does not update the live URL. Apps Script keeps serving the last published snapshot.

**Fix:** Deploy → Manage deployments → ✏️ → Version: **New version** → Deploy.

### Multiple deployments

If two deployments exist, each has its own URL and its own version. You can update one while your Shortcut points at the other.

**Fix:** Manage deployments, confirm which URL your Shortcut uses, and update that one. Archive the rest.

**How to tell:** the version marker in every response. If the URL returns an old version string, you are updating the wrong deployment.

---

## Charts

### A chart renders with its title but no data

Three possible causes, in order of likelihood.

**1. The source range is hidden.** Sheets will not plot data from hidden rows or columns. Helper blocks therefore live in far-right columns that are visible but out of the way, not hidden ones.

**2. The range was written in the same execution.** Values may not be committed when the chart is built.

```javascript
sheet.getRange(2, COL, rows.length, 2).setValues(rows);
SpreadsheetApp.flush();          // commit before the chart reads it
sheet.insertChart(chart);
```

**3. `chartArea` is over-constrained.** A `chartArea` with explicit offsets plus a right-hand legend can compute to an unrenderable plot area, giving you a titled but empty card. Removing the option lets Sheets size it.

### Two charts overlap

Charts are positioned by row and column but sized in pixels, so their footprint depends on row heights. A 330px chart starting at row 12 covers roughly 11 rows at 30px each.

**Fix:** either reduce the chart height or push the next section's start row down.

---

## Data problems

### `#REF!` in the summary cells

The summary block lives in columns G and H on the same rows as the entries. Deleting entry rows by hand removes part of it.

**Fix:** use **💰 Tracker → Delete selected entries**, which clears only columns A–E and closes the gap. If it has already happened, run a refresh — the summary is rebuilt from scratch each time.

### New entries land several rows below the data

`appendRow()` and `getLastRow()` consider the whole sheet, including helper columns. With a short entry list and a longer helper block, new rows land past the gap.

**Fix:** find the target row from column A alone.

```javascript
function lastDataRow(sheet) {
  const v = sheet.getRange("A:A").getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]).trim() !== "") return i + 1;
  }
  return 1;
}
```

Clearing a row's contents also leaves its formatting behind, which counts as content. Delete rows properly rather than clearing them.

### Dates sort in the wrong order

Storing dates as text sorts them alphabetically: `3 Sep` lands before `14 Aug`.

**Fix:** write a real `Date` value and apply a display format with `setNumberFormat("d mmm yyyy")`.

### The month picker throws an error mid-rebuild

**Symptom:** `Exception: Pick a month to explore`. The dashboard renders down to the picker and stops.

**Cause:** `sheet.clear()` does not remove data validation. The old rule lists only the months that existed previously, so setting a newly added month fails validation, and the exception aborts everything after it.

**Fix:** clear validations first, allow invalid values, and guard the call.

```javascript
picker.clearDataValidations();
picker.setValue(pick);
try {
  picker.setDataValidation(rule);
} catch (e) {
  picker.clearDataValidations();
}
```

### "August 2026" turns into a date

Sheets parses month-year strings as dates.

**Fix:** `range.setNumberFormat("@")` before setting the value.

---

## Shortcuts

### Set Variable will not accept typed text

Its input field only takes variables. Put the literal in a **Text** action first and point Set Variable at that.

### Menu Result picks the wrong menu

With several menus, a generic `Menu Result` pill binds to the most recent one at that point in the flow, which is not always what you meant.

**Fix:** inside each branch, use a Text action with the value and a Set Variable that names it. Explicit, and it survives rearranging.

### Multi-word values arrive truncated

Anything after a space is dropped, along with any parameters following it.

**Fix:** URL Encode free text such as the description. For controlled values such as categories, keep them single-word or hyphenated.

### The variable picker looks empty

It scrolls horizontally. Swipe left along the toolbar above the keyboard to see the rest.

---

## Diagnosing anything else

**Version markers.** Every response includes a version string. If it disagrees with the file you just pasted, the deployment is stale.

**`?echo=1`.** Reports exactly what arrived without writing. The difference between a missing key and an empty one tells you whether the parameter was dropped in transit or sent empty.

**Quick Look.** Add a Quick Look action after Get Contents of URL to see the raw response on the phone. Point it at your Text action instead to see the URL actually being sent.

**Bisect.** Add parameters or actions one at a time until it breaks. Nearly every bug above was found this way rather than by reading logs.
