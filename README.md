# JAG Attendance Report Generator

A static, browser-only attendance report generator for **Junior Adventures Group**.

It accepts a MagicBooking / Daily Attendance Excel export, detects the reporting structure, and generates:

- a detailed `.xlsx` workbook in the established weekly attendance format;
- a JAG-branded `.pdf` summary with a detailed weekly appendix;
- an optional `.zip` containing both files.

## What is detected automatically

The generator does not assume a fixed date range or a fixed set of JAG products. It detects:

- school / centre name;
- minimum and maximum attendance date from the actual rows;
- programme families such as `Rise then Shine` and `Stay and Play`;
- individual bookable session variants;
- operating days;
- attendance entries per date and session;
- programme subtotals where a programme contains multiple bookable session variants.

The school name can be corrected in the interface before generating the report.

## Privacy

The app is intentionally static. The spreadsheet is parsed in the user's browser and is **not uploaded to a server**.

The generator loads its JavaScript libraries and Montserrat web font from public CDNs, so an internet connection is required when the page first loads. Attendance data itself remains in the browser.

## Deploy on GitHub Pages

1. Create a new GitHub repository, for example `jag-attendance-generator`.
2. Upload everything in this folder to the root of the repository.
3. Open **Settings → Pages**.
4. Under **Build and deployment**, choose **Deploy from a branch**.
5. Select your main branch and `/ (root)`, then save.
6. GitHub will provide the Pages URL after deployment completes.

No build step, npm install, database, backend or server function is required.

## Files

- `index.html` — application layout and CDN dependencies
- `styles.css` — JAG UI plus PDF page styling
- `app.js` — Excel parsing, report logic, XLSX generation and PDF generation
- `assets/jag-logo.png` — supplied JAG logo asset
- `.nojekyll` — tells GitHub Pages to serve the files directly

## Expected export columns

The parser searches for a header row rather than assuming a fixed row number. It expects equivalent columns for:

- session (`Sessions` / `Session`)
- booked or attendance date (`BookedDate`, `Booking Date`, `DateBooked`, etc.)
- centre / school / site (`Centre`, `Center`, `School`, `Site`)

`Label` and `BookingId` are used when present but are not mandatory.

## Counting logic

Each unique booking ID is counted once per date and session. If no booking ID is present, the source row is used as the attendance entry.

`Day total` is the sum of all detected session attendance entries on that date. A programme subtotal such as `Stay and Play Only` is created when a programme family contains multiple bookable sessions.

`NA` is shown when a programme family has no attendance entries at all on that day; a zero within an otherwise active programme is shown as `0`.

## Libraries

Loaded from jsDelivr in `index.html`:

- xlsx-js-style
- html2canvas
- jsPDF
- JSZip

The PDF is rendered from branded HTML pages into jsPDF. This allows the generated report to retain the Montserrat/JAG visual style without bundling font files in the repository.
