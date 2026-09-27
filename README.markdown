# export-mbox-to-pdf

This program turns one or more mbox mailboxes into a single print-oriented master PDF. The layout, image geometry, and quote bars follow a long-used Perl converter. The Python program is the implementation you run. A small bash wrapper (`export-mbox-to-pdf`) is the usual entry point: it activates the project virtualenv, passes paths, and forwards flags.

The output is meant as an archival paper image of the mailbox, not as a substitute for the mbox. Bodies, headers, images, PDFs, and convertible office files are drawn or embedded on DIN-A pages (default A0). Calendars, vCards, and zip archives are summarised in the body. A companion folder tree can hold the raw extracted files. Deduplication can write a sibling mailbox without strong duplicates so later runs do not walk the same payload twice.

The remainder of this document is split. First comes the end-user interface: configuration file, command-line switches, wrapper rules, and what a run does. After that comes implementation: parsers, fingerprints, ImageMagick montages, Ghostscript fit, qpdf rewrite, and the limits of large PDFs.

## What a successful run produces

The converter writes the PDF named in the configuration under `output_path`. If that file already exists it is renamed with a suffix taken from its modification time (`stem-DD-MM-YYYY-HH-MM-SS.pdf`) so a new assemble does not overwrite the previous master. With `--finalize`, a second file `stem-master.pdf` is written after qpdf has rebuilt the cross-reference as PDF 1.7 object streams; the unfinalized assemble is rotated the same way.

Two log files live in `output_path`. `convert.log` receives the full run. `errors.log` receives errors and warnings that need a human look (failed converts, empty mails, strong duplicates when review is requested, unreadable images). Existing logs are rotated by mtime before the new files are opened.

If `archive_attachments` is true, the program creates `output_path/attachments/`. A folder already present under that name is renamed as a whole (`attachments-DD-MM-YYYY-HH-MM-SS`) so each archive-enabled run starts empty. Inside the new tree, each processed mail gets `Email #<listing-number> (<when>)`, where `<when>` is built from the `Date` header using day.month.year and, when present, hours and minutes with a Unicode ratio colon (`∶`). Seconds are written only when the header contains them. A header that is only a year becomes `(2016)`. A missing or unparseable date becomes `(unknown date)`. Copies of the original plain and HTML parts are stored as `01_original.txt` and `02_original.html`. Attachments are numbered copies of the extracted files. Modification times on those archive copies prefer EXIF or PDF metadata, then the email `Date`, then the copy clock. Temporary files under `/tmp` are not stamped.

If `use_deduplicated_mbox` is true, the program may also write, next to each source mbox, a file whose name is the source stem plus `-deduplicated.mbox`, and a sibling `-skipped.mbox` holding messages dropped as strong content duplicates while building that cache. Those cache files are inputs for later runs when they are newer than the source.

`--list-emails` writes nothing except tab-separated rows on stdout.

## End-user API

### How you invoke the program

The Python file is not meant to be called with only a mailbox path. It always needs four path-like arguments: the project base directory (fonts live under `Build/assets/fonts` relative to that base), the working directory used to resolve a relative configuration path, the configuration file itself, and the LibreOffice binary used for office conversion. The bash wrapper fills those in from the project layout and from two variables at the top of the wrapper: `working_directory` and `configuration_file`.

The wrapper refuses to start a conversion unless you pass either `--test` or `--finalize`. Those two are exclusive. `--test` is a wrapper-only name: it is not sent to Python. It means “produce the configured master PDF without the finalize qpdf rewrite.” `--finalize` is sent through and turns on PNG compression plus the qpdf object-stream rewrite. `--fast` of any multiplicity cannot be combined with `--finalize`. Listing (`--list-emails` or `-l`) does not require `--test` or `--finalize`.

If the wrapper is started with no arguments it prints its help text and exits. The Python program does the same when it sees an empty argument vector.

### Configuration file

The configuration is one JSON object. Paths in `input_files` and `output_path` are used as written; make them absolute or ensure the process current directory is the one you intend. A minimal file looks like this.

```json
{
  "input_files": [
    "Emails/Ju.mbox"
  ],
  "output_path": "Ausfertigung",
  "output_filename": "Briefwechsel (high-resolution Master).pdf",
  "dimensions": "A0",
  "image_dpi": 300,
  "prefer_plain_text": true,
  "use_montages_for_images": true,
  "archive_attachments": true,
  "use_deduplicated_mbox": true,
  "duplicate_emails": {
    "enabled": true,
    "action": "warn"
  },
  "email_to_name": {
    "Ju@example.org": "Ju"
  }
}
```

`input_files` is either a string or an array of strings. Each entry is a path to an mbox. The program refuses a path whose file name already ends with `-deduplicated.mbox` or `-skipped.mbox`. Several sources are concatenated, then sorted by the `Date` header so listing numbers are chronological across all inputs. Missing files abort the run.

`output_path` is the directory for the PDF, the logs, and the optional `attachments` tree. It is created if needed.

`output_filename` is the assemble name only, not a path. The finalized name is derived by inserting `-master` before the suffix.

`dimensions` selects a ReportLab DIN-A page. Allowed values are `A0` through `A10`. The page is a real size in PostScript points. Geometry that the Perl program expressed in millimetres on a triple-width A4-like sheet is scaled by the ratio of this page width to that reference width (`210 mm` at `300` dots per inch). Changing `dimensions` scales fonts, margins, quote bars, and montage cells together. It does not by itself change how many pixels are embedded in a photograph; that is `image_dpi`.

`image_dpi` is the raster cap used when ImageMagick builds montage PNGs: pixels per inch of the size those images will occupy on the master sheet. The Perl program used a document density near `300` on its reference width. Keep `300` for a print master on A0 if you want the same pixel budget as that reference. `document_dpi` is accepted as a legacy synonym if `image_dpi` is absent. `--fast` once overrides this value to `10` for the run and does not write it back to the JSON file.

`prefer_plain_text` chooses which body to typeset when both `text/plain` and `text/html` exist. `true` (the intended archival mode) uses plain text and only falls back to a stripped HTML rendering when plain is empty. `false` prefers HTML. In both cases HTML is converted to plain-like text with line breaks from `<br>` and `<p>`, quotes from `<blockquote>`, and tags removed. The ReportLab HTML parser is not used.

`use_montages_for_images` controls how many photographs share a page. `true` packs up to nine images in the same grid the Perl program used (one, then two in a column, then two-by-two, three-by-two, three-by-three). `false` puts each image on its own page. A single image still goes through the montage code path so sizing matches a one-cell grid.

`archive_attachments` enables the dated `Email #N (…)` tree described above. It does not change the PDF. Turning it on rotates any existing `attachments` directory at the start of the conversion.

`use_deduplicated_mbox` enables the cache next to each source. When the cache exists and its mtime is at least the source mtime, conversion and `--list-emails` read the cache instead of the source. When the cache is missing or older, a conversion run rebuilds it from scratch by walking every source message, hashing body plus saved attachments, writing keepers to `-deduplicated.mbox` and strong duplicates to `-skipped.mbox`. Empty messages (no body and no saved attachments) are keepers; they are never collapsed by this cache. `--select-year` and `--select-emails` do not limit the rebuild: the cache always reflects the whole source. `--list-emails` will use an existing cache but will not build one.

`duplicate_emails` is a second, in-run filter after year and index selection. If the value is a boolean, `true` means enabled with action `warn`. If it is an object, `enabled` turns the pass on and `action` is either `warn` or `skip`. `warn` keeps every selected mail and prints a warning when a non-empty mail has the same body hash and attachment-hash pair as an earlier kept mail. `skip` omits those later copies from the PDF. The same pass always warns when subject and body are both empty, and when body and attachments are both empty, and it never auto-skips those empty cases. A “likely duplicate” (same body, different attachments, or the reverse) is logged at info level so `--quiet` hides it.

`email_to_name` maps a lower-cased address to the display name drawn in the From/To/Cc block. If the address is in the map, that string is used. Otherwise the name parsed from the header is used unless it is identical to the address, in which case the name column is left empty. `--debug` fills missing names with an obvious placeholder string so you can see unmapped addresses.

Unknown JSON keys are ignored. A missing required key (`output_path`, `output_filename`, or a usable `input_files` list) aborts.

### Program switches

`--base-directory` is the project root the wrapper already knows. Fonts are loaded from `Build/assets/fonts/OpenSans-Regular.ttf`, `OpenSans-Bold.ttf`, and `OpenSans-LightItalic.ttf` under that root. If registration fails the PDF falls back to Helvetica.

`--working-directory` is prepended to a relative `--configuration-file`. An absolute configuration path is used as-is.

`--configuration-file` is the JSON file above.

`--document-converter-binary` is the LibreOffice (or compatible) executable. The program calls it headless with `--convert-to pdf`. Only the formats classified as office are sent there. If your installation has Writer but not Calc, spreadsheets will log a load failure and remain hashed rather than embedded.

`--select-year` (wrapper aliases `--year` and `--select-year=N`) keeps messages whose `Date` parses to that calendar year. It applies to listing and to conversion. It does not restrict cache rebuild. Value `-1` or omission means all years. A message with an unparseable date is dropped when a year filter is active.

`--select-emails` (wrapper aliases `--select`, `--select-emails=…`, and a following token without `=`) selects by the 1-based index after chronological sort of the loaded mailboxes. The current parser accepts a single integer `N` (that one mail) or a single closed range `A-B`. The string `0` means all. A token that does not parse is treated as all, which is easy to misuse; prefer an explicit `0` or omit the switch. Union lists, open ranges, and `!` exclusions have been discussed and are not implemented. These indices are the first column of `--list-emails`. Filtering happens after sort and after choosing the deduplicated cache if that cache is in use, so number `772` is “the 772nd remaining message in the mailbox you actually read,” not “the 772nd message in the raw source” when the cache has dropped copies.

`--list-emails` (wrapper `-l`, `--list`, optional `=TEXT` or a following word that does not start with a dash) prints one line per selected mail and exits. Columns are listing number, date as `DD-MM-YYYY-HH-MM-SS`, `From → To`, subject, and attachment file names, separated by tabs. A following string is a literal substring filter on that whole line, case-insensitive, not a regular expression. Attachment names include real filenames and, for `message/rfc822` parts without a name, the token `attached-email`. The wrapper adds `--quiet` so status lines do not mix with the table. Listing does not render a PDF and does not invoke LibreOffice.

`--verbose` prints extra progress: tools found, temp directory, each saved part, montage paths, Ghostscript fit lines. It does not imply developer traces.

`--debug` implies `--verbose` and adds command lines, fingerprints, image pixel sizes and DPI, and mailbox family counts (source, deduplicated, skipped). It also substitutes the debug display-name placeholders. Combined with `--quiet`, info lines are still suppressed; debug and verbose lines still print.

`--quiet` hides info-level lines (including “likely duplicate” and “type guessed from extension”). Warnings, errors, verbose (if requested), and debug (if requested) still appear. Listing mode forces quiet in the wrapper so the TSV stays clean.

`--fast` may be repeated. Written `-f`, `-ff`, `-fff`, or several `--fast` tokens, the count is what matters. One occurrence sets `image_dpi` to `10` for that run, keeps page geometry, and writes uncompressed PNG montages. Two occurrences skip ImageMagick raster and draw one placeholder page per image batch with the label `Image`, keeping the same page count as a full image section. Three occurrences also skip LibreOffice and Ghostscript embedding and draw one placeholder page per PDF or office attachment, labelled `PDF`, `Document`, and so on. Timeouts stay at twenty-four hours at every fast level. `--finalize` cannot be combined with any fast level. The configuration file is not modified.

`--finalize` is the production rewrite. Montages are written with PNG quality 100. After the pypdf assemble, qpdf is run with `--object-streams=generate`, `--min-version=1.7`, `--compress-streams=y`, and `--recompress-flate`. `--linearize` is not used. Linearization hint tables store some offsets in 32 bits and cannot describe a file of many gigabytes; they are for small documents meant to start painting during an HTTP download. The rewrite writes to the run’s temp directory, then copies to `stem-master.pdf` in `output_path` after rotating any existing master and rotating the unfinalized assemble by its mtime. Every conversion that has qpdf available, including test runs, still runs `qpdf --check --warning-exit-0` on the file that remains in the destination. That check does not overwrite the file.

`-h` and `--help` print help and exit.

## External programs the user must have

ImageMagick (`magick` or `convert`) builds every photograph cell and montage. Ghostscript (`gs`) fits attached PDFs onto the master page size at compatibility 1.7. qpdf checks the result and, with `--finalize`, rewrites it. LibreOffice converts office families when fast level is below three. The Unix `file` program, if present, contributes MIME types when magic bytes are ambiguous. OpenSans files are optional but expected in the project tree.

A run that cannot find ImageMagick, Ghostscript, or qpdf exits before conversion. Listing mode does not require those tools.

## Implementation

### Reading mailboxes

Messages are read with `RobustMbox`, a separate module that splits on `From ` lines in binary and feeds each block to `email.parser.BytesParser`. That avoids `mailbox.mbox` treating the file as text in a locale encoding and dropping or corrupting 8-bit payloads. Messages that fail to parse are still yielded with an `X-Robust-Mbox-Error` header and a warning. `write_mbox` writes the cache files in the same `From `-separated form.

### MIME walk

`iter_logical_parts` flattens multipart trees but yields a `message/rfc822` part as a unit so it can be treated as an attached mail. `text/plain` and `text/html` are decoded with the part charset and accumulated. HTML is reduced to text as described above, including `__BLOCKQUOTE_START__` and `__BLOCKQUOTE_END__` markers that the canvas turns into quote depth.

A part whose payload is already a nested `Message`, or whose declared type is `message/rfc822` / `message/news`, or whose decoded bytes look like headers (`From `, `Received:`, `Subject:`), is walked as another mail and merged into the same body and attachment buckets. That is why a forward named `attachment.bin` is no longer an empty payload: `get_payload(decode=True)` is `None` for a nested message object, which used to abort the save and leave the parent looking empty.

Other parts are written under the temp directory, named `{listing}_{sha12}_{stem}.{ext}`, then classified by magic (PDF header, JPEG/PNG/GIF, OLE, ZIP/OOXML/ODF, `BEGIN:VCARD`, `BEGIN:VCALENDAR`, `From `) with `file(1)` and extension as fallbacks. Role mismatches between magic and declared MIME are warnings for review. Images go to the montage list. PDFs and successful LibreOffice outputs go to the embed list. ICS and vCard text become German addenda in the body. Zip files contribute name, size, and SHA-256 only. Video is hashed and archived, not embedded.

### Deduplication fingerprints

A strong fingerprint is the SHA-256 of a normalised body (CRLF folded, trailing spaces stripped, runs of blank lines collapsed) plus the SHA-256 of a sorted list of `original_name:size:sha256` for every saved attachment. Cache rebuild and the optional in-run filter use that pair. Walking a nested rfc822 changes the parent fingerprint because the inner body and files now exist. A forward that is only a wrapper around a mail already in the box can therefore leave the cache and appear in `-skipped.mbox`. That shrinks `-deduplicated.mbox` compared with an older cache that treated the wrapper as empty.

### Page drawing

ReportLab canvas draws a blue header band (listing number, date, page, From/To/Cc with optional wrapping of long display names), a white subject band that grows with wrapped subject lines, a dark rule, then the body. Quote depth is a pair of markers in the text. The canvas keeps a start `y` per depth. On each non-empty source line it draws a short tick beside that line, then wraps words. Blank lines only move `y` by `0.7 * leading`. When a quote closes, when the page fills, and after the last line, it strokes from the stored start `y` to the current `y`. Those long strokes are what keep the bar continuous across paragraph gaps, matching the Perl `write_long_text` routine.

Images use ImageMagick only. Each cell is sized from pixel width and height divided by the image’s own DPI (ImageMagick `resolution` and `units`, centimetres converted to inches, invalid DPI forced to 72, DPI under 50 multiplied by ten except under `--fast`). Physical inches times the layout’s document density give a target in points. If that target exceeds the cell, it is scaled down uniformly. The bitmap is resized without forcing aspect, centred on a white canvas of cell size, then `montage` tiles the cells. The PNG is placed with `drawImage` using the montage’s pixel size converted back through `image_dpi`. Profiles are left intact. `--finalize` compresses those PNGs; otherwise they are written uncompressed.

Attached PDFs are passed through Ghostscript `-dPDFFitPage` onto the master media box, then merged with pypdf. The first page of each such document is overlaid with a title bar `PDF: filename`.

### Assembling and large files

Each mail is a small PDF. Those pages are appended with pypdf into the destination file. pypdf writes a classic cross-reference table whose offsets are specified as exactly ten decimal digits. The PDF standard does not allow an eleven-digit classic offset. A file larger than 9 999 999 999 bytes cannot have a correct classic table. Viewers often reconstruct by scanning objects; qpdf `--check` reports `file is damaged` and `invalid xref`. That is a producer limit, not a missing “big PDF” version.

PDF 1.5 and later allow a cross-reference stream whose offset fields can be several bytes wide. qpdf implements that when `--object-streams=generate` is set. `--finalize` is that rewrite. Linearization is a different feature: hint tables with 32-bit positions, unusable past a few gigabytes, and not applied here.

pypdf stream-length caps are raised at import so large embedded object streams do not abort the merge. Pillow’s pixel cap is raised because ReportLab still decodes montage PNGs when drawing.

### Timeouts

External commands use a twenty-four hour budget that does not count time spent with the child in the `T` (stopped) state, so `Ctrl+Z` and `fg` do not burn the timer.

## Practical notes

A `--select-emails` conversion still rebuilds a stale dedup cache over the whole source, which will call LibreOffice on every office part in that source. A fresh cache plus a selector only converts the selected mail.

`--fast` once is the cheap visual check: same structure, coarse images. `--fast` twice or thrice is for structure only.

`qpdf --check` on an unfinalized multi-gigabyte assemble is expected to warn. Judge the master after `--finalize`.

Do not point `input_files` at the cache files. Do not combine `--finalize` with `--fast`. Do not expect linearize to repair a 15 GB file.

The listing number in the PDF header, the archive folder name, and the first TSV column are the same index after the same sort of the same loaded mailbox.
