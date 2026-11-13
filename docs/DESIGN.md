# Design

## What it does
A caller sends: a pupil's `.pptx`, a JSON **spec** (requirements, marks, reference images), and optionally the pupil's other files (HANDOUTS/NOTES/FULL PAGE PDFs, EVIDENCE) plus `{name, class}`. It returns a **report**: every check is `pass`, `fail` or `review`, with marks, a one-line reason, and totals by learning objective.

`review` means "a person (or later a vision/LLM judge) must look". The engine never guesses a pass. Confirmed marks and marks pending review are reported separately.

## Stack
Python 3.10+ · `zipfile` + `lxml` (read the .pptx XML directly; nothing is executed) · `pydantic` (spec/report) · `Pillow` (perceptual image matching, contrast) · `pypdf` (printed PDFs) · `FastAPI`/`uvicorn` (service) · `pytest`. LibreOffice (headless) is planned for rendering slides to images for the visual judge.

## Spec format (`specs/*.json`)
```json
{ "id": "...", "title": "...",
  "assets": { "TURTLE": { "file": "pearl-kiosk/assets/TURTLE.jpg" } },   // or { "data_base64": "..." }
  "sections": [ { "id": "T1", "title": "...", "checks": [
      { "id": "1b", "description": "...", "marks": 1, "tags": ["LO1"],
        "type": "master.background_image", "params": { "asset": "BACKGROUND", "transparency_pct": [55, 65] } } ] } ] }
```
`type: "all_of"` with `params.checks` combines sub-checks (any fail → fail; else any review → review). Used where a mark needs both a file fact and a screenshot.

Reference images are matched by perceptual hash, so a picture that PowerPoint recompressed or resized still matches. With a reference image supplied, only the image *content* counts; the file name in alt text is a fallback when no reference was given.

## Check catalogue
| type | checks |
|---|---|
| `deck.slide_ratio` | slide size ratio (16:10 = 1.6; compares real dimensions, not PowerPoint's `type` label) |
| `master.background_image` | effective background image + transparency, per slide |
| `master.title_style` | effective title bold/italic/dark-blue and contrast vs the blended background |
| `master.logo` | reference image visible (slide/layout/master) in the top right |
| `master.footer_text` / `master.slide_number` | footer text + centred, auto date, slide number right, title slide clean |
| `layout.placeholder_hidden` / `_added` | date removed from a layout (and slides); picture placeholder added and used |
| `slides.layouts` · `short_bullets` · `images` · `image_replaces_text` · `image_size_cm` · `fit_for_purpose` | structure, text length, images, exact sizes, font size/overlap/stretch |
| `anim.same_entrance` · `after_previous_in_order` · `duration` · `transition.auto_advance` | Animation Pane facts and slide transition timing |
| `links.slide_links` · `home_button` · `external` · `file` | hyperlinks (text or whole-shape), action buttons, URLs/mailto, file links |
| `notes.present` | speaker notes |
| `pdf.print` · `print.name_class` | page count, orientation, slide titles/notes text in printed PDFs; name+class in footer |
| `evidence.screenshot` · `answer.rubric` | need a judge → `review` (or `fail` if EVIDENCE not supplied) |

## What can be automated (PEARL KIOSK, Section A, 32 marks)
Deterministic from the .pptx: 1b–1i, 2a–2e, 3a–3d, 4a–4c, 4e, 4f, 5a (≈ 22 marks).
Deterministic from the PDFs: 5c–5e (3). Needs the candidate's name/class: 5b (1).
Partly automatable, final call needs eyes: 2f (measurements automated, "looks good" not).
Screenshots in EVIDENCE: 1a (file fact is automated), 1j, 3e, 4d (judge needed). Section B (13 marks, short answers): judge needed.

## Findings from the model answer (things the spec must tolerate)
- Menu links are on whole shapes (rounded rectangles), not text runs: both are accepted.
- Slide size is stored as `type="screen4x3"` with 16:10 dimensions: check the dimensions.
- Date is `datetime1` on slides, `datetimeFigureOut` on the master: any `datetime*` field is auto-updating.
- Footer text sits on each slide's footer placeholder; the master's footer placeholder is empty. Judge the slide, not the master.
- TURTLE and JELLYFISH are free pictures (not in placeholders), although task 2(c) says "placeholders". Only slide 2's image is required to be in a placeholder (matching mark scheme 1i). **Decide whether 2(c) should require placeholders on 3 and 6**; the model answer itself does not.
- On slide 5 the Home button overlaps the MAP by about 15% of the button. Reported as a soft overlap for review, not a fail.

## Not yet built
1. Judge for screenshots, "fit for purpose" visuals and Section B (render with LibreOffice, then a Claude vision/text call behind a `Judge` interface).
2. Section B spec entries (keep the rubric text out of this public repo).
3. Batch mode (a class set / DINO upload) and a minimal upload page.
4. Auth, rate limiting and storage policy for the hosted service. Pupil files are processed in memory and not stored or logged.

## Public repo rules
Mark schemes, model answers and pupil work are never committed (`private/` is git-ignored, as are reference images). The committed spec contains only what the pupil-facing task sheet already states.
