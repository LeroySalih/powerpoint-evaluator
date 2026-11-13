# powerpoint-evaluator

Upload a pupil's `.pptx` (plus optional PDFs/screenshots) and a JSON spec of requirements; get back a per-check report: what passed, what failed, what needs a teacher's eye.

```bash
pip install -e '.[dev]'
pptx-eval evaluate --spec specs/pearl-kiosk.json --pptx "PEARL KIOSK.pptx" \
    --extra HANDOUTS=HANDOUTS.pdf --extra NOTES=NOTES.pdf --extra "FULL PAGE=FULL PAGE.pdf"
uvicorn pptx_eval.api:app          # POST /evaluate (multipart)
```

See `docs/DESIGN.md` for the spec format, the check catalogue and what can and cannot be automated.
