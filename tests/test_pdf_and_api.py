import io
import json
import shutil
import subprocess
import tempfile
from pathlib import Path

import pytest
from pypdf import PdfReader, PdfWriter

from conftest import MODEL, SPEC, needs_private
from pptx_eval.engine import evaluate
from pptx_eval.inspect import BadPackage, inspect_pptx
from pptx_eval.spec import Spec

soffice = shutil.which("soffice")


@pytest.fixture(scope="module")
def slides_pdf() -> bytes:
    if not (soffice and MODEL.is_file()):
        pytest.skip("needs LibreOffice and the private model answer")
    with tempfile.TemporaryDirectory() as d:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", d, str(MODEL)],
                       check=True, capture_output=True, timeout=120, env={"HOME": d, "PATH": "/usr/bin:/bin"})
        return (Path(d) / (MODEL.stem + ".pdf")).read_bytes()


def pages(pdf: bytes, idx: list[int]) -> bytes:
    r, w = PdfReader(io.BytesIO(pdf)), PdfWriter()
    for i in idx:
        w.add_page(r.pages[i])
    out = io.BytesIO(); w.write(out)
    return out.getvalue()


def run(extras):
    spec = Spec.model_validate_json(SPEC.read_text(encoding="utf8"))
    return {r.id: r for r in evaluate(spec, MODEL.read_bytes(), {"EVIDENCE": b"x"} | extras, None, SPEC.parent).results}


@needs_private
def test_full_page_pass_and_fail(slides_pdf):
    assert run({"FULL PAGE.pdf": pages(slides_pdf, [0])})["5e"].status == "pass"
    assert run({"FULL PAGE.pdf": pages(slides_pdf, [1])})["5e"].status == "fail"  # wrong slide
    assert run({"FULL PAGE.pdf": pages(slides_pdf, [0, 1])})["5e"].status == "fail"  # 2 pages


@needs_private
def test_handouts_must_be_portrait_two_pages(slides_pdf):
    res = run({"HANDOUTS.pdf": slides_pdf})  # 6 landscape pages: the wrong layout
    assert res["5c"].status == "fail" and "6 page" in res["5c"].detail


@needs_private
def test_extra_names_are_forgiving():
    for name in ("handouts.PDF", "HANDOUTS", "Handouts.pdf"):
        assert "not supplied" not in run({name: b"%PDF"})["5c"].detail


def test_not_a_pptx_is_rejected():
    with pytest.raises(BadPackage):
        inspect_pptx(b"this is not a zip")


@needs_private
def test_api_roundtrip():
    from fastapi.testclient import TestClient
    from pptx_eval.api import app
    c = TestClient(app)
    spec = json.loads(SPEC.read_text(encoding="utf8"))
    from conftest import ASSETS
    import base64
    for name, a in spec["assets"].items():  # embed reference images in the JSON, as a remote caller would
        a["data_base64"] = base64.b64encode((SPEC.parent / a.pop("file")).read_bytes()).decode()
    r = c.post("/evaluate", data={"spec": json.dumps(spec), "candidate": json.dumps({"name": "A", "class": "B"})},
               files=[("submission", ("k.pptx", MODEL.read_bytes())), ("files", ("EVIDENCE.pdf", b"x"))])
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["marks_available"] == 32 and body["marks_awarded"] >= 20
    assert c.post("/evaluate", data={"spec": "{}"}, files=[("submission", ("k.pptx", b"x"))]).status_code == 422
    assert c.post("/evaluate", data={"spec": json.dumps(spec)}, files=[("submission", ("k.pptx", b"nope"))]).status_code == 422
