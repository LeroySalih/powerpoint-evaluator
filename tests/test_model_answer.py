"""The model answer must pass what it should; each mutation must trip exactly the check it targets."""
import io

import pytest
from PIL import Image

from conftest import ASSETS, MODEL, SPEC, many, mutate, needs_private
from pptx_eval.engine import evaluate
from pptx_eval.spec import Spec

pytestmark = needs_private
EVID = {"EVIDENCE": b"x"}  # screenshots exist -> those marks become 'review', not 'fail'


def run(data: bytes | None = None, extras=None, candidate=None):
    spec = Spec.model_validate_json(SPEC.read_text(encoding="utf8"))
    rep = evaluate(spec, data or MODEL.read_bytes(), EVID | (extras or {}), candidate, SPEC.parent)
    return {r.id: r for r in rep.results}


def status(res, *ids):
    return {i: res[i].status for i in ids}


def test_model_answer_clean():
    res = run()
    auto = [i for i in res if i[0] in "1234" and i not in ("1a", "1j", "2f", "3e", "4d")]
    assert {i: res[i].status for i in auto if res[i].status != "pass"} == {}
    assert res["2f"].status == "review" and res["1a"].status == "review"
    assert res["5a"].status == "pass"


def mut(path, *pairs):
    return run(mutate(MODEL.read_bytes(), {path: many(*pairs)}))


@pytest.mark.parametrize("path,pairs,check", [
    ("ppt/presentation.xml", [('cy="5715000"', 'cy="5143500"')], "1a"),  # 16:9
    ("ppt/slideMasters/slideMaster1.xml", [('amt="40000"', 'amt="100000"')], "1b"),  # opaque background
    ("ppt/slideMasters/slideMaster1.xml", [('b="1" i="1"><a:solidFill><a:srgbClr val="073F5C"', 'b="1" i="1"><a:solidFill><a:srgbClr val="000000"')], "1c"),  # black title
    ("ppt/slideMasters/slideMaster1.xml", [(' b="1" i="1"><a:solidFill><a:srgbClr val="073F5C"', ' b="1"><a:solidFill><a:srgbClr val="073F5C"')], "1d"),  # not italic
    ("ppt/slides/slide2.xml", [('type="datetime1"', 'type="notadate"')], "1f"),  # typed date
    ("ppt/slides/slide3.xml", [('cx="3600000" cy="2700000"', 'cx="3000000" cy="2700000"')], "2e"),
    ("ppt/slides/slide3.xml", [('dur="2000"', 'dur="500"')], "3c"),
    ("ppt/slides/slide3.xml", [('nodeType="afterEffect"', 'nodeType="clickEffect"')], "3b"),
    ("ppt/slides/slide3.xml", [('spid="3"', 'spid="2"')], "3a"),  # animates the title instead
    ("ppt/slides/slide1.xml", [('advClick="0"', 'advClick="1"')], "3d"),
    ("ppt/slides/slide1.xml", [('advTm="5000"', 'advTm="3000"')], "3d"),
    ("ppt/slides/slide6.xml", [('<a:off x="216000"', '<a:off x="300000"')], "4c"),  # button moved
    ("ppt/slides/_rels/slide6.xml.rels", [('http://www.pearlcoast.service', 'http://www.pearlcoast.servce')], "4e"),
    ("ppt/slides/_rels/slide4.xml.rels", [('TICKET PRICES.xlsx', 'https://example.com/prices')], "4f"),
])
def test_mutation_fails_target_check(path, pairs, check):
    res = mut(path, *pairs)
    assert res[check].status == "fail", res[check]


def test_home_button_to_wrong_slide_rejected():
    res = mut("ppt/slides/_rels/slide3.xml.rels", ('Target="slide2.xml"', 'Target="slide4.xml"'))
    assert res["4c"].status == "fail"


def test_home_button_to_first_slide_action_rejected():
    res = mut("ppt/slides/slide6.xml", ('<a:hlinkClick action="ppaction://hlinksldjump" r:id="rId5"/>',
                                        '<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=firstslide"/>'))
    assert res["4c"].status == "fail"


def test_wrong_menu_link_rejected():
    res = mut("ppt/slides/_rels/slide2.xml.rels", ('Target="slide3.xml"', 'Target="slide6.xml"'))
    assert res["4a"].status == "fail"


def _swap_media(name, src_asset):
    return {f"ppt/media/{name}": lambda _b: (ASSETS / src_asset).read_bytes()}


def test_map_replaced_by_other_image_fails_2d():
    res = run(mutate(MODEL.read_bytes(), _swap_media("image4.png", "TURTLE.jpg")))
    assert res["2d"].status == "fail"


def test_recompressed_resized_image_still_matches():
    im = Image.open(ASSETS / "TURTLE.jpg").resize((600, 450))
    buf = io.BytesIO(); im.save(buf, "JPEG", quality=55)
    res = run(mutate(MODEL.read_bytes(), {"ppt/media/image3.jpg": lambda _b: buf.getvalue()}))
    assert res["2c"].status == "pass" and res["2e"].status == "pass"


def test_wrong_image_on_slide3_fails_2c():
    res = run(mutate(MODEL.read_bytes(), _swap_media("image3.jpg", "CORAL.jpg")))
    assert res["2c"].status == "fail"


def test_stretched_image_flagged_in_2f():
    res = mut("ppt/slides/slide3.xml", ('cx="3600000" cy="2700000"', 'cx="3600000" cy="1500000"'))
    assert res["2f"].status == "fail" and "stretched" in res["2f"].detail


def test_missing_notes_fails_5a():
    res = run(mutate(MODEL.read_bytes(), {"ppt/slides/_rels/slide4.xml.rels": many(('notesSlide', 'notesXXX'))}))
    assert res["5a"].status == "fail"


def test_candidate_name_in_footer():
    res = run(candidate={"name": "Zed Pupil", "class": "10X"})
    assert res["5b"].status == "fail"  # model answer footer has no pupil name
