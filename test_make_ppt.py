"""Model-free checks: python test_make_ppt.py"""
import make_ppt as mp


def test_split_source_keeps_everything():
    text = "# Source: a.pdf\n\n" + "".join(f"## Page {i}\n{'x' * 900}\n" for i in range(1, 11))
    parts = mp.split_source(text, 2500)
    assert "".join(parts) == text and len(parts) > 1
    assert all(len(p) <= 2500 * 1.15 for p in parts)      # a short tail joins the last part
    assert all(p.startswith("#") for p in parts)          # cut at page breaks
    assert mp.split_source("y" * 7000, 3000) == ["y" * 3000, "y" * 3000, "y" * 1000]  # no breaks: hard cut


def test_diversify_and_palette():
    lead = {"type": "bullets", "title": "Forms", "bullets": ["Form 1: apply", "Form 2 – agree", "Form 3 - sign"]}
    spec = {"palette": "forest | coral | teal", "slides": [
        {"type": "title", "title": "T"}, lead,
        {"type": "bullets", "title": "How to apply - steps", "bullets": ["A: one", "B: two", "C: three"]},
        {"type": "bullets", "title": "Loose", "bullets": ["no separator here", "B: two", "C: three"]}]}
    mp.normalize(spec)
    assert spec["palette"] == "forest"
    assert mp.diversify(spec) == 2
    assert lead["type"] == "cards" and lead["cards"][1] == {"title": "Form 2", "text": "agree"}
    assert [s["type"] for s in spec["slides"]] == ["title", "cards", "steps", "bullets"]
    assert not mp.validate(spec, mp.Path("."))


def test_script_toggle_and_hints():
    import json
    seen = []

    def fake(model, prompt, reply=0):
        seen.append(prompt)
        n = int(prompt.split("holding exactly")[1].split()[0])
        return json.dumps({"notes": ["Explained. Reference: a.pdf, p1"] * n, "art": ["a house"] * n})

    mp.ask_llm = fake
    mk = lambda: [{"type": "bullets", "title": "A", "bullets": ["x"]}, {"type": "bullets", "title": "B", "bullets": ["y"]}]
    on = mk()
    mp.add_scripts("m", "T", on, "src", "short and funny", art=True, script=True)
    assert all("Reference" in s["notes"] and s["art"] == "a house" for s in on)
    assert "speaker script" in seen[0] and "short and funny" in seen[0]      # hints reach the script pass
    off = mk()
    mp.add_scripts("m", "T", off, "src", "", art=True, script=False)         # pictures only: no script asked for
    assert all("notes" not in s and s["art"] for s in off) and "speaker script" not in seen[1]
    assert mp.SCRIPT_HINT.search("make the script for each slide") and not mp.SCRIPT_HINT.search("keep it short")


if __name__ == "__main__":
    test_script_toggle_and_hints()
    test_split_source_keeps_everything()
    test_diversify_and_palette()
    print("ok")
