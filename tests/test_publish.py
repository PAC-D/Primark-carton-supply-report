import datetime
import json
import subprocess

import pytest

from pathlib import Path

from filtering import build_table, default_duration, range_months
from models import Record
from publish import commit_site, publish_site

PROJECT_ROOT = Path(__file__).parent.parent


def make_records():
    return [
        Record("M&U", "S1", "F1", 2026, 1, 10.0),
        Record("M&U", "S1", "F1", 2026, 2, 20.0),
        Record("M&U", "S2", "F2", 2026, 1, 5.0),
    ]


def test_publish_writes_all_files(tmp_path):
    out = publish_site(make_records(), tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    for name in ("index.html", "app.js", "logic.js", "data.json", "site-manifest.json"):
        assert (out / name).exists()


def test_publish_data_matches_build_table(tmp_path):
    records = make_records()
    out = publish_site(records, tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    data = json.loads((out / "data.json").read_text(encoding="utf-8"))
    df = build_table(records)
    assert data["columns"] == df.columns.tolist()
    assert data["rows"] == df.values.tolist()
    frm, to = default_duration(records)
    assert data["months"] == [list(p) for p in range_months(frm, to)]
    assert len(data["months"]) == len(data["columns"]) - 5


def test_publish_manifest(tmp_path):
    out = publish_site(make_records(), tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    manifest = json.loads((out / "site-manifest.json").read_text(encoding="utf-8"))
    assert manifest == {"published": "2026-08-19 14:30", "row_count": 2}


def test_publish_empty_records(tmp_path):
    out = publish_site([], tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    data = json.loads((out / "data.json").read_text(encoding="utf-8"))
    assert data["rows"] == []
    assert data["columns"] == ["SL", "Packaging Supplier", "Supplier", "Factory", "Total"]
    assert data["months"] == []
    manifest = json.loads((out / "site-manifest.json").read_text(encoding="utf-8"))
    assert manifest["row_count"] == 0


def test_publish_copies_templates(tmp_path):
    out = publish_site(make_records(), tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    for name in ("index.html", "app.js", "logic.js"):
        assert (out / name).read_text(encoding="utf-8") == (PROJECT_ROOT / "site_template" / name).read_text(encoding="utf-8")


def _git(repo, *args):
    return subprocess.run(
        ["git", *args], cwd=repo, capture_output=True, text=True, check=True
    )


def test_commit_site(tmp_path):
    repo = tmp_path / "repo"
    repo.mkdir()
    _git(repo, "init", "-b", "main")
    _git(repo, "config", "user.email", "test@example.com")
    _git(repo, "config", "user.name", "Test")
    out = publish_site(make_records(), repo / "docs" / "site", datetime.datetime(2026, 8, 19, 14, 30))
    commit_site(repo, "docs/site", datetime.datetime(2026, 8, 19, 14, 30))
    log = _git(repo, "log", "--oneline", "-1")
    assert "Updated at 2026-08-19 14:30" in log.stdout
    tracked = _git(repo, "ls-files", "docs/site")
    assert "docs/site/data.json" in tracked.stdout