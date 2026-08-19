import datetime
import json
import shutil
import subprocess
from pathlib import Path

from filtering import build_table, default_duration, range_months

TEMPLATE_DIR = Path(__file__).parent / "site_template"


def _table_data(records):
    duration = default_duration(records)
    if duration is None:
        return ["SL", "Packaging Supplier", "Supplier", "Factory", "Total"], [], []
    frm, to = duration
    months = range_months(frm, to)
    df = build_table(records)
    return df.columns.tolist(), df.values.tolist(), months


def publish_site(records, out_dir, now):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    columns, rows, months = _table_data(records)
    (out_dir / "data.json").write_text(
        json.dumps({"columns": columns, "months": months, "rows": rows}),
        encoding="utf-8",
    )
    (out_dir / "site-manifest.json").write_text(
        json.dumps({
            "published": now.isoformat(sep=" ", timespec="minutes"),
            "row_count": len(rows),
        }),
        encoding="utf-8",
    )
    for name in ("index.html", "app.js", "logic.js"):
        (out_dir / name).write_text(
            (TEMPLATE_DIR / name).read_text(encoding="utf-8"), encoding="utf-8"
        )
    for name in ("css/style.css", "favicon.png", "pacd.png", "assets/primark-logo.png"):
        (out_dir / name).parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(TEMPLATE_DIR / name, out_dir / name)
    return out_dir


def commit_site(repo_dir, out_dir, now):
    stamp = now.strftime("%Y-%m-%d %H:%M")
    add = subprocess.run(
        ["git", "add", str(out_dir)],
        cwd=repo_dir, capture_output=True, text=True,
    )
    if add.returncode != 0:
        raise RuntimeError(f"git add failed: {add.stderr.strip()}")
    commit = subprocess.run(
        ["git", "commit", "-m", f"Updated at {stamp}"],
        cwd=repo_dir, capture_output=True, text=True,
    )
    if commit.returncode != 0:
        raise RuntimeError(f"git commit failed: {commit.stderr.strip()}")
    return commit.stdout.strip()