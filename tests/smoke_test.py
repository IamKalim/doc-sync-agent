#!/usr/bin/env python3
"""End-to-end smoke test: builds sample docs + a git repo in a temp dir, then
runs init -> scan -> apply (both modes, all edit types) and checks the results.

    python tests/smoke_test.py
"""
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parent.parent
CLI = [sys.executable, str(ROOT / "docsync.py")]


def run(args, env, expect=0):
    r = subprocess.run(CLI + args, capture_output=True, text=True, env=env, stdin=subprocess.DEVNULL)
    if r.returncode != expect:
        print(r.stdout, r.stderr)
        raise SystemExit(f"FAILED: docsync {' '.join(args)} (exit {r.returncode}, expected {expect})")
    return r.stdout + r.stderr


def make_module_guide(path, module):
    d = Document()
    d.add_paragraph(f"{module} Module Guide", style="Title")
    d.add_heading("Overview", 1)
    d.add_paragraph(f"The {module} module lets users work with {module.lower()} features in the portal.")
    d.add_heading("Prerequisites", 1)
    d.add_paragraph("A user account with the correct role and access to the admin panel.")
    d.add_heading("Configuration", 1)
    d.add_paragraph("Open Settings and set the options listed below before first use.")
    d.add_paragraph("Enable the module", style="List Bullet")
    d.add_paragraph("Set the default role", style="List Bullet")
    d.add_heading("FAQ", 1)
    d.add_paragraph("Contact the support team if the module does not appear in the menu.")
    path.parent.mkdir(parents=True, exist_ok=True)
    d.save(path)


def main():
    tmp = Path(tempfile.mkdtemp(prefix="docsync_test_"))
    docs, repo, state = tmp / "docs", tmp / "code", tmp / "state"
    env = dict(os.environ, DOCSYNC_HOME=str(state))

    make_module_guide(docs / "Modules/Login/Login_Module_Guide.docx", "Login")
    make_module_guide(docs / "Modules/Reports/Reports_Module_Guide.docx", "Reports")

    api = Document()
    api.add_heading("API Guide", 0)
    api.add_heading("Authentication", 1)
    api.add_paragraph("Call the login endpoint to obtain a token. Tokens expire after 30 minutes.")
    api.add_heading("Endpoints", 1)
    t = api.add_table(rows=2, cols=2)
    t.style = "Table Grid"
    t.cell(0, 0).text, t.cell(0, 1).text = "Endpoint", "Purpose"
    t.cell(1, 0).text, t.cell(1, 1).text = "/login", "Get a token"
    api.add_heading("Errors", 1)
    api.add_paragraph("Errors are returned as JSON with a code and message.")
    (docs / "API").mkdir(parents=True, exist_ok=True)
    api.save(docs / "API/API_Guide.docx")

    idx = Document()
    idx.add_heading("Documentation Index", 1)
    idx.add_heading("Modules", 2)
    idx.add_paragraph("Login_Module_Guide", style="List Bullet")
    idx.add_paragraph("Reports_Module_Guide", style="List Bullet")
    idx.add_paragraph("API_Guide", style="List Bullet")
    idx.save(docs / "Index.docx")

    (docs / "Notes.md").write_text("# Notes\n\n## Release process\n\nTag the release.\n\n## Contacts\n\n- Alice\n", "utf-8")

    # code repo with two commits
    repo.mkdir()
    g = lambda *a: subprocess.run(["git", "-C", str(repo), *a], check=True, capture_output=True)
    g("init", "-q")
    g("config", "user.email", "t@t.t")
    g("config", "user.name", "Tester")
    (repo / "auth.py").write_text("TOKEN_TTL_MINUTES = 30\n")
    g("add", ".")
    g("commit", "-qm", "Add auth")
    (repo / "auth.py").write_text("TOKEN_TTL_MINUTES = 60\n")
    (repo / "payments.py").write_text("def charge(): ...\n")
    g("add", ".")
    g("commit", "-qm", "Token expiry 60m; add Payments module")

    # --- init: non-interactive without --root must ask for input
    out = run(["init"], env, expect=1)
    assert "NEEDS_INPUT" in out, out
    run(["init", "--root", str(docs), "--code-repo", str(repo)], env)

    # --- scan
    out = run(["scan"], env)
    index = (state / "index.md").read_text()
    structure = (state / "structure.md").read_text()
    changes = (state / "changes.md").read_text()
    assert "DOC: Modules/Login/Login_Module_Guide.docx" in index
    assert "Token expiry 60m" in changes and "TOKEN_TTL_MINUTES = 60" in changes
    assert "Overview" in structure and "Words_With_Underscores" in structure
    assert "Index.docx" in structure, structure
    print("scan OK")

    # --- apply needs a mode
    edits = {
        "task": "Token expiry + Payments module",
        "edits": [
            {"type": "update_section", "path": "API/API_Guide.docx", "heading": "Authentication",
             "new_text": "Call the login endpoint to obtain a token. Tokens now expire after **60 minutes**.",
             "reason": "TOKEN_TTL_MINUTES 30 -> 60"},
            {"type": "add_section", "path": "API/API_Guide.docx", "heading": "Refresh tokens", "level": 1,
             "after_heading": "Authentication", "body": "Tokens can be refreshed before expiry.\n\n- Call /refresh\n- Use the new token"},
            {"type": "new_document", "path": "Modules/Payments/Payments_Module_Guide.docx",
             "template_from": "Modules/Login/Login_Module_Guide.docx", "title": "Payments Module Guide",
             "sections": [
                 {"heading": "Overview", "level": 1, "body": "The Payments module handles charges."},
                 {"heading": "Configuration", "level": 1,
                  "body": "Set the gateway key.\n\n| Setting | Value |\n|---|---|\n| Gateway | TODO: confirm with developer |\n\n- Enable the module"},
                 {"heading": "FAQ", "level": 1, "body": "TODO: confirm with developer"}]},
            {"type": "append_to_section", "path": "Index.docx", "heading": "Modules",
             "body": "- Payments_Module_Guide"},
            {"type": "update_section", "path": "Notes.md", "heading": "Release process",
             "new_text": "Tag the release, then publish the notes."},
            {"type": "add_section", "path": "Notes.md", "heading": "Payments", "level": 2,
             "after_heading": "Release process", "body": "Payments ships in the next release."},
        ],
    }
    ef = tmp / "edits.json"
    ef.write_text(json.dumps(edits))
    out = run(["apply", str(ef)], env, expect=1)
    assert "NEEDS_INPUT" in out

    run(["set-mode", "new-version"], env)
    out = run(["apply", str(ef), "--dry-run"], env)
    assert "-Call the login" in out or "- Call the login" in out or "before" in out
    assert not (docs / "API/API_Guide_v2.docx").exists()
    print("dry-run OK")

    # --- bad edit aborts everything
    bad = tmp / "bad.json"
    bad.write_text(json.dumps({"edits": [edits["edits"][0],
                                         {"type": "update_section", "path": "API/API_Guide.docx", "heading": "Nope", "new_text": "x"}]}))
    out = run(["apply", str(bad)], env, expect=2)
    assert "heading not found" in out and not (docs / "API/API_Guide_v2.docx").exists()
    print("atomic abort OK")

    # --- real apply, new-version mode
    out = run(["apply", str(ef)], env)
    assert (docs / "API/API_Guide.docx").exists() and (docs / "API/API_Guide_v2.docx").exists()
    assert (docs / "Index_v2.docx").exists() and (docs / "Notes_v2.md").exists()
    assert (docs / "Modules/Payments/Payments_Module_Guide.docx").exists()

    orig = Document(docs / "API/API_Guide.docx")
    assert "30 minutes" in " ".join(p.text for p in orig.paragraphs), "original must be untouched"
    v2 = Document(docs / "API/API_Guide_v2.docx")
    texts = [p.text for p in v2.paragraphs]
    assert any("60 minutes" in t for t in texts) and not any("30 minutes" in t for t in texts)
    heads = [p.text for p in v2.paragraphs if p.style.name.startswith("Heading")]
    assert heads.index("Refresh tokens") == heads.index("Authentication") + 1, heads
    assert len(v2.tables) == 1, "table must survive"

    new = Document(docs / "Modules/Payments/Payments_Module_Guide.docx")
    nh = [(p.style.name, p.text) for p in new.paragraphs if p.text.strip()]
    assert nh[0] == ("Title", "Payments Module Guide"), nh
    assert ("Heading 1", "Overview") in nh and ("Heading 1", "FAQ") in nh, nh
    assert any(s == "List Bullet" for s, _ in nh), "bullets should reuse list style"
    assert len(new.tables) == 1
    assert not any("Login" in p.text for p in new.paragraphs), "template content must be cleared"

    iv2 = Document(docs / "Index_v2.docx")
    assert [p.text for p in iv2.paragraphs][-1] == "Payments_Module_Guide"
    assert "Payments" in (docs / "Notes_v2.md").read_text()
    assert "publish the notes" in (docs / "Notes_v2.md").read_text()
    assert "Tag the release.\n" in (docs / "Notes.md").read_text()
    changelog = (docs / "_docsync/changelog.md").read_text()
    assert "Payments module" in changelog and "API_Guide_v2.docx" in changelog
    print("new-version mode OK")

    # --- scan again: only new commits (none), superseded versions detected
    run(["scan"], env)
    assert "No new commits" in (state / "changes.md").read_text() and "Nothing new was found in git" in (state / "changes.md").read_text()
    assert "SUPERSEDED by API/API_Guide_v2.docx" in (state / "index.md").read_text()

    # --- in-place mode with backup, targeting the latest version
    run(["set-mode", "in-place"], env)
    e2 = tmp / "e2.json"
    e2.write_text(json.dumps({"task": "in-place test", "edits": [
        {"type": "update_section", "path": "Modules/Login/Login_Module_Guide.docx", "heading": "FAQ",
         "new_text": "Contact support@example.com if the module does not appear."}]}))
    run(["apply", str(e2)], env)
    login = Document(docs / "Modules/Login/Login_Module_Guide.docx")
    assert any("support@example.com" in p.text for p in login.paragraphs)
    backups = list((docs / "_docsync/backups").rglob("Login_Module_Guide.docx"))
    assert len(backups) == 1
    assert not any("support@example.com" in p.text for p in Document(backups[0]).paragraphs)
    assert not (docs / "Modules/Login/Login_Module_Guide_v2.docx").exists()
    print("in-place mode OK")

    # --- safety: path traversal + media guard + duplicate new doc
    e3 = tmp / "e3.json"
    e3.write_text(json.dumps({"edits": [{"type": "new_document", "path": "../evil.md", "sections": []}]}))
    assert "escapes" in run(["apply", str(e3)], env, expect=2)
    e4 = tmp / "e4.json"
    e4.write_text(json.dumps({"edits": [{"type": "new_document", "path": "Index.docx", "sections": []}]}))
    assert "already exists" in run(["apply", str(e4)], env, expect=2)
    print("safety checks OK")

    print(f"\nALL TESTS PASSED  (sandbox: {tmp})")


if __name__ == "__main__":
    main()
