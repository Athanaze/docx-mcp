"""Test DOCUMENT_ROOT path resolution (#27)."""
import sys, os
sys.stdout.reconfigure(line_buffering=True)
sys.path.insert(0, '/home/s/Office-Word-MCP-Server')

from word_document_server.utils.file_utils import (
    ensure_docx_extension, resolve_document_path, get_document_root
)

PASS = FAIL = 0
def test(name, cond, detail=""):
    global PASS, FAIL
    if cond: PASS += 1; print(f"  ✅ {name}", flush=True)
    else: FAIL += 1; print(f"  ❌ {name} — {detail}", flush=True)

# === Test 1: No DOCUMENT_ROOT (backwards compatible) ===
print("[1] No DOCUMENT_ROOT set (default behavior)", flush=True)
if 'DOCUMENT_ROOT' in os.environ:
    del os.environ['DOCUMENT_ROOT']

result = ensure_docx_extension("report")
test("Adds .docx", result.endswith("report.docx"), result)
test("Returns absolute path", os.path.isabs(result), result)
test("Resolves against CWD", result == os.path.join(os.getcwd(), "report.docx"), result)

result2 = ensure_docx_extension("/absolute/path/report.docx")
test("Absolute path unchanged", result2 == "/absolute/path/report.docx", result2)

# === Test 2: DOCUMENT_ROOT set ===
print("\n[2] DOCUMENT_ROOT=/tmp/test_docroot", flush=True)
os.environ['DOCUMENT_ROOT'] = '/tmp/test_docroot'

result = ensure_docx_extension("report")
test("Resolves against DOCUMENT_ROOT", result == "/tmp/test_docroot/report.docx", result)
test("Directory created", os.path.isdir("/tmp/test_docroot"))

result2 = ensure_docx_extension("subdir/report.docx")
test("Subdirectory relative path", result2 == "/tmp/test_docroot/subdir/report.docx", result2)

result3 = ensure_docx_extension("/absolute/report.docx")
test("Absolute path still absolute", result3 == "/absolute/report.docx", result3)

# === Test 3: DOCUMENT_ROOT with ~ ===
print("\n[3] DOCUMENT_ROOT=~/Documents", flush=True)
os.environ['DOCUMENT_ROOT'] = '~/Documents'
root = get_document_root()
test("Tilde expanded", root and '~' not in root, f"root={root}")
test("Path is absolute", root and os.path.isabs(root), f"root={root}")

# === Test 4: End-to-end with actual document creation ===
print("\n[4] End-to-end: create document via tool with DOCUMENT_ROOT", flush=True)
docroot = '/tmp/test_docroot_e2e'
os.environ['DOCUMENT_ROOT'] = docroot

import asyncio
from word_document_server.tools.document_tools import create_document
result = asyncio.run(create_document("e2e_test.docx", title="Test"))
print(f"    create_document result: {result}", flush=True)
expected_path = os.path.join(docroot, "e2e_test.docx")
test("Document created at DOCUMENT_ROOT", os.path.exists(expected_path), f"expected: {expected_path}")
if os.path.exists(expected_path):
    test("File has content", os.path.getsize(expected_path) > 0)

# Cleanup
import shutil
for d in ['/tmp/test_docroot', '/tmp/test_docroot_e2e']:
    if os.path.exists(d): shutil.rmtree(d)
del os.environ['DOCUMENT_ROOT']

print(f"\n{'='*50}", flush=True)
print(f"Results: {PASS}/{PASS+FAIL} passed, {FAIL}/{PASS+FAIL} failed", flush=True)
if FAIL == 0: print("🎉 All passed!", flush=True)
else: print(f"⚠️ {FAIL} failed", flush=True); sys.exit(1)
