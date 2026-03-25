"""Test convert_to_pdf after #26 fixes."""
import sys, os, shutil, time
sys.stdout.reconfigure(line_buffering=True)
sys.path.insert(0, '/home/s/Office-Word-MCP-Server')

import asyncio
from word_document_server.tools.extended_document_tools import convert_to_pdf

FIXTURES = '/home/s/Office-Word-MCP-Server/tests/fixtures'
SRC = os.path.join(FIXTURES, 'par-hlink-frags.docx')

PASS = FAIL = 0
def test(name, cond, detail=""):
    global PASS, FAIL
    if cond: PASS += 1; print(f"  ✅ {name}", flush=True)
    else: FAIL += 1; print(f"  ❌ {name} — {detail}", flush=True)

# Test 1: Basic conversion
print("[1] Basic absolute path conversion", flush=True)
out1 = os.path.join(FIXTURES, 'test_out1.pdf')
if os.path.exists(out1): os.remove(out1)
t0 = time.time()
r = asyncio.run(convert_to_pdf(SRC, out1))
dt = time.time() - t0
print(f"    Result: {r}", flush=True)
print(f"    Time: {dt:.1f}s", flush=True)
test("PDF created", os.path.exists(out1))
test("Success message format", "Output:" in r and "KB" in r)
if os.path.exists(out1):
    test("PDF has content", os.path.getsize(out1) > 1000, f"{os.path.getsize(out1)} bytes")
    os.remove(out1)

# Test 2: Two rapid successive calls (the lock file scenario)
print("\n[2] Two rapid successive calls (lock file test)", flush=True)
out2a = os.path.join(FIXTURES, 'test_out2a.pdf')
out2b = os.path.join(FIXTURES, 'test_out2b.pdf')
for f in [out2a, out2b]:
    if os.path.exists(f): os.remove(f)

t0 = time.time()
r1 = asyncio.run(convert_to_pdf(SRC, out2a))
dt1 = time.time() - t0
print(f"    Call 1 ({dt1:.1f}s): {r1}", flush=True)
test("First call succeeded", os.path.exists(out2a))

t0 = time.time()
r2 = asyncio.run(convert_to_pdf(SRC, out2b))
dt2 = time.time() - t0
print(f"    Call 2 ({dt2:.1f}s): {r2}", flush=True)
test("Second call succeeded (no lock contention)", os.path.exists(out2b))

for f in [out2a, out2b]:
    if os.path.exists(f): os.remove(f)

# Test 3: No leftover temp dirs
print("\n[3] Cleanup check", flush=True)
import glob
lo_dirs = glob.glob('/tmp/lo_mcp_*')
test("No leftover temp dirs", len(lo_dirs) == 0, f"Found: {lo_dirs}")

# Summary
print(f"\n{'='*50}", flush=True)
print(f"Results: {PASS}/{PASS+FAIL} passed, {FAIL}/{PASS+FAIL} failed", flush=True)
if FAIL == 0: print("🎉 All passed!", flush=True)
else: print(f"⚠️ {FAIL} failed", flush=True); sys.exit(1)
