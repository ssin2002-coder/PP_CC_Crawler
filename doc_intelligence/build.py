"""
build.py — Frontend Vite build for doc_intelligence.

# Required commands at runtime:
#   python doc_intelligence/build.py
#   (or from project root: python -m doc_intelligence.build)
#
# Prerequisites (one-time setup):
#   - Node.js >= 18 with npm
#   - From repo root: `npm install --prefix doc_intelligence/web/frontend`
#     (this script will also run install if node_modules is missing)
#
# Output:
#   Bundles emitted to doc_intelligence/web/static/ (consumed by run.py).
"""
import os
import subprocess
import sys


def main() -> int:
    here = os.path.dirname(os.path.abspath(__file__))
    frontend_dir = os.path.join(here, "web", "frontend")
    if not os.path.isdir(frontend_dir):
        print(f"[build] frontend directory missing: {frontend_dir}", file=sys.stderr)
        return 1

    node_modules = os.path.join(frontend_dir, "node_modules")
    npm_cmd = "npm.cmd" if os.name == "nt" else "npm"

    if not os.path.isdir(node_modules):
        print(f"[build] installing dependencies in {frontend_dir}")
        try:
            subprocess.check_call(
                [npm_cmd, "install", "--no-audit", "--no-fund"],
                cwd=frontend_dir,
            )
        except subprocess.CalledProcessError as exc:
            print(f"[build] npm install failed: {exc}", file=sys.stderr)
            raise

    print(f"[build] running vite build in {frontend_dir}")
    try:
        subprocess.check_call([npm_cmd, "run", "build"], cwd=frontend_dir)
    except subprocess.CalledProcessError as exc:
        print(f"[build] npm run build failed: {exc}", file=sys.stderr)
        raise
    print("[build] done")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
