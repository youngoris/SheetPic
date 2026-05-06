"""
Auto-increment patch version in APP_VERSION, commit, and create tag.
Usage: python scripts/auto_bump.py
"""
import re
import subprocess
import sys
import os

SOURCE_FILE = "sheetpic.py"
VERSION_PATTERN = r'(APP_VERSION\s*=\s*")(\d+\.\d+\.\d+)(")'


def get_current_version():
    with open(SOURCE_FILE, "r") as f:
        m = re.search(VERSION_PATTERN, f.read())
    if not m:
        print(f"ERROR: APP_VERSION not found in {SOURCE_FILE}")
        sys.exit(1)
    return m.group(2)


def bump_patch(version):
    major, minor, patch = version.split(".")
    return f"{major}.{minor}.{int(patch) + 1}"


def update_file(new_version):
    with open(SOURCE_FILE, "r") as f:
        content = f.read()
    new_content = re.sub(VERSION_PATTERN, f'\\g<1>{new_version}\\3', content)
    with open(SOURCE_FILE, "w") as f:
        f.write(new_content)


def git(*args):
    r = subprocess.run(["git"] + list(args), capture_output=True, text=True)
    if r.returncode != 0:
        print(f"git {' '.join(args)} failed: {r.stderr}")
        sys.exit(1)
    return r.stdout.strip()


def main():
    os.chdir(os.path.join(os.path.dirname(__file__), ".."))

    old = get_current_version()
    new = bump_patch(old)
    tag = f"v{new}"

    # Check if tag already exists
    r = subprocess.run(["git", "rev-parse", tag], capture_output=True)
    if r.returncode == 0:
        print(f"Tag {tag} already exists, skip bump")
        return

    update_file(new)
    print(f"Version: {old} -> {new}")

    git("add", SOURCE_FILE)
    git("commit", "-m", f"Bump to {tag} [skip ci]")
    git("push")

    git("tag", tag)
    git("push", "origin", tag)
    print(f"Pushed tag {tag}")


if __name__ == "__main__":
    main()
