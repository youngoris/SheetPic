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


def get_latest_tag_version():
    """Get the highest existing version tag from git."""
    r = subprocess.run(
        ["git", "tag", "--list", "v*.*.*", "--sort=-v:refname"],
        capture_output=True, text=True
    )
    for line in r.stdout.strip().splitlines():
        tag = line.strip()
        ver = tag.lstrip("v")
        parts = ver.split(".")
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            return ver
    return None


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

    file_ver = get_current_version()
    tag_ver = get_latest_tag_version()

    # Bump from whichever is higher: file version or latest git tag
    if tag_ver and _ver_tuple(tag_ver) >= _ver_tuple(file_ver):
        base = tag_ver
    else:
        base = file_ver

    new = bump_patch(base)
    tag = f"v{new}"

    # Check if tag already exists (shouldn't, but guard)
    r = subprocess.run(["git", "rev-parse", tag], capture_output=True)
    if r.returncode == 0:
        print(f"Tag {tag} already exists, skip bump")
        return

    if new != file_ver:
        update_file(new)
    print(f"Version: {base} -> {new}")

    git("add", SOURCE_FILE)
    git("commit", "-m", f"Bump to {tag} [skip ci]")
    git("push")

    git("tag", tag)
    git("push", "origin", tag)
    print(f"Pushed tag {tag}")


def _ver_tuple(v):
    return tuple(int(x) for x in v.split("."))


if __name__ == "__main__":
    main()
