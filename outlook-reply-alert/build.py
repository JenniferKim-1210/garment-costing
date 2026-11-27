"""Build <Module>.bas (ASCII only, CRLF) from every src/<Module>.src.bas (UTF-8).

The VBA editor reads .bas files in the PC's legacy code page, so Korean text in the
file breaks on PCs whose "language for non-Unicode programs" is not Korean.
This script rewrites every non-ASCII string literal as U("\\uXXXX...") (decoded at run
time) and drops non-ASCII text from comments, so the module works on any Windows.

    python build.py
"""
import pathlib
import re

HERE = pathlib.Path(__file__).parent


def split_comment(line):
    """Return (code, comment) where comment starts at the first ' outside a string."""
    in_str = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i], line[i:]
    return line, ""


def escape(text):
    return "".join(ch if ord(ch) < 128 else "\\u%04X" % ord(ch) for ch in text)


def convert_code(code):
    def repl(m):
        body = m.group(1)
        if all(ord(c) < 128 for c in body):
            return m.group(0)
        return 'U("%s")' % escape(body)
    return re.sub(r'"((?:[^"]|"")*)"', repl, code)


def convert_comment(comment):
    if all(ord(c) < 128 for c in comment):
        return comment
    kept = "".join(c for c in comment if ord(c) < 128)
    kept = re.sub(r"\s+", " ", kept).rstrip()
    # keep only if something meaningful (letters/digits) is left
    return kept if re.search(r"[A-Za-z0-9]", kept[1:]) else ""


def build(src, dst):
    out = []
    for n, line in enumerate(src.read_text(encoding="utf-8").splitlines(), 1):
        code, comment = split_comment(line)
        new_code = convert_code(code)
        if new_code != code and re.match(r"\s*(Public |Private )?Const\b", code):
            raise SystemExit("%s line %d: Korean text in a Const (use a variable)" % (src.name, n))
        new_comment = convert_comment(comment)
        if comment and not new_comment and not code.strip():
            continue  # comment-only line that was entirely Korean
        new_line = (new_code + new_comment).rstrip() if new_comment else new_code.rstrip()
        assert all(ord(c) < 128 for c in new_line), (n, new_line)
        if len(new_line) > 1000:
            raise SystemExit("%s line %d is too long for the VBA editor" % (src.name, n))
        out.append(new_line)
    dst.write_bytes(("\r\n".join(out) + "\r\n").encode("ascii"))
    # paste-ready copy (no Attribute line) for PCs where File > Import does not work
    paste = dst.with_name(dst.stem + "_paste.txt")
    paste.write_bytes(("\r\n".join(l for l in out if not l.startswith("Attribute VB_Name")) + "\r\n").encode("ascii"))
    print("wrote", dst.name, "and", paste.name, len(out), "lines")


def main():
    for src in sorted((HERE / "src").glob("*.src.bas")):
        build(src, HERE / (src.name[: -len(".src.bas")] + ".bas"))


if __name__ == "__main__":
    main()
