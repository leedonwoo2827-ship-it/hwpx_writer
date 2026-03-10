#!/usr/bin/env python3
"""
HWPX 네임스페이스 후처리 유틸리티

python-hwpx가 생성한 HWPX 파일의 XML 네임스페이스 프리픽스를
한컴오피스 표준 프리픽스로 교체한다.

사용법:
  CLI:    python fix_namespaces.py <file.hwpx>
  Import: fix_hwpx_namespaces("output.hwpx")
"""

import zipfile
import os
import re
import sys


def fix_hwpx_namespaces(hwpx_path):
    """
    HWPX 파일의 ns0:/ns1: 등 자동 생성 프리픽스를
    한컴오피스 표준 프리픽스(hh/hc/hp/hs)로 교체한다.
    """
    NS_MAP = {
        "http://www.hancom.co.kr/hwpml/2011/head": "hh",
        "http://www.hancom.co.kr/hwpml/2011/core": "hc",
        "http://www.hancom.co.kr/hwpml/2011/paragraph": "hp",
        "http://www.hancom.co.kr/hwpml/2011/section": "hs",
    }

    tmp_path = hwpx_path + ".tmp"

    with zipfile.ZipFile(hwpx_path, "r") as zin:
        items = zin.infolist()

        mimetype_items = [i for i in items if i.filename == "mimetype"]
        other_items    = [i for i in items if i.filename != "mimetype"]
        ordered_items  = mimetype_items + other_items

        with zipfile.ZipFile(tmp_path, "w") as zout:
            for item in ordered_items:
                data = zin.read(item.filename)

                if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                    text = data.decode("utf-8")

                    ns_aliases = {}
                    for match in re.finditer(r'xmlns:(ns\d+)="([^"]+)"', text):
                        alias, uri = match.group(1), match.group(2)
                        if uri in NS_MAP:
                            ns_aliases[alias] = NS_MAP[uri]

                    for old_prefix, new_prefix in ns_aliases.items():
                        text = text.replace(f"xmlns:{old_prefix}=", f"xmlns:{new_prefix}=")
                        text = text.replace(f"<{old_prefix}:", f"<{new_prefix}:")
                        text = text.replace(f"</{old_prefix}:", f"</{new_prefix}:")

                    data = text.encode("utf-8")

                if item.filename == "mimetype":
                    info = zipfile.ZipInfo("mimetype")
                    info.compress_type = zipfile.ZIP_STORED
                    zout.writestr(info, data)
                else:
                    zout.writestr(item, data, compress_type=zipfile.ZIP_DEFLATED)

    os.replace(tmp_path, hwpx_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python fix_namespaces.py <file.hwpx>", file=sys.stderr)
        print("  Fixes namespace prefixes for Hangul Viewer compatibility.", file=sys.stderr)
        sys.exit(1)

    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"Error: File not found: {path}", file=sys.stderr)
        sys.exit(1)

    fix_hwpx_namespaces(path)
    print(f"Fixed namespaces: {path}", file=sys.stderr)
