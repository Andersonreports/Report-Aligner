"""
Assert zero occurrences of known real values remain in a de-identified docx.
Scans every <w:t> text node in every XML part of the zip (document, headers,
footers, footnotes, endnotes) - not just python-docx's paragraph/table
traversal - to catch hyperlink-wrapped runs or anything structurally unusual.

Usage: python verify_deidentified.py <file.docx>
"""
import sys
import zipfile
import re

MUST_NOT_CONTAIN = [
    "Kerina",
    "AVP0000008750",
    "Sri Baalaji",
    "Villupuram",
    "260604571",
    "22/05/2026",
    "23/05/2026",
    "16/06/2026",
    "SH3TC2",
    "c.2710C>T",
    "Arg904Ter",
    "c.3287_3308del",
    "Asn1096MetfsTer13",
    "Charcot-Marie-Tooth",
    "601596",
    "608206",
    "149027022",
    "149010289",
    "NC_000005",
    "NM_024577",
    "21696",
    "29515423",
    "37.72%",
    "36.52%",
]


def main():
    path = sys.argv[1]
    z = zipfile.ZipFile(path)
    all_text = []
    for name in z.namelist():
        if not name.endswith('.xml'):
            continue
        xml = z.read(name).decode('utf-8', errors='ignore')
        all_text.extend(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml))
    blob = ''.join(all_text)

    failures = [needle for needle in MUST_NOT_CONTAIN if needle in blob]
    if failures:
        print("FAILED - real values still present:", failures)
        sys.exit(1)
    print("PASSED - no known real values found in", path)


if __name__ == "__main__":
    main()
