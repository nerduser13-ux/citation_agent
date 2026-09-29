"""report.py

Generates the verification / review CSV required by the workflow.
"""
import csv


class Report:
    COLUMNS = [
        "Citation Number", "Location", "Citation Text", "Reference Text",
        "URL", "Action", "Status", "Notes",
    ]

    def __init__(self):
        self.rows = []
        self.summary = {}

    def add(self, citation_number, location, citation_text, reference_text,
            url, action, status, notes):
        self.rows.append([
            citation_number, location, citation_text, reference_text,
            url, action, status, notes,
        ])

    def set_summary(self, **kw):
        self.summary.update(kw)

    def write(self, path):
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(self.COLUMNS)
            for r in self.rows:
                w.writerow(r)
            w.writerow([])
            w.writerow(["SUMMARY"])
            for k, v in self.summary.items():
                w.writerow([k, v])
