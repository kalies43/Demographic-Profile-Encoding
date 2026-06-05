# Demographic-Profile-Encoding
Enrollment Demographic Tally App
A lightweight Python desktop application built with Tkinter for the College of Our Lady of Mercy of Pulilan Foundation Inc. to streamline the encoding of student demographic profile questionnaires during enrollment.
Features

Encodes student data (program, religion, residency, previous school, student type) through a simple form interface
Directly tallies enrollment counts into the monthly Excel summary sheet (Demographic-Profile-SUMMARY.xlsm) without breaking its print layout
Tracks previous schools and municipalities in dedicated tally sheets (School Tally, Residency Tally) with searchable, auto-updating dropdowns
Supports per-sheet configuration for monthly layout changes
Runs entirely offline — no internet connection required

Requirements
pip install openpyxl
Usage
Place tally_app.py in the same folder as Demographic-Profile-SUMMARY.xlsm and run:
python tally_app.py
Tech Stack
Python · Tkinter · openpyxl
