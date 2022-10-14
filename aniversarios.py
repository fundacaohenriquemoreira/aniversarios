#-*- coding: utf-8 -*-
# aniversarios.py  (c)2021  Henrique Moreira

"""
aniversarios - leitor de aniversarios.
Basic Libre/Excel reader (using filing.xcelent 'openpyxl' wrapper).
"""

# pylint: disable=no-self-use, missing-function-docstring

import sys
import os
from os import environ
import datetime
import openpyxl
import waxpage.redit as redit
import filing.xcelent as xcelent


def main():
    """ Main function! """
    run(sys.stdout, sys.stderr, sys.argv[1:])

def run(out, err, args):
    fname = what_aniv()["aniversarios"]
    #print(f"# fname: '{fname}'")
    param = args
    if not param:
        param = [fname]
    for name in param:
        dump_aniv(out, err, name)
    return 0

def dump_aniv(out, err, fname:str) -> int:
    print("# reading:", fname)
    names = list()
    wbk = openpyxl.load_workbook(fname)
    libre = xcelent.Xcel(wbk)
    sheet = libre.get_sheet(1)
    table = xcelent.Xsheet(sheet)
    idx = 0
    for row in table.rows:
        idx += 1
        shown = [simpler(item.value) if item.value else "-" for item in row]
        #addup = [item.value for item in row if item.value]
        if not shown:
            continue
        if len(shown[0]) <= 1:
            continue
        name, dash, dtstr = shown[0], shown[1], shown[2]
        assert name
        assert dash == "-", f"Invalid null cell: {shown}"
        if dtstr == "-":
            continue
        if dtstr.endswith("-") and 1 <= dtstr.count("-") <= 2:
            day, month = dtstr.rstrip("-").split("-")
        elif dtstr.count("-") == 2:
            day, month, _ = dtstr.split("-")
        else:
            day, month = "0", "0"
        day, month = int(day), int(month)
        astr = f"{month:02}.{day:02} {name:.<20} {dtstr}"
        names.append(astr)
    for astr in sorted(names):
        out.write(astr + "\n")
    return 0

def simpler(astr, default="") -> str:
    if not astr:
        return default
    if isinstance(astr, (datetime.date, datetime.datetime)):
        new = f"{astr.day:02}-{astr.month:02}-{astr.year:04}"
        return new
    return redit.char_map.simpler_ascii(astr)

def what_aniv() -> dict:
    """ Returns dictionary from ~/.config/misc.conf
    """
    res = {}
    home = environ["USERPROFILE"] if environ.get("HOME") is None else environ.get("HOME")
    path = os.path.join(home, ".config", "misc.conf")
    with open(path, "r", encoding="ascii") as fdin:
        lines = [line.rstrip() for line in fdin.readlines() if line.strip() and line[0] != "#"]
        for line in lines:
            tup = line.split("=", maxsplit=1)
            if len(tup) <= 1:
                print("# Ignored:", line)
                continue
            left, right = tup[0].strip(), tup[1].strip()
            res[left] = right
    return res

# Main script
if __name__ == "__main__":
    main()
