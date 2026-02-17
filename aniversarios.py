#!/usr/bin/python3
# -*- coding: utf-8 -*-
#
# (c) 2021..2026  Henrique Moreira

""" aniversarios.py  (c) 2021..2025  Henrique Moreira

aniversarios - leitor de aniversarios.
Basic Libre/Excel reader (using filing.xcelent 'openpyxl' wrapper).
"""

# pylint: disable=missing-function-docstring

import sys
import os
from os import environ
import datetime
import openpyxl
import filing.xcelent
from waxpage.redit import char_map

DEBUG = 0

EXCL_EXPR = {
    "Isabel": "RIP",
}


def main():
    """ Main function! """
    run(sys.stdout, sys.stderr, sys.argv[1:])

def run(out, err, args):
    """ Main script run """
    dct = what_aniv()
    fname = dct["aniversarios"]
    param = args
    if not param:
        param = [fname]
    for name in param:
        print("# reading:", name, "(no misc.conf)" if dct["conf-file"] is None else "(.config/misc.conf)")
        dump_aniv(out, err, name, "aniversarios.txt")
    return 0

def dump_aniv(out, err, fname:str, outname:str="") -> int:
    debug = DEBUG
    names = []
    wbk = openpyxl.load_workbook(fname)
    libre = filing.xcelent.Xcel(wbk)
    sheet = libre.get_sheet(1)
    table = filing.xcelent.Xsheet(sheet)
    for row in table.rows:
        shown = [
            simpler(item.value) if item.value else "-" for item in row
        ]
        if not shown:
            continue
        if len(shown[0]) <= 1:
            continue
        name, dash, dtstr = shown[0], shown[1], shown[2]
        assert name
        if dash == "#":
            continue
        assert dash == "-", f"Invalid null cell: {shown}"
        if dtstr == "-":
            continue
        if dtstr.endswith("-") and 1 <= dtstr.count("-") <= 2:
            day, month, year = dtstr.rstrip("-").split("-"), ""
        elif dtstr.count("-") == 2:
            day, month, year = dtstr.split("-")
        else:
            day, month = "0", "0"
        day, month = int(day), int(month)
        astr = f"{month:02}.{day:02} {name:.<20} {dtstr}\n"
        names.append((astr, int(year) if year else 0))
    ours = ""
    for astr, year in sorted(names):
        out.write(astr)
        line = astr
        line = ("+" if year <= 1974 else "-") + "  " + astr
        if excluded(line, EXCL_EXPR, debug=debug):
            continue
        ours += line
    print("# Output:", [outname])
    if not outname:
        return 0
    with open(outname, "wb") as fdout:
        fdout.write(bytes(ours.encode("ascii")))
    return 0

def simpler(astr, default="") -> str:
    if not astr:
        return default
    if isinstance(astr, (datetime.date, datetime.datetime)):
        new = f"{astr.day:02}-{astr.month:02}-{astr.year:04}"
        return new
    return char_map.simpler_ascii(astr)


def what_aniv() -> dict:
    """ Returns dictionary from ~/.config/misc.conf
    """
    home = environ["USERPROFILE"] if environ.get("HOME") is None else environ.get("HOME")
    path = os.path.join(home, ".config", "misc.conf")
    res = {
        "conf-file": None,
        "aniversarios": os.path.join(home, "aniversarios.xlsx"),
    }
    if os.path.isfile(path):
        res = dict_from_conf_file(path)
    return res

def dict_from_conf_file(path):
    """ Returns the dictionary from tuples LValue = RValue """
    res = {
        "conf-file": path,
    }
    with open(path, "r", encoding="ascii") as fdin:
        lines = [
            line.rstrip() for line in fdin.readlines()
            if line.strip() and line[0] != "#"
        ]
        for line in lines:
            tup = line.split("=", maxsplit=1)
            if len(tup) <= 1:
                print("# Ignored:", line)
                continue
            left, right = tup[0].strip(), tup[1].strip()
            res[left] = right
    return res


def excluded(astr:str, subexprs, debug=0) -> bool:
    for key in subexprs:
        if " " + key in astr:
            if debug > 0:
                print(f"Excluded {[subexprs[key]]}:", key, [astr])
            return True
    return False


# Main script
if __name__ == "__main__":
    main()
