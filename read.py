import pandas as pd
from pprint import pprint


def read_sheet(filename, sheet):
    path = 'C:\\Users\\Jelte\\Desktop\\soldeerdingen\\spreadsheets\\{}.xlsx'.format(
        filename)

    xls = pd.ExcelFile(path)
    name = xls.sheet_names[sheet]

    resistors = pd.read_excel(
        path,
        sheet_name=sheet,
        header=1,
        index_col=0,
        usecols='A:B',
        nrows=32).dropna()

    ceramics = pd.read_excel(
        path,
        sheet_name=sheet,
        header=1,
        index_col=0,
        usecols='F:G',
        nrows=8
    ).dropna()

    films = pd.read_excel(
        path,
        sheet_name=sheet,
        header=10,
        index_col=0,
        usecols='F:G',
        nrows=16
    ).dropna()

    electros = pd.read_excel(
        path,
        sheet_name=sheet,
        header=27,
        index_col=0,
        usecols='F:G',
        nrows=6
    ).dropna()

    miscs = pd.read_excel(
        path,
        sheet_name=sheet,
        header=1,
        index_col=0,
        usecols='K:L',
        nrows=32).dropna()

    pots = pd.read_excel(
        path,
        sheet_name=sheet,
        header=1,
        index_col=0,
        usecols='P:Q',
        nrows=10).dropna()

    hws = pd.read_excel(
        path,
        sheet_name=sheet,
        header=14,
        index_col=0,
        usecols='P:Q',
        nrows=19).dropna()

    return {
        'pedal_name': name,
        'resistor': resistors,
        'ceramic': ceramics,
        'film': films,
        'electro': electros,
        'misc': miscs,
        'pot': pots,
        'hw': hws
    }


if __name__ == '__main__':
    cats = ['resistor', 'ceramic', 'film', 'electro', 'misc', 'pot', 'hw']
    sheets = range(0, 7)
    filename = 'fuzzdog'

    totals = {}

    for c in cats:
        for s in sheets:
            d = read_sheet(filename, s)[c]
            if c in totals.keys():
                totals[c] = totals[c].add(d, fill_value=0)
            else:
                totals[c] = d

    writer = pd.ExcelWriter(
        'C:\\Users\\Jelte\\Desktop\\soldeerdingen\\spreadsheets\\{}_resulsts.xlsx'.format(filename))

    for k in totals.keys():
        totals[k].to_excel(writer, k)

    writer.save()
