import os
import numpy as np
from collections import defaultdict
from xlsxwriter import Workbook

prefix = 'mch-'
species_of_interest = 'mch'
consumption_desired = 20

moles_file = '{}Moles.csv'.format(prefix)
rop_file = '{}ROP.csv'.format(prefix)
percent_file = '{}Percent2.xlsx'.format(prefix)

moles = np.loadtxt(moles_file, dtype=float, delimiter=',', skiprows=1)
rop = np.loadtxt(rop_file, dtype=float, delimiter=',', skiprows=1)

with open(moles_file, 'r') as m:
    moles_header_line = m.readline()

with open(rop_file, 'r') as r:
    rop_header_line = r.readline()

moles_header = [x.strip() for x in moles_header_line.split(',')]
rop_header = [x.strip() for x in rop_header_line.split(',')]

fuel_conversion_index = moles_header.index('Molar_conversion_{}_(percent)'.format(species_of_interest))
fuel_conversion = moles[:, fuel_conversion_index]
consumption_index = next(i for i,x in enumerate(fuel_conversion) if x >= consumption_desired)

time = moles[0:consumption_index, 0]
integ = np.trapz(rop[0:consumption_index,:], time, axis=0)

species = []
reactions = {}
useful_columns = {}

for head in moles_header:
    if 'Mole_fraction' in head:
        spec = head.split('_')[2]
        if spec not in species:
            species.append(spec)

for i, head in enumerate(rop_header):
    elems = head.split('_')
    # If len(elems) < 2, its the first column and we don't want it
    if len(elems) > 2:
        if 'GasRxn#' in elems[3]:
            rxn_num = int(elems[3].strip('GasRxn#'))-1
            spec_num = species.index(elems[0])
            reaction = elems[2]
            useful_columns[i] = (spec_num,rxn_num)
            if rxn_num not in reactions:
                reactions[rxn_num] = reaction

num_reactions = max(reactions.keys()) + 1
percents = np.zeros((num_reactions, len(species)))

crea = np.zeros((len(species)))
dest = np.zeros((len(species)))

for col, (spec_num, rxn_num) in useful_columns.items():
    if integ[col] >= 0:
        crea[spec_num] += integ[col]
    else:
        dest[spec_num] += abs(integ[col])

for col, (spec_num, rxn_num) in useful_columns.items():
    if integ[col] >= 0:
        percents[rxn_num, spec_num] = integ[col]/crea[spec_num]*100
    else:
        percents[rxn_num, spec_num] = integ[col]/dest[spec_num]*100

wb = Workbook(percent_file)
ws = wb.add_worksheet()

for col, spec in enumerate(species):
    col += 2
    ws.write(0, col, spec)

for rxn_num, reaction in reactions.items():
    row = rxn_num + 1
    ws.write(row, 0, row)
    ws.write(row, 1, reaction)

it = np.nditer(percents, flags=['multi_index'])

while not it.finished:
    ws.write(it.multi_index[0] + 1, it.multi_index[1] + 2, it[0])
    it.iternext()

wb.close()
