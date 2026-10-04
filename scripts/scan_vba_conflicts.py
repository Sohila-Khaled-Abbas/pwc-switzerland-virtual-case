import os
import glob
import re

vba_files = glob.glob(r'vba\*.bas')
module_procs = {}
module_consts = {}

for f in vba_files:
    mname = os.path.basename(f)
    procs = []
    consts = []
    with open(f, 'r', encoding='utf-8', errors='ignore') as fp:
        for line in fp:
            line_s = line.strip()
            # check procs
            m_p = re.match(r'^(Public\s+)?(Sub|Function)\s+([a-zA-Z0-9_]+)', line_s, re.I)
            if m_p:
                procs.append(m_p.group(3))
            # check consts
            m_c = re.match(r'^Public\s+Const\s+([a-zA-Z0-9_]+)', line_s, re.I)
            if m_c:
                consts.append(m_c.group(1))
    module_procs[mname] = procs
    module_consts[mname] = consts

print('=== PROCEDURE DUPLICATES ACROSS MODULES ===')
all_procs = {}
for m, procs in module_procs.items():
    for p in procs:
        all_procs.setdefault(p.lower(), []).append((m, p))

dup_count = 0
for p_lower, sources in sorted(all_procs.items()):
    if len(sources) > 1:
        dup_count += 1
        print(f'Procedure "{sources[0][1]}" in: {[s[0] for s in sources]}')
print(f'Total duplicate procedures: {dup_count}')

print('\n=== CONSTANT DUPLICATES ACROSS MODULES ===')
all_consts = {}
for m, consts in module_consts.items():
    for c in consts:
        all_consts.setdefault(c.lower(), []).append((m, c))

dup_const_count = 0
for c_lower, sources in sorted(all_consts.items()):
    if len(sources) > 1:
        dup_const_count += 1
        print(f'Constant "{sources[0][1]}" in: {[s[0] for s in sources]}')
print(f'Total duplicate constants: {dup_const_count}')
