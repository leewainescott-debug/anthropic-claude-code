set -e
# design-A actuals on the 1.x tabs, then the owner's 2707 adoptions that need the built
# blocks in place: the Actuals column on the Portfolio Summary tables, his live levers,
# the 0.2 COE spend repoints and the 1.13 bar
python3 actuals.py cand.xlsx pA.xlsx A
python3 post2707.py pA.xlsx pB.xlsx
# the design sweep again, idempotently: actuals and post2707 create the actuals table at
# the top of each 1.x tab, the K/L pair down the squad tables and the G Actuals cells,
# all after chain2's design pass has run, so the same label-anchored sweep runs once more
# to dress what they built
python3 design2707.py pB.xlsx pC.xlsx
# his words, not mine: "ledger" was never his term for the role mapping
python3 plainwords.py pC.xlsx pD.xlsx
python3 -c "
import wbio
rc,st=wbio.build('pD.xlsx','cand_A.xlsx'); print('  A injected',st)
e,b=wbio.audit('cand_A.xlsx'); print('  A errors',len(e),'blank',len(b))
for x in e[:10]: print('    ',x)
"

# the mechanical audit: figures against the ledger, controls, modelling standards
python3 audit.py cand_A.xlsx || true

# what of his own typing is not in the file - his question, answered every build
python3 whatsgone.py base_2707.xlsx cand_A.xlsx --later rev.xlsx || true
python3 whatsgone.py rev.xlsx cand_A.xlsx || true
