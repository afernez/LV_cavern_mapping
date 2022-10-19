# Check the cavern mapping for typos (that can't be easily handled), then (once
# user fixes typos) check the surface mapping vs the cavern mapping (and also
# against any other references, eg. the sheet used to test the produced PEPI
# cables) PPP info. Output typos and cavern differences to files in fixme
# folder. Also produce new cavern sheet(s) with fixes implemented in
# unformatted_fixed_cavern folder.
# The surface mapping is being used for the nominal PPP mapping for now, but as
# the mapping may change (depending on what is easiest to fix the cavern
# cables), the script will try to be flexible to using other input files
# (hopefully with minimal fixes required; a function will have to be added to
# parse any new input file if it has a different format, though).
# Note: LVR info is ignored throughout. To check cavern LVR info is correct
# in current mapping, need some file that associates lines with LVR directly
# from Phoebe's mapping...

import pandas, os, fnmatch
from argparse import ArgumentParser
from parseXls import * # steal what Mark did to parse the cavern file

### grab command line args
parser = ArgumentParser(description='Produce computer-readable cavern mappings')
# designate nominal and cavern mappings
parser.add_argument('mapping', help='specify cavern mapping to be used as input')
args = parser.parse_args()
# note that only C side needs to be compared; A side mapping is equivalent
nominal = 'surface_LV_power_tests_PMH_Formatting_wflex_flat_C_side.xlsx'
cavern = 'LVR_PPP_Underground_Mapping_PPPSorted_Samtec_cables__01-04-22.xlsx'
# also set sheets that will be compared for consistency
compare = []
if (args.doCompare).lower()=='true':
    compare = os.listdir('compare')
    compare = ['compare/'+file for file in compare]

# basic class that will uniquely identify (C-side) power lines, along with
# PPP connector and pin; LVR info isn't included for now
class line:

    def __init__(self, bp_con, ibbp2b2, flex, load, msa, ppp, ppp_pin):
        self.bp_con = bp_con
        self.ibbp2b2 = ibbp2b2
        self.flex = flex
        self.load = load
        self.msa = msa
        self.ppp = ppp
        self.ppp_pin = ppp_pin

    def __eq__(self, other):
        return (self.bp_con==other.bp_con and self.ibbp2b2==other.ibbp2b2 and
                self.flex==other.flex and self.load==other.load and
                self.msa==other.msa and self.ppp==other.ppp and
                self.ppp_pin==other.ppp_pin)

####### Helpers


####### Main Checking Functions

def cavern_typo_check():
    return True

def cavern_check_fix(do_compare):
    return True

print('\n\n')
no_typos = cavern_typo_check()
if no_typos: cavern_check_fix(do_compare)
