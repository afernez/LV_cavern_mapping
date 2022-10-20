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

import pandas, os, fnmatch, math
import csv
from argparse import ArgumentParser

### grab command line args
parser = ArgumentParser(description='Check and fix cavern mapping')
# designate nominal and cavern mappings
parser.add_argument('nominal', help='specify nominal mapping to be used as input')
parser.add_argument('cavern', help='specify cavern mapping to be used as input')
# note that only C side needs to be compared; A side mapping is equivalent
# also set sheets that will be compared for consistency
parser.add_argument('doCompare', help='specify if compare mappings should be checked')
args = parser.parse_args()
nominal = args.nominal
cavern = args.cavern
compare = []
if (args.doCompare).lower()=='true':
    compare = os.listdir('compare')
    compare = ['compare/'+file for file in compare]

# from parseXls import * # steal what Mark did to parse the cavern file
# actually, don't import here; messes up argparse, and I'm too lazy to debug it

# basic class that will uniquely identify (with redundancy) power lines, along
# with PPP connector and pin; LVR info isn't included for now
class line:

    def __init__(self, x, y, z, bp, bp_con, ibbp2b2, flex, load, msa, ppp,
                 ppp_pin):
        self.x = x
        self.y = y
        self.z = z
        self.bp = bp
        self.bp_con = bp_con
        self.ibbp2b2 = ibbp2b2
        self.flex = flex
        self.load = load
        self.msa = msa
        self.ppp = ppp
        self.ppp_pin = ppp_pin

    def equal_minus_ppp(self, other):
        return (self.x==other.x and self.y==other.y and self.z==other.z and
                self.bp==other.bp and self.bp_con==other.bp_con and
                self.ibbp2b2==other.ibbp2b2 and self.flex==other.flex and
                self.load==other.load and self.msa==other.msa)

    def __eq__(self, other):
        return (self.x==other.x and self.y==other.y and self.z==other.z and
                self.bp==other.bp and self.bp_con==other.bp_con and
                self.ibbp2b2==other.ibbp2b2 and self.flex==other.flex and
                self.load==other.load and self.msa==other.msa and
                self.ppp==other.ppp and self.ppp_pin==other.ppp_pin)

    def __hash__(self):
        return hash(self.x+self.y+self.z+self.bp+self.bp_con+self.ibbp2b2+
                    self.flex+self.load+self.msa+self.ppp+self.ppp_pin)


####### Helpers

# just going to copy & paste to steal Mark's parsing functions...
# ideally would be importing these... TODO

#
# This part is common to the hybrids and DCBs
#
def parseSheet( xls, sheet ):
    dfIn = pandas.read_excel(xls, sheet , usecols="C:G", skiprows=[0,1] )
    cols = list(dfIn.columns)
    cols[-1] = 'LVR'
    dfIn.columns = cols
    # get rid of empty rows
    dfIn = dfIn.dropna(subset=['LVR'])
    # make separate columns with LVR ID plus connector, pin & channel for LVR and PPP connectors
    dfIn['LVR ID - Connector - Pin'] = dfIn['LVR ID - Connector - Pin'].str.replace(' ', '').str.replace( '2.5V:','')
    dfIn[['LVR ID','LVR Connector', 'LVR Pin']] = dfIn['LVR ID - Connector - Pin'].str.split('-', expand=True)
    dfIn['LVR ID - Connector - Pin'] = dfIn['LVR ID - Connector - Pin'].str.replace('-',' - ')
    dfIn['LVR Channel'] = (dfIn['LVR Connector'].str.replace('J','').astype(int) - 12)*4 + ( 8 - dfIn['LVR Pin'].str[0:1].astype(int) ) / 2 +1
    dfIn[['PPP Positronic', 'PPP Src/Ret']] = dfIn['PPP Connector - Pin'].str.replace(' ','').str.split('-', expand=True)
    return dfIn


#
# Read sheet for DCBs and add additional information to data frame
#
def parseDCBs( xls, sheet ):
    dfIn = parseSheet( xls, sheet )
    # set voltage
    dfIn['Voltage'] = '1V5'
    dfIn.loc[ dfIn['PPP Name'].str.contains( '2V5' ), 'Voltage' ] = "2V5"
    dfIn['M/S/A'] = 'A'
    dfIn.loc[ dfIn['PPP Name'].str.contains( 'Master' ), 'M/S/A' ] = "M"
    dfIn.loc[ dfIn['PPP Name'].str.contains( 'Slave' ), 'M/S/A' ] = "S"

    # set DCB and position
    ###dfIn['LVR Name'] = dfIn['LVR Name'].str.replace( 'alhpa', 'alpha' )
    aa = dfIn['LVR Name'].str.split('_', expand=True)
    aa.loc[ aa[6].isnull(), 'DCB' ] = aa[1]
    aa.loc[ ~aa[6].isnull(), 'DCB' ] = aa[2] + '&' + aa[1]

    aa['Pos'] = aa[1].astype(float)
    aa.loc[ aa['DCB'].str.contains('&'), 'Pos' ] += 1.5
    dfIn[['DCB','Pos']] = aa[['DCB','Pos']]
    dfIn = dfIn.sort_values( by=['Pos', 'PPP Name' ] )
    dfIn['BP Connector'] = 'JD' + dfIn['DCB'].str.replace('&','-')
    dfIn['iBB/P2B2 Connector'] = dfIn['PPP Name'].apply( lambda x: x[x.find( 'J', 1 ):x.find( '_', x.find( 'J', 1 ) )] )
    dfIn['SBC FLEX NAME'] = 'n/a'

    # delete columns with repeated/unnecessary information
    del dfIn['Pos']
    #del dfIn['LVR']
    return dfIn


#
# Read sheet for DCBs and add additional information to data frame
#
def parseHybrids( xls, sheet ):
    dfIn = parseSheet( xls, sheet )
    # get flex and 4-asic group
    aa = dfIn['LVR Name'].str.split("_",expand=True)
    dfIn['SBC FLEX NAME'] = aa[3]
    dfIn['4-asic group'] = aa[4]

    # set type - M/S for the 8-asic hybrids in alpha, otherwise A
    dfIn['M/S/A'] = 'A'
    # alpha, X0S and S0S
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '0S_P1W' ), 'M/S/A' ] = "M"
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '0S_P1E' ), 'M/S/A' ] = "S"
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '0S_P2E' ), 'M/S/A' ] = "M"
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '0S_P2W' ), 'M/S/A' ] = "S"
    # alpha X0M and SOM
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '0M_P1W' ), 'M/S/A' ] = "M"
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '0M_P1E' ), 'M/S/A' ] = "S"
    # alpha X1S and S1S
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '1S_P1W' ), 'M/S/A' ] = "M"
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '1S_P1E' ), 'M/S/A' ] = "S"
    # alpha X1M and S1M
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '1M_P1W' ), 'M/S/A' ] = "M"
    dfIn.loc[ dfIn['LVR Name'].str.contains( 'alpha' ) & dfIn['LVR Name'].str.contains( '1M_P1E' ), 'M/S/A' ] = "S"

    # Sort by backplane connection
    dfIn[ ['BP Connector', 'iBB/P2B2 Connector' ] ] = dfIn['PPP Name'].str.split('_', expand=True)[ [0, 1] ]
    dfIn['Pos'] = dfIn['BP Connector'].str.replace( 'JP', '' ).astype(int)
    dfIn = dfIn.sort_values( by=['Pos', '4-asic group' ] )#, 'LVR Name'] )

    # delete columns with duplicate/unnecessary information
    del dfIn['Pos']
    del dfIn['LVR']
    del dfIn['LVR ID - Connector - Pin']
    del dfIn['PPP Connector - Pin']
    return dfIn

# for C side
def z_truemir_to_y_z(z, truemir):
    z = z.lower()
    truemir = truemir.lower()
    if z=='mag' and truemir=='true': return ('top', z)
    if z=='mag' and truemir=='mirror': return ('bot', z)
    if z=='ip' and truemir=='true': return ('bot', z)
    if z=='ip' and truemir=='mirror': return ('top', z)
    print(f'Couldn\'t identify {z}, {truemir}...')
    return '??'

def true_mirror(x, y, z):
    if x=='C':
        if y=='top':
            if z=='mag': return 'True'
            if z=='ip': return 'Mirror'
            return '??'
        if y=='bot':
            if z=='mag': return 'Mirror'
            if z=='ip': return 'True'
            return '??'
        return '??'
    if x=='A':
        if y=='top':
            if z=='mag': return 'Mirror'
            if z=='ip': return 'True'
            return '??'
        if y=='bot':
            if z=='mag': return 'True'
            if z=='ip': return 'Mirror'
            return '??'
        return '??'
    return '??'

def ppp_ret_pin(src_pin):
    src = int(src_pin)
    ret = src+8
    return str(ret)


####### Functions to Parse different sheets

# returns a list of lines for surface mapping
def parse_surface(file):
    xlsx = pandas.ExcelFile(file)
    sheets = xlsx.sheet_names
    lines = []
    for sheet in sheets:
        sheet_info = sheet.split('-')
        x, y, z, bp = sheet_info[0], sheet_info[2], sheet_info[1], sheet_info[3]
        df = pandas.read_excel(xlsx, sheet, usecols='A:G')
        # get rid of empty rows
        df = df.dropna(how='all')
        for ind, row in df.iterrows():
            flex = row['SBC FLEX NAME']
            if isinstance(flex, float) and math.isnan(flex): flex='n/a'
            l=line(x, y, z, bp, row['BP Connector'],
                   row['iBB/P2B2 Connector'], flex,
                   row['4ASIC-group (hybrid)/DCB power'], row['M/S/A'],
                   row['PPP Positronic'], str(row['Positronic Src'])[0])
            lines.append(l)
    return lines


# returns a list of lines for cavern mapping
def parse_cavern(file):
    xlsx = pandas.ExcelFile(file)
    sheets = xlsx.sheet_names
    # separate dcb and hybrid sheets
    dcb_sheets = fnmatch.filter( sheets, "DCB - *" )
    hyb_sheets = fnmatch.filter( sheets, "Hybrid - *" )

    lines = []
    for dcb_sheet in dcb_sheets:
        x = 'C' # take advantage of only doing C-side...
        sheet_info = dcb_sheet.split(' - ')
        y, z = z_truemir_to_y_z(sheet_info[1], sheet_info[2])
        df = parseDCBs(xlsx, dcb_sheet)
        for ind, row in df.iterrows():
            lvr_name = (row['LVR Name']).split('_')
            bp = lvr_name[2]
            # bp_con = 'JD'+lvr_name[1]
            if '25' in row['LVR Name']:
                bp = lvr_name[3]
            #     bp_con = 'JD'+lvr_name[2]+'-'+lvr_name[1]
            # ppp_name = (row['PPP Name']).split('_')
            # ibbp2b2 = ppp_name[1]
            # load = ppp_name[2]
            # if '2V5' in row['PPP Name']:
            #     ibb_p2b2 = ppp_name[2]
            #     load = ppp_name[3]
            flex = 'n/a' # DCBs
            # msa = 'A'
            # if 'Master' in row['PPP Name']: msa = 'M'
            # if 'Slave' in row['PPP Name']: msa = 'S'
            # ppp_connector_pin = (row['PPP Connector - Pin']).split(' - ')
            # ppp = ppp_connector_pin[0]
            # ppp_pin = ppp_connector_pin[1][0]
            bp_con = row['BP Connector']
            ibbp2b2 = row['iBB/P2B2 Connector']
            load = row['Voltage']
            msa = row['M/S/A']
            ppp = row['PPP Positronic']
            ppp_pin = row['PPP Src/Ret'][0]
            l=line(x, y, z, bp, bp_con, ibbp2b2, flex, load, msa, ppp, ppp_pin)
            lines.append(l)

    for hyb_sheet in hyb_sheets:
        x = 'C' # take advantage of only doing C-side...
        sheet_info = hyb_sheet.split(' - ')
        y, z = z_truemir_to_y_z(sheet_info[1], sheet_info[2])
        df = parseHybrids(xlsx, hyb_sheet)
        for ind, row in df.iterrows():
            lvr_name = (row['LVR Name']).split('_')
            # ppp_name = (row['PPP Name']).split('_')
            bp = lvr_name[2]
            # bp_con = ppp_name[0]
            # ibbp2b2 = ppp_name[1]
            # flex = lvr_name[3]
            # load = lvr_name[4]
            bp_con = row['BP Connector']
            ibbp2b2 = row['iBB/P2B2 Connector']
            flex = row['SBC FLEX NAME']
            load = row['4-asic group']
            msa = row['M/S/A']
            ppp = row['PPP Positronic']
            ppp_pin = row['PPP Src/Ret'][0]
            l=line(x, y, z, bp, bp_con, ibbp2b2, flex, load, msa, ppp, ppp_pin)
            lines.append(l)

    return lines


# returns a list of lines for cable test mapping
def parse_cable_test(file):
    xlsx = pandas.ExcelFile(file)
    sheets = xlsx.sheet_names
    lines = []
    for sheet in sheets:
        continue
    return lines


### Associate files with a function for parsing
parse_func = {}
files = [nominal, cavern] + compare
for file in files:
    if 'surface_LV_power_tests' in file: parse_func[file]=parse_surface
    elif 'LVR_PPP_Underground' in file: parse_func[file]=parse_cavern
    elif 'lvr_testing' in file: parse_func[file]=parse_cable_test
    else: print(f'\n\nDON\'T RECONGNIZE {file} FORMAT\n\n')



####### Main Checking Functions

def cavern_typo_check():
    print(f'\n\nChecking {cavern} for typos...\n\n')
    nominal_lines = parse_func[nominal](nominal)
    cavern_lines = parse_func[cavern](cavern) # parse func will do some checks TODO
    # now, make sure you can find every line in cavern lines! (but, don't
    # require the PPP to be the same; this will be checked later)
    found_all_lines_ok = True
    for nom_line in nominal_lines:
        found_line = False
        for cav_line in cavern_lines:
            if cav_line.equal_minus_ppp(nom_line): found_line=True
        if not found_line:
            print(nom_line.x+nom_line.y+nom_line.z+nom_line.bp+
                  nom_line.bp_con+nom_line.ibbp2b2+nom_line.flex+
                  nom_line.load+nom_line.msa) # TODO text file in FIXME...
            found_all_lines_ok = False
    return found_all_lines_ok


def cavern_check_fix():
    print(f'\n\nChecking {nominal} vs {cavern}...\n\n')
    nominal_lines = parse_func[nominal](nominal)
    cavern_lines = parse_func[cavern](cavern)
    ppp_wrong_cavern_lines = {} # map from nom line to cav line
    for nom_line in nominal_lines:
        cav_line = None
        found = 0
        for cl in cavern_lines:
            if cl.equal_minus_ppp(nom_line):
                cav_line = cl
                found += 1
        if found==0 : print(f'Couldn\'t find '+
                      f'{nom_line.x+nom_line.y+nom_line.z+nom_line.bp}'+
                      f'{nom_line.bp_con+nom_line.ibbp2b2+nom_line.flex}'+
                      f'{nom_line.load+nom_line.msa}???')
        # finding more than 1 is ok if hybrid M/s (spliced at LVR, but cavern
        # mapping includes both lines coming out of LVR); if not west/east pair
        # printed, then you should be worried...
        if found > 1: print(f'Found more than 1 '+
                      f'{nom_line.x+nom_line.y+nom_line.z+nom_line.bp}'+
                      f'{nom_line.bp_con+nom_line.ibbp2b2+nom_line.flex}'+
                      f'{nom_line.load+nom_line.msa}???')
        if not cav_line==nom_line:
            print(f'\nFound cavern line with wrong PPP!\n')
            ppp_wrong_cavern_lines[nom_line] = cav_line

    # write out all the wrong lines, and what they should be!
    # TODO use pandas...
    fixme_lines = []
    fixme_lines.append(['True/Mir', 'Mag/IP', 'BP', 'BP Connector',
                        'iBB/P2B2 Connector', 'SBC Flex Name',
                        '4-asic group / DCB power', 'M/S/A',
                        'Cav. Map. PPP Positronic', 'Cav. Map. PPP Pins',
                        'Surf. Map. (Correct) PPP Positronic',
                        'Surf. Map. (Correct) PPP Pins'])
    for nl in ppp_wrong_cavern_lines:
        fixme_lines.append([true_mirror(nl.x, nl.y, nl.z), nl.z, nl.bp,
                            nl.bp_con, nl.ibbp2b2, nl.flex, nl.load,
                            nl.msa, ppp_wrong_cavern_lines[nl].ppp,
                            ppp_wrong_cavern_lines[nl].ppp_pin + ',' +
                            ppp_ret_pin(ppp_wrong_cavern_lines[nl].ppp_pin),
                            nl.ppp, nl.ppp_pin + ',' + ppp_ret_pin(nl.ppp_pin)])
    fixme_ppp_csv = open('fixme/ppp_fixes.csv', 'w')
    writer = csv.writer(fixme_ppp_csv)
    for row in fixme_lines: writer.writerow(row)
    fixme_ppp_csv.close()

    for comp in compare:
        print(f'\n\nChecking {nominal} vs {comp}...\n\n')
        # TODO comparisons...

    print(f'\n\nWriting (unformatted) fixed cavern mapping...\n\n')
    fixed = 'fixed-'+cavern


no_typos = cavern_typo_check()
if no_typos: cavern_check_fix()
