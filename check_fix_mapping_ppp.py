# Check the cavern mapping for typos (that can't be easily handled), then (once
# user fixes typos) check the surface mapping vs the cavern mapping (and also
# against any other references, eg. the sheet used to test the produced PEPI
# cables) PPP info. Output cavern differences to files in fixme folder,
# including directions for shifters fixing the cables by moving labels.
# Note: it is assumed here that the LVR info in the cavern mapping is correct;
# another script would need to be prepared to check this vs Phoebe's schematic.

import pandas, os, fnmatch, math
import csv
from argparse import ArgumentParser

# problem seems to be restricted to hybrid mag mirror (stereo+straight)
only_hyb_mag_mir = True

# see if flipping stereo<->straight helps
check_stereo_straight_flip = False

### grab command line args
parser = ArgumentParser(description='Check and fix cavern mapping')
# designate nominal and cavern mappings
parser.add_argument('nominal', help='specify nominal mapping to be used as input')
parser.add_argument('cavern', help='specify cavern mapping to be used as input')
# note that only C side needs to be compared; A side mapping is equivalent
# also set sheets that will be compared for consistency
parser.add_argument('doCompare', help='specify if compare mappings should be checked')
parser.add_argument('swap', default='NA', help='specify how Posistronix are swapping')
args = parser.parse_args()
nominal = args.nominal
cavern = args.cavern
compare = []
if (args.doCompare).lower()=='true':
    compare = os.listdir('compare')
    compare = ['compare/'+file for file in compare]
swap_pos = args.swap

# from parseXls import * # steal what Mark did to parse the cavern file
# actually, don't import here; messes up argparse, and I'm too lazy to debug it

# basic class that will uniquely identify (with redundancy) power lines, along
# with PPP connector and pin; LVR/length info not available in all sheets, so
# don't set those variables on construction
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
        self.length_c = '-1' # set this explicitly later
        self.length_a = '-1' # set this explicitly later
        self.lvr = '-1' # set this explicitly later
        self.lvr_ch = '-1' # set this explicitly later
        self.ppp_label = 'NA' # Petr's PPP label for cavern; redundant info
        self.lvr_label = 'NA' # Petr's LVR label for cavern; redundant info

    def equal_minus_ppp(self, other):
        return (self.x==other.x and self.y==other.y and self.z==other.z and
                self.bp==other.bp and self.bp_con==other.bp_con and
                self.ibbp2b2==other.ibbp2b2 and self.flex==other.flex and
                self.load==other.load and self.msa==other.msa)

    def equal_pepi_ppp(self, other):
        return (self.x==other.x and self.y==other.y and self.z==other.z and
                self.ppp==other.ppp and self.ppp_pin==other.ppp_pin)

    def set_length(self, length_c, length_a):
        self.length_c = str(length_c)
        if (not self.length_c == 'splice') and (not '|' in self.length_c):
            # not sure what python is doing with floats here...
            self.length_c = one_dec_str(self.length_c)
        self.length_a = str(length_a)
        if (not self.length_a == 'splice') and (not '|' in self.length_a):
            self.length_a = one_dec_str(self.length_a)

    def set_lvr(self, lvr, lvr_ch):
        self.lvr = str(lvr)
        self.lvr_ch = str(lvr_ch)

    def set_labels(self, ppp_label, lvr_label):
        self.ppp_label = ppp_label
        self.lvr_label = lvr_label

    def flip_stereo_straight_line(self):
        # leave all other info the same, just change flex and BP con
        if "JD" in self.bp_con: return self # leave DCBs alone
        mirror=False
        if self.x=='C' and ((self.y=='bot' and self.z=='mag') or
                            (self.y=='top' and self.z=='ip')):
            mirror = True
        if self.x=='A' and ((self.y=='bot' and self.z=='ip') or
                            (self.y=='top' and self.z=='mag')):
            mirror = True
        new_flex = self.flex
        if new_flex[0]=='X': new_flex = 'S'+new_flex[1:]
        else: new_flex = 'X'+new_flex[1:]
        new_bp_con = bp_con_alt_to_JP(new_flex, mirror)
        new_line = line(self.x, self.y, self.z, self.bp, new_bp_con,
                        self.ibbp2b2, new_flex, self.load, self.msa,
                        self.ppp, self.ppp_pin)
        new_line.set_length(self.length_c, self.length_a)
        new_line.set_lvr(self.lvr, self.lvr_ch)
        return new_line

    def __eq__(self, other):
        return (self.x==other.x and self.y==other.y and self.z==other.z and
                self.bp==other.bp and self.bp_con==other.bp_con and
                self.ibbp2b2==other.ibbp2b2 and self.flex==other.flex and
                self.load==other.load and self.msa==other.msa and
                self.ppp==other.ppp and self.ppp_pin==other.ppp_pin)

    def __hash__(self):
        return hash(self.x+self.y+self.z+self.bp+self.bp_con+self.ibbp2b2+
                    self.flex+self.load+self.msa+self.ppp+self.ppp_pin+
                    self.length_c+self.length_a+self.lvr+self.lvr_ch)


####### Helpers

def one_dec_str(float_str):
    return str(round(float(float_str), 1))

# stealing my old code from generating the surface mappings
def bp_con_JP_to_alt(JP, mirror): # converts JP # to alt BP connector notation
  index=int(JP[2:])
  if mirror:
    if (index//2)%2==0: index+=2
    else: index-=2
  alt_labels = ["X0M","X0S","S0S","S0M","X1M","X1S","S1S","S1M","X2M","X2S",
                "S2S","S2M"]
  return alt_labels[index]

def bp_con_alt_to_JP(alt, mirror): # converts alt BP connector notation to JP #
  # take advantage of mapping being 1-to-1
  for JP in range(12):
    if bp_con_JP_to_alt("JP"+str(JP), mirror)==alt: return "JP"+str(JP)
  return None # should never return here

# just going to copy & paste to steal Mark's parsing functions... and edit
# slightly
#
# This part is common to the hybrids and DCBs
#
def parseSheet( xls, sheet ):
    dfIn = pandas.read_excel(xls, sheet , usecols="C:G,L,N:O", skiprows=[0,1] )
    cols = list(dfIn.columns)
    cols[-4] = 'LVR'
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
    # del dfIn['LVR']
    # del dfIn['LVR ID - Connector - Pin']
    # del dfIn['PPP Connector - Pin']
    return dfIn

# for C side only
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
        if only_hyb_mag_mir:
            if not 'C-mag-bot' in sheet:
                continue
        sheet_info = sheet.split('-')
        x, y, z, bp = sheet_info[0], sheet_info[2], sheet_info[1], sheet_info[3]
        df = pandas.read_excel(xlsx, sheet, usecols='A:G')
        # get rid of empty rows
        df = df.dropna(how='all')
        for ind, row in df.iterrows():
            if only_hyb_mag_mir:
                if 'JD' in row['BP Connector']: continue
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
        if only_hyb_mag_mir: continue
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
            l.set_length(row['C L (m)'], row['A L (m)'])
            l.set_lvr(row['LVR ID'], row['LVR Channel'])
            ppp_label = row['PPP Connector - Pin'] + ' / ' + \
                        (row['LVR Name']).replace('_LV_SRC/RET','')
            lvr_label = row['LVR ID - Connector - Pin'] + ' ' + \
                        row['SBC section'] + ' / ' + \
                        (row['LVR Name']).replace('_LV_SRC/RET','')
            l.set_labels(ppp_label, lvr_label)
            lines.append(l)

    for hyb_sheet in hyb_sheets:
        if only_hyb_mag_mir:
            if not 'Hybrid - Mag - Mirror' in hyb_sheet: continue
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
            l.set_length(row['C L (m)'], row['A L (m)'])
            l.set_lvr(row['LVR ID'], int(row['LVR Channel']))
            ppp_label = row['PPP Connector - Pin'] + ' | ' + \
                        (row['LVR Name']).replace('_LV_SRC/RET','')
            lvr_label = row['LVR ID - Connector - Pin'] + ' ' + \
                        row['SBC section'] + ' | ' + \
                        (row['LVR Name']).replace('_LV_SRC/RET','')
            l.set_labels(ppp_label, lvr_label)
            lines.append(l)

    # combine splice lines by going through lines, checking which are have
    # length = 'splice', and adding LVR ch to other spliced line. then, go
    # through lines and delete redundant splice lines
    combined_splice_lines = []
    for cl in lines:
        if cl.length_c == 'splice':
            splice_found = 0
            for other_cl in lines:
                if (other_cl == cl) and (other_cl.length_c != cl.length_c):
                    splice_found += 1
                    other_cl.set_lvr(other_cl.lvr, other_cl.lvr_ch+' Y '+cl.lvr_ch)
                    other_cl.set_labels(other_cl.ppp_label, other_cl.lvr_label+\
                                        '   Y   '+cl.lvr_label)
            if splice_found==0 or splice_found>1: print(f'Splice problem??')
    for cl in lines:
        if not cl.length_c == 'splice': combined_splice_lines.append(cl)

    return combined_splice_lines

# return the cavern_lines with PPP positronic swapped according to input file
def parse_swap_pos(file, cavern_lines):
    xlsx = pandas.ExcelFile(file)
    sheets = xlsx.sheet_names
    lines = []
    swap_sheet = sheets[0] # only 1 sheet
    df = pandas.read_excel(xlsx, swap_sheet, usecols='A,D')
    for ind, row in df.iterrows():
        old_pos = 'P'+str(int(row['Positronic']))
        new_pos = 'P'+str(int(row['Swap to']))
        for l in cavern_lines:
            if l.ppp == old_pos:
                ml = line(l.x, l.y, l.z, l.bp, l.bp_con, l.ibbp2b2, l.flex,
                          l.load, l.msa, new_pos, l.ppp_pin)
                ml.set_lvr(l.lvr, l.lvr_ch)
                ml.set_length(l.length_c, l.length_a)
                ml.set_labels(l.ppp_label, l.lvr_label) # don't move ppp_label!
                lines.append(ml)
                # print(l.x+l.y+l.z+l.bp+l.bp_con+l.ibbp2b2+l.flex+l.load+l.msa+
                #       l.ppp+l.ppp_pin+'  '+nl.x+nl.y+nl.z+nl.bp+nl.bp_con+
                #       nl.ibbp2b2+nl.flex+nl.load+nl.msa+nl.ppp+nl.ppp_pin)
    return lines

# returns a list of lines for cable test mapping
def parse_cable_test(file):
    xlsx = pandas.ExcelFile(file)
    sheets = xlsx.sheet_names
    lines = []
    for sheet in sheets:
        continue
    return lines


# Associate files with a function for parsing
parse_func = {}
files = [nominal, cavern, swap_pos] + compare
for file in files:
    if 'surface_LV_power_tests' in file: parse_func[file]=parse_surface
    elif 'LVR_PPP_Underground' in file: parse_func[file]=parse_cavern
    elif 'lvr_testing' in file: parse_func[file]=parse_cable_test
    elif 'swap_positronic' in file: parse_func[file]=parse_swap_pos
    else: print(f'\n\nDON\'T RECONGNIZE {file} FORMAT\n\n')


# def print_ppp_corrected_cavern_lines(xlfile, cav_to_nom):
#     return True

def sort_by_surf_ppp_layer(lines, pos_ind, pins_ind, layer_ind):
    # lines is just 2D array; should return 2D array sorted by positronic info
    # and separating stereo and straight
    # also, add an empty row between each ppp
    lines = sorted(lines, key=lambda l:l[pins_ind])
    lines_with_layer_refcol = [l+[int(l[pos_ind][1:])] for l in lines]
    lines = sorted(lines_with_layer_refcol, key=lambda l:l[-1])
    lines = [l[:-1] for l in lines] # drop refcol
    lines_with_layer_refcol = [l + [l[layer_ind][0]] for l in lines]
    lines = sorted(lines_with_layer_refcol, key=lambda l:l[-1])
    lines = [l[:-1] for l in lines] # drop refcol
    row_len = len(lines[0])
    empty_row = [' ' for j in range(row_len)]
    prev_ppp = 'P0'
    rows_to_insert_empty_row = []
    for i in range(len(lines)):
        l = lines[i]
        ppp = l[pos_ind]
        if ppp != prev_ppp: rows_to_insert_empty_row.append(i)
        prev_ppp = ppp
    offset = 0 # each time adding an empty row, lines grows
    for i in rows_to_insert_empty_row:
        lines.insert(i+offset, empty_row)
        offset += 1

    return lines

def count_positronic(lines, pos, pos_ind, len_ind):
    # counts the occurrence of positronic pos at len_ind col in lens, not
    # adding to count when the line is spliced
    count = 0
    for l in lines:
        if pos == l[pos_ind] and l[len_ind] != 'splice': count += 1
    return count

def add_pop_col(lines, pos_ind, len_ind):
    # adds a col to lines that counts the number of populated pins for the
    # associated PPP positronic
    return [l+[count_positronic(lines, l[pos_ind], pos_ind, len_ind)] for l \
                                                                       in lines]

####### Main Checking Functions

def cavern_typo_check():
    print(f'\n\nChecking {cavern} for typos...\n\n')
    nominal_lines = parse_func[nominal](nominal)
    cavern_lines = parse_func[cavern](cavern) # parse func will do some checks (TODO)
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
    # ppp_wrong_cavern_lines = {} # map from nom line to cav line
    # for nom_line in nominal_lines:
    #     cav_line = None
    #     found = 0
    #     for cl in cavern_lines:
    #         if cl.equal_minus_ppp(nom_line):
    #             cav_line = cl
    #             found += 1
    #     if found==0 : print(f'Couldn\'t find '+
    #                   f'{nom_line.x+nom_line.y+nom_line.z+nom_line.bp}'+
    #                   f'{nom_line.bp_con+nom_line.ibbp2b2+nom_line.flex}'+
    #                   f'{nom_line.load+nom_line.msa}???')
    #     # finding more than 1 is ok if hybrid M/s (spliced at LVR, but cavern
    #     # mapping includes both lines coming out of LVR); if not west/east pair
    #     # printed, then you should be worried...
    #     if found > 1: print(f'Found more than 1 '+
    #                   f'{nom_line.x+nom_line.y+nom_line.z+nom_line.bp}'+
    #                   f'{nom_line.bp_con+nom_line.ibbp2b2+nom_line.flex}'+
    #                   f'{nom_line.load+nom_line.msa}???')
    #     if not cav_line==nom_line:
    #         print(f'\nFound cavern line with wrong PPP!\n')
    #         ppp_wrong_cavern_lines[nom_line] = cav_line

    # write out all the wrong lines, and what they should be!
    # TODO use pandas...
    # fixme_lines = []
    # fixme_lines.append(['True/Mir', 'Mag/IP', 'BP', 'BP Connector',
    #                     'iBB/P2B2 Connector', 'SBC Flex Name',
    #                     '4-asic group / DCB power', 'M/S/A',
    #                     'Cav. Map. PPP Positronic', 'Cav. Map. PPP Pins',
    #                     'Surf. Map. PPP Positronic',
    #                     'Surf. Map. PPP Pins'])
    # for nl in ppp_wrong_cavern_lines:
    #     fixme_lines.append([true_mirror(nl.x, nl.y, nl.z), nl.z, nl.bp,
    #                         nl.bp_con, nl.ibbp2b2, nl.flex, nl.load,
    #                         nl.msa, ppp_wrong_cavern_lines[nl].ppp,
    #                         ppp_wrong_cavern_lines[nl].ppp_pin + ',' +
    #                         ppp_ret_pin(ppp_wrong_cavern_lines[nl].ppp_pin),
    #                         nl.ppp, nl.ppp_pin + ',' + ppp_ret_pin(nl.ppp_pin)])
    # fixme_ppp_csv = open('fixme/ppp_fixes.csv', 'w')
    # writer = csv.writer(fixme_ppp_csv)
    # for row in fixme_lines: writer.writerow(row)
    # fixme_ppp_csv.close()
    # ppp_wrong_cavern_lines = {} # map from (wrong) cav line to nom line
    ppp_corrected_cavern_lines = {} # from (all) cav line to nom line
    # also do comparison with all cavern lines flipped stereo<->straight
    # note: for this, compare against nominal line matched to non-flipped
    # cavern line!
    # ppp_wrong_cavern_lines_flip = {}
    ppp_corrected_cavern_lines_flip = {}
    ppp_corrected_cavern_lines_moved = {}
    for cav_line in cavern_lines:
        nom_line = line('na', 'na', 'na', 'na', 'na', 'na', 'na', 'na', 'na',
                        'na', 'na')
        found = 0
        for nl in nominal_lines:
            if nl.equal_minus_ppp(cav_line):
                found += 1
                # also, set the LVR info, lengths, and labels while you're at it
                nl.set_lvr(cav_line.lvr, cav_line.lvr_ch)
                nl.set_length(cav_line.length_c, cav_line.length_a)
                cav_ppp_label = cav_line.ppp_label.split(' | ')
                ppp_label = nl.ppp + ' - ' + nl.ppp_pin + '/' + \
                            ppp_ret_pin(nl.ppp_pin) + ' | ' + cav_ppp_label[1]
                nl.set_labels(ppp_label, cav_line.lvr_label)
                nom_line = nl
        if found==0: print(f'Couldn\'t find '+
                           f'{cav_line.x+cav_line.y+cav_line.z+cav_line.bp}'+
                           f'{cav_line.bp_con+cav_line.ibbp2b2+cav_line.flex}'+
                           f'{cav_line.load+cav_line.msa}???')
        if found>1: print(f'Found more than one '+
                          f'{cav_line.x+cav_line.y+cav_line.z+cav_line.bp}'+
                          f'{cav_line.bp_con+cav_line.ibbp2b2+cav_line.flex}'+
                          f'{cav_line.load+cav_line.msa}???')
        if not cav_line==nom_line:
            print(f'\nFound cavern line with wrong PPP!\n')
            # ppp_wrong_cavern_lines[cav_line] = nom_line
        ppp_corrected_cavern_lines[cav_line] = nom_line
        if check_stereo_straight_flip:
            cav_line_flip = cav_line.flip_stereo_straight_line()
            if not cav_line_flip==nom_line:
                print(f'\nFound flipped cavern line with wrong PPP!\n')
                # ppp_wrong_cavern_lines_flip[cav_line_flip] = nom_line
            ppp_corrected_cavern_lines_flip[cav_line_flip] = nom_line

    # also, if the user is specifying where the positronic are being swapped
    # to, want to figure out where shifters should move the LVR/PPP labels
    # (so that, with the pre-moved positronics taken into account, the
    # cables going into the given PPP location are the correct lines/LVRs)
    if swap_pos != 'NA':
        moved_cavern_lines = parse_func[swap_pos](swap_pos, cavern_lines)
        for cav_line in moved_cavern_lines:
            # should have already checked above (with print statements) that
            # a unique nom_line is found for each cav_line, but need to recreate
            # the map from cavern lines to nominal lines because of changed PPP
            # Note that labels for nominal lines are already set above!
            nom_line = line('na', 'na', 'na', 'na', 'na', 'na', 'na', 'na', 'na',
                            'na', 'na')
            found = 0
            for nl in nominal_lines:
                if nl.equal_pepi_ppp(cav_line):
                    found += 1
                    nom_line = nl
            if found==0: print(f'Couldn\'t find nominal line at '+
                               f'{cav_line.x+cav_line.y+cav_line.z}'+
                               f'{cav_line.ppp+cav_line.ppp_pin}???')
            if found>1: print(f'Found more than one nominal line at '+
                              f'{cav_line.x+cav_line.y+cav_line.z}'+
                              f'{cav_line.ppp+cav_line.ppp_pin}???')
            if not cav_line==nom_line:
                print(f'\nFound (moved) cavern line with wrong PPP!\n')
            ppp_corrected_cavern_lines_moved[cav_line] = nom_line

    # write out all cavern lines
    # TODO should put this in a separate function...
    # print_ppp_corrected_cavern_lines('fixme/cavern_mapping_ppp_fixes.xlsx',
    #                                  ppp_corrected_cavern_lines)
    cav_lines = []
    cav_lines.append(['True/Mir', 'Mag/IP', 'BP', 'BP Con.',
                      'iBB/P2B2 Con.', 'SBC Flex Name',
                      '4-asic group / DCB power', 'M/S/A',
                      'Cav. Map. PPP Pos.', 'Cav. Map. PPP Pins',
                      'Surf. Map. PPP Pos.',
                      'Surf. Map. PPP Pins', 'LVR', 'LVR Ch.',
                      'C Len (m)', 'A Len (m)'])
    for cl in ppp_corrected_cavern_lines:
        cav_lines.append([true_mirror(cl.x, cl.y, cl.z), cl.z, cl.bp,
                          cl.bp_con, cl.ibbp2b2, cl.flex, cl.load,
                          cl.msa, cl.ppp, cl.ppp_pin + ',' +
                          ppp_ret_pin(cl.ppp_pin),
                          ppp_corrected_cavern_lines[cl].ppp,
                          ppp_corrected_cavern_lines[cl].ppp_pin + ',' +
                          ppp_ret_pin(ppp_corrected_cavern_lines[cl].ppp_pin),
                          cl.lvr, cl.lvr_ch, cl.length_c, cl.length_a])
    cav_lines = [cav_lines[0] + ['Cav. Map. PPP Pop.']] + \
                 add_pop_col(cav_lines[1:],
                      cav_lines[0].index('Cav. Map. PPP Pos.'),
                      cav_lines[0].index('C Len (m)'))
    cav_lines = [cav_lines[0] + ['Surf. Map. PPP Pop.']] + \
                 add_pop_col(cav_lines[1:],
                      cav_lines[0].index('Surf. Map. PPP Pos.'),
                      cav_lines[0].index('C Len (m)'))
    cav_lines = [cav_lines[0]] + \
                 sort_by_surf_ppp_layer(cav_lines[1:],
                      cav_lines[0].index('Surf. Map. PPP Pos.'),
                      cav_lines[0].index('Surf. Map. PPP Pins'),
                      cav_lines[0].index('SBC Flex Name'))

    fixme = open('fixme/ppp_fixes.csv', 'w')
    writer = csv.writer(fixme)
    for row in cav_lines: writer.writerow(row)
    fixme.close()

    # write out all flipped cavern lines
    if check_stereo_straight_flip:
        cav_lines_flip = []
        cav_lines_flip.append(['True/Mir', 'Mag/IP', 'BP', 'BP Con.',
                               'iBB/P2B2 Con.', 'SBC Flex Name',
                               '4-asic group / DCB power', 'M/S/A',
                               'Cav. Map. PPP Pos.', 'Cav. Map. PPP Pins',
                               'Surf. Map. PPP Pos.',
                               'Surf. Map. PPP Pins', 'LVR', 'LVR Ch.',
                               'C Len (m)', 'A Len (m)'])
        for cl in ppp_corrected_cavern_lines_flip:
            cav_lines_flip.append([true_mirror(cl.x, cl.y, cl.z), cl.z, cl.bp,
                                   cl.bp_con, cl.ibbp2b2, cl.flex, cl.load,
                                   cl.msa, cl.ppp, cl.ppp_pin + ',' +
                                   ppp_ret_pin(cl.ppp_pin),
                                   ppp_corrected_cavern_lines_flip[cl].ppp,
                                   ppp_corrected_cavern_lines_flip[cl].ppp_pin + ',' +
                                   ppp_ret_pin(ppp_corrected_cavern_lines_flip[cl].ppp_pin),
                                   cl.lvr, cl.lvr_ch, cl.length_c, cl.length_a])
        cav_lines_flip = [cav_lines_flip[0] + ['Cav. Map. PPP Pop.']] + \
                          add_pop_col(cav_lines_flip[1:],
                               cav_lines_flip[0].index('Cav. Map. PPP Pos.'),
                               cav_lines_flip[0].index('C Len (m)'))
        cav_lines_flip = [cav_lines_flip[0] + ['Surf. Map. PPP Pop.']] + \
                          add_pop_col(cav_lines_flip[1:],
                               cav_lines_flip[0].index('Surf. Map. PPP Pos.'),
                               cav_lines_flip[0].index('C Len (m)'))
        cav_lines_flip = [cav_lines_flip[0]] + \
                          sort_by_surf_ppp_layer(cav_lines_flip[1:],
                             cav_lines_flip[0].index('Surf. Map. PPP Pos.'),
                             cav_lines_flip[0].index('Surf. Map. PPP Pins'),
                             cav_lines_flip[0].index('SBC Flex Name'))

        fixme_flip = open('fixme/ppp_fixes_flipped.csv', 'w')
        writer_flip = csv.writer(fixme_flip)
        for row in cav_lines_flip: writer_flip.writerow(row)
        fixme_flip.close()

    # write out where labels should move to
    if swap_pos != 'NA':
        cav_lines_moved = []
        # keep a few extra columns (mostly just for sorting)
        # also keep track of actual cable lengths, not just required!
        cav_lines_moved.append(['True/Mir', 'Mag/IP', 'BP', 'BP Con.',
                                'iBB/P2B2 Con.', 'SBC Flex Name',
                                '4-asic group / DCB power', 'M/S/A',
                                'PPP Pos. (Correct)', 'PPP Pins (Correct)',
                                'LVR', 'LVR Ch.', 'Actual C L (m)',
                                'Actual A L (m)', 'C Len (m)', 'A Len (m)',
                                'Cav. Map. PPP Label (After Moving Pos.)',
                                'Replace w/ PPP Label',
                                'Cav. Map. LVR Label (After Moving Pos.)',
                                'Replace w/ LVR Label'])
        for cl in ppp_corrected_cavern_lines_moved:
            nl = ppp_corrected_cavern_lines_moved[cl]
            cav_lines_moved.append([true_mirror(nl.x, nl.y, nl.z), nl.z, nl.bp,
                              nl.bp_con, nl.ibbp2b2, nl.flex, nl.load,
                              nl.msa, nl.ppp, nl.ppp_pin + ',' +
                              ppp_ret_pin(nl.ppp_pin), nl.lvr, nl.lvr_ch,
                              cl.length_c, cl.length_a, nl.length_c,
                              nl.length_a, cl.ppp_label, nl.ppp_label,
                              cl.lvr_label, nl.lvr_label])
        cav_lines_moved = [cav_lines_moved[0]] + \
                           sort_by_surf_ppp_layer(cav_lines_moved[1:],
                              cav_lines_moved[0].index('PPP Pos. (Correct)'),
                              cav_lines_moved[0].index('PPP Pins (Correct)'),
                              cav_lines_moved[0].index('SBC Flex Name'))

        fixme_moved = open('fixme/move_labels.csv', 'w')
        writer_moved = csv.writer(fixme_moved)
        for row in cav_lines_moved: writer_moved.writerow(row)
        fixme_moved.close()



    for comp in compare:
        print(f'\n\nChecking {nominal} vs {comp}...\n\n')
        # TODO comparisons...

    # print(f'\n\nWriting (unformatted) fixed cavern mapping...\n\n')
    # fixed = 'fixed-'+cavern


no_typos = cavern_typo_check()
if no_typos: cavern_check_fix()
