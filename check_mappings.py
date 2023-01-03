# Check the cavern mapping for typos (that can't be easily handled), then (once
# user fixes typos) check the surface mapping vs the cavern mapping (and also
# against any other references, eg. the sheet used to test the produced PEPI
# cables) PPP info. Output cavern differences to files in fixme folder,
# including directions for shifters fixing the cables by moving labels.
# Note: this script can also check the LV cavern LVR-load mapping is correct vs
# Phoebe's schematics.
# Note: the script also now checks that Petr's LV labels were generated
# correctly. Here, the PPP color info isn't checked.
# Note: script also outputs a table for the underground power.

import pandas, os, fnmatch, math, copy, csv
from argparse import ArgumentParser

# problem seems to be restricted to hybrid mag mirror (stereo+straight)
only_hyb_mag_mir = False

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
parser.add_argument('doCheckLines', help='specify if cavern mapping LVR-load should be checked')
parser.add_argument('-sense', '-c', default='NA', help='specify cavern sense mapping to be checked')
parser.add_argument('-swap', '-s', default='NA', help='specify how Posistronix are swapping')
args = parser.parse_args()
nominal = args.nominal
cavern = args.cavern
compare = []
if (args.doCompare).lower()=='true':
    compare = os.listdir('compare')
    compare = ['compare/'+file for file in compare]
schem_ip = 'nominal/PEPI_a_SIDE_g3.NET'
schem_mag = 'nominal/PEPI_b_SIDE_g3.NET'
tbb_schem = 'nominal/TelemetryBB_Mirror_FINAL_mpeco.NET'
cavern_sense = args.sense
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
        self.z = z # z of load being powered, not of LVR in SBC
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

    def get_lvr_pins(self):
        chs = self.lvr_ch.split(' Y ')
        res = ''
        for chn in chs:
            ch = int(float(chn))
            con = 'n/a'
            pins = 'n/a'
            if ch<5:
                con = 'J12'
                pin = 10-2*ch
                pins = f'{pin}/{pin-1}'
            else:
                con = 'J13'
                pin = 18-2*ch
                pins = f'{pin}/{pin-1}'
            res += f'{self.lvr} - {con} - {pins}  Y  '
        res = res[:-5] # get rid of final '  Y  '
        return res

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


# basic class that will uniquely identify (with redundancy) sense lines
# will have one senseline object for 4 actual sense lines: one in each SB; the
# only true/mir difference is the PPP connector and load served (ie. what the
# mirror vs true tBB maps to in the JP notation)
class senseline:

    def __init__(self, crate, slot, lvr, lvr_con, lvr_twistpair, spltr,
                 out_spltr, in_spltr, in_twistpair, in_label, ppp_true,
                 ppp_mir, tbb_con):
        self.crate = crate
        self.slot = slot
        self.lvr = lvr
        self.lvr_con = lvr_con # J10 or J16
        self.lvr_twistpair = lvr_twistpair # 1-2, 4-5, 3-6, or 7-8
        self.spltr = spltr # the full splitter label
        self.out_spltr = out_spltr
        self.in_spltr = in_spltr
        self.in_twistpair = in_twistpair
        self.in_label = in_label # contains bp info
        self.ppp_true = ppp_true
        self.ppp_mir = ppp_mir
        self.tbb_con = tbb_con

    def in_spltr_lab(self):
        spltr_type=get_spltr_type(self.spltr)
        if spltr_type=='direct': return '-'
        elif spltr_type in ['1', '4']: return self.spltr
        elif spltr_type in ['2', '3', '6']: return f'{self.spltr}_{self.in_spltr}'
        else: print('what?')

    def out_spltr_lab(self):
        spltr_type=get_spltr_type(self.spltr)
        if spltr_type=='direct': return spltr_type
        elif spltr_type in ['1', '2', '3', '4', '6']:
            return f'Type {spltr_type} - OUT {self.out_spltr}'
        else: print('huh?')

    def get_bp(self):
        return self.in_label.split('_')[0]

    # from tBB schematic! For DCBs I have flex = 1V5 or 2V5
    def get_flex(self, tbb_map):
        tbb_twistpair = f'{self.tbb_con}_{self.in_twistpair}'
        if not tbb_twistpair in tbb_map:
            print(f'couldnt find {tbb_twistpair}!')
            return 'n/a'
        return (tbb_map[tbb_twistpair].split('_'))[0]

    def get_load(self, tbb_map):
        tbb_twistpair = f'{self.tbb_con}_{self.in_twistpair}'
        if not tbb_twistpair in tbb_map:
            print(f'couldnt find {tbb_twistpair}!')
            return 'n/a'
        return (tbb_map[tbb_twistpair].split('_'))[1]

    def get_msa(self, lines): return


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

# returns channel that sense line is associated w/ on LVR
def lvr_twistpair_to_ch(sense_con, twistpair):
    if sense_con=='J10':
        if twistpair=='1-2': return '1'
        elif twistpair=='4-5': return '2'
        elif twistpair=='3-6': return '3'
        elif twistpair=='7-8': return '4'
        else: return '???'
    elif sense_con=='J16':
        if twistpair=='1-2': return '5'
        elif twistpair=='4-5': return '6'
        elif twistpair=='3-6': return '7'
        elif twistpair=='7-8': return '8'
        else: return '????'
    else:
        return '?????'

def get_spltr_type(spltr_lab):
    if not 'S' in spltr_lab: return 'direct'
    else: return spltr_lab[1] # all splitter types are 1 digit

# return pin in 'J12_4' etc format (only src pin)
def lvr_ch_to_pin(chn):
    ch = int(float(chn))
    con = 'n/a'
    pin = 'n/a'
    if ch<5:
        con = 'J12'
        pin = f'{10-2*ch}'
    else:
        con = 'J13'
        pin = f'{18-2*ch}'
    return f'{con}_{pin}'

# define the connections in the splitter boards; return 'n/a' if out twisted pair
# maps nowhere, else return the in [port, twisted pair] that it maps to
def spltr1(out_port, out_twistpair):
    if out_port=='a':
        if out_twistpair=='1-2': return ['-', '1-2']
        if out_twistpair=='3-6': return ['-', '3-6']
    elif out_port=='b':
        if out_twistpair=='1-2': return ['-', '4-5']
        if out_twistpair=='3-6': return ['-', '7-8']
    else:
        return False # ugly to mix types, but fine

def spltr2(out_port, out_twistpair):
    if out_port=='a':
        if out_twistpair=='1-2': return ['1', '1-2']
        if out_twistpair=='4-5': return ['1', '4-5']
        if out_twistpair=='3-6': return ['2', '3-6']
        if out_twistpair=='7-8': return ['1', '7-8']
    elif out_port=='b':
        if out_twistpair=='1-2': return ['3', '1-2']
        if out_twistpair=='4-5': return ['3', '4-5']
        if out_twistpair=='3-6': return ['2', '7-8']
        if out_twistpair=='7-8': return ['3', '7-8']
    else:
        return False

def spltr3(out_port, out_twistpair):
    if out_port=='a':
        if out_twistpair=='1-2': return ['1', '1-2']
        if out_twistpair=='4-5': return ['2', '1-2']
        if out_twistpair=='3-6': return ['1', '3-6']
    elif out_port=='b':
        if out_twistpair=='1-2': return ['1', '4-5']
        if out_twistpair=='4-5': return ['2', '4-5']
        if out_twistpair=='3-6': return ['1', '7-8']
        if out_twistpair=='7-8': return ['2', '7-8']
    else:
        return False

# identical to type 1 board... but make sure input isn't ['1', '7-8']!
def spltr4(out_port, out_twistpair):
    ret = spltr1(out_port, out_twistpair)
    if (not ret==False) and ret[1]=='7-8':
        print(f'Found a supposed type 4 splitter w 7-8 input used!')
        return False
    return ret

# type 5 splitter no longer in use!

def spltr6(out_port, out_twistpair):
    if out_port=='a':
        if out_twistpair=='1-2': return ['1', '4-5']
        if out_twistpair=='3-6': return ['1', '7-8']
    elif out_port=='b':
        if out_twistpair=='1-2': return ['1', '1-2']
        if out_twistpair=='4-5': return ['3', '1-2']
        if out_twistpair=='3-6': return ['1', '3-6']
        if out_twistpair=='7-8': return ['2', '1-2']
    elif out_port=='c':
        if out_twistpair=='1-2': return ['2', '4-5']
        if out_twistpair=='3-6': return ['2', '7-8']
    elif out_port=='d':
        if out_twistpair=='1-2': return ['3', '4-5']
        if out_twistpair=='3-6': return ['3', '7-8']
    else:
        return False

# convert how Phoebe labels loads to my convention (for sense check)
def load_label_phoebe_to_me(lab):
    splt = lab.split('_')
    if 'dcb' in lab.lower():
        if '25' in lab: # 2V5 only
            return f'{splt[3]}_2V5_{splt[2]}-{splt[1]}'
        else: # 1V5 only
            return f'{splt[2]}_1V5_{splt[1]}'
    else: # hybrids only
        return f'{splt[2]}_{splt[3]}_{splt[4]}'

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

def twisted_ret(twisted_src):
    src = int(twisted_src)
    ret = 'n/a'
    if src==1: ret = '2'
    elif src==4: ret = '5'
    elif src==3: ret = '6'
    elif src==7: ret = '8'
    else: print('bad twisted src')
    return ret

def petr_filename_to_xyz(file):
    # all filenames have same length; get rid of any parent folders
    file = file[-15:]
    x, y, z = file[0], file[1], file[2]
    # change y,z to your naming convention
    if y=='B': y='bot'
    if y=='T': y='top'
    if z=='M': z='mag'
    if z=='I': z='ip'
    return [x,y,z]

# PPP color uniquely defined by where the Pos is: mag/IP and Pos num.
def check_ppp_color(z, ppp, color):
    ppp_num = int(ppp[1:])
    if ppp_num <= 18:
        if z=='mag':
            if color!='blu': return False
        if z=='ip':
            if color!='grn': return False
    else:
        if z=='mag':
            if color!='red': return False
        if z=='ip':
            if color!='yel': return False
    return True

# returns True if there is a (A or M) load for the associated LVR ch, else False
def senseline_used(lvr, con, twistpair, power_map):
    # make sure that the corresponding power line exists, and then make sure it
    # isn't a slave line (taking into account the lines which are incorrectly
    # assigned to slave loads)
    lvr_power_pin = f'{lvr}_{lvr_ch_to_pin(lvr_twistpair_to_ch(con, twistpair))}'
    # Phoebe's mapping is extremely annoying, so just put in the slave lines
    # by hand...
    slaves = ['43_J12_6', '43_J12_2', '43_J13_6', '43_J13_2',
              '44_J12_6', '44_J12_2', '44_J13_6', '44_J13_2',
              '53_J12_6', '53_J12_2', '53_J13_6', '53_J13_2',
              '54_J12_6', '54_J12_2', '54_J13_6', '54_J13_2',
              '45_J12_6', '45_J12_2', '55_J12_6', '55_J12_2',
              '57_J12_6', '57_J12_2', '57_J13_6', '57_J13_2',
              '58_J12_6', '58_J12_2', '58_J13_6', '58_J13_2',
              '59_J12_6', '59_J12_2', '59_J13_6', '59_J13_2',
              '60_J12_6', '60_J12_2', '60_J13_6', '60_J13_2',
              '62_J12_6', '62_J12_2', '62_J13_6', '62_J13_2',
              '63_J12_6', '63_J12_2', '63_J13_6', '63_J13_2',
              '64_J12_6', '64_J12_2', '64_J13_6', '64_J13_2',
              '61_J12_2', '61_J13_2',
              '9_J12_6', '9_J12_2', '9_J13_6', '9_J13_2',
              '10_J12_6', '10_J12_2', '10_J13_6', '10_J13_2',
              '21_J12_6', '21_J12_2', '21_J13_6', '21_J13_2',
              '22_J12_6', '22_J12_2', '22_J13_6', '22_J13_2',
              '11_J12_6', '11_J12_2', '23_J12_6', '23_J12_2',
              '25_J12_6', '25_J12_2', '25_J13_6', '25_J13_2',
              '26_J12_6', '26_J12_2', '26_J13_6', '26_J13_2',
              '27_J12_6', '27_J12_2', '27_J13_6', '27_J13_2',
              '30_J12_6', '30_J12_2', '30_J13_6', '30_J13_2',
              '28_J12_6', '28_J12_2', '28_J13_6', '28_J13_2',
              '33_J12_6', '33_J12_2', '33_J13_6', '33_J13_2',
              '31_J12_6', '31_J12_2', '31_J13_6', '31_J13_2',
              '29_J12_2', '29_J13_2', '32_J12_2', '32_J13_2']
    if not lvr_power_pin in power_map or lvr_power_pin in slaves: return False
    return True


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
            lvr_name = ((row['LVR Name']).replace('_P/N_S','')).replace('_P/N','')
            ppp_label = row['PPP Connector - Pin'] + ' | ' + lvr_name
            lvr_label = row['LVR ID - Connector - Pin'] + ' ' + \
                        row['SBC section'] + ' | ' + lvr_name
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
        pos = 'P'+str(int(row['Positronic']))
        for l in cavern_lines:
            pos_tmp = pos
            if (l.x=='C' and l.y=='bot' and l.z=='mag' and l.flex!='n/a'):
                pos_tmp = 'P'+str(int(row['Swap to'])) # don't move non-HMM Pos!
            if l.ppp == pos:
                ml = line(l.x, l.y, l.z, l.bp, l.bp_con, l.ibbp2b2, l.flex,
                          l.load, l.msa, pos_tmp, l.ppp_pin)
                ml.set_lvr(l.lvr, l.lvr_ch)
                ml.set_length(l.length_c, l.length_a)
                ml.set_labels(l.ppp_label, l.lvr_label) # don't move ppp_label!
                lines.append(ml)
                # print(l.x+l.y+l.z+l.bp+l.bp_con+l.ibbp2b2+l.flex+l.load+l.msa+
                #       l.ppp+l.ppp_pin+'  '+nl.x+nl.y+nl.z+nl.bp+nl.bp_con+
                #       nl.ibbp2b2+nl.flex+nl.load+nl.msa+nl.ppp+nl.ppp_pin)
    return lines

# returns a list of lines for cable test mapping; TODO
def parse_cable_test(file):
    xlsx = pandas.ExcelFile(file)
    sheets = xlsx.sheet_names
    lines = []
    for sheet in sheets:
        continue
    return lines

# parse Phoebe's netlists; return a map of LVR ch to load
def parse_netlist(netlist):
    lvrch_load = {}
    with open(netlist) as fp:
        myline=fp.readline()
        netmap = {}
        while myline:
            if (myline.find("PCBComponent") > 0):
                connector = myline.split()[3]
                array=[]
                myline=fp.readline()
                while (myline.find("(") > 0):
                    if(myline.find("J1") ==-1):
                        line = myline.split()
                        array.append(f'LVPin{line[1]}|{line[2]}')
                    else: array.append("x")
                    myline=fp.readline()
                netmap[connector]=array
            myline = fp.readline()
        #target = "LVReg_X-Y_1.2_"
        #for j in range(13,25,1):
        #    print("\n"+str(j))
        #    for i in reversed(netmap["J12_"+target+str(j)]): print(i)
        #    for i in reversed(netmap["J13_"+target+str(j)]): print(i)
        for conn in netmap.keys():
            for net in netmap[conn]:
                if ("_SRC" in net) or (net.endswith("_P")):
                    if("PT_" in net) or ("DCB_" in net):
                        net_parts = net.split('|')
                        pin = (net_parts[0])[-1]
                        load = net_parts[1]
                        load = load.replace('_LV_SRC','')
                        load = load.replace('_25_P', '_25')
                        load = load.replace('_b_P', '_b')
                        load = load.replace('_a_P', '_a')
                        conn_parts = conn.split('_')
                        lvr_out_con = conn_parts[0]
                        lvr = conn_parts[-1]
                        lvrch_load[f'{lvr}_{lvr_out_con}_{pin}'] = load
    return lvrch_load

# parse and check Petr's LVR labels
def parse_check_petr_lvr(file, correct_lines):
    my_txt = file[:-4]+'_alex.txt'
    with open(file) as f, open(my_txt, 'w') as mf:
        xyz = petr_filename_to_xyz(file)
        x, y, z = xyz[0], xyz[1], xyz[2]
        lvr = ''
        for txt_line in f.readlines():
            # set the lvr number!
            if 'LVR ' in txt_line:
                mf.write('\n'+txt_line)
                lvr = (txt_line.split())[1]
            # skip useless lines
            if not ('J12' in txt_line or 'J13' in txt_line): continue
            txt_line_split = txt_line.split()
            lvr_ch = txt_line_split[0]
            lvr_pin = txt_line_split[1]
            ppp = txt_line_split[2]
            ppp_pin = txt_line_split[3][0] # only src pin
            ppp_color = txt_line_split[4]
            sbc_sec = txt_line_split[5]
            length_c = txt_line_split[6]
            # sometimes Petr converts floats to ints...
            if not '|' in length_c: length_c = str(float(length_c))
            # if ppp in ['P20', 'P21', 'P23', 'P24', 'P26', 'P27', 'P29', 'P30',
            #            'P32', 'P33', 'P35', 'P36']: continue # skip dcbs
            # print(f'LVR{lvr} ch{lvr_ch} (pin {lvr_pin}) Pos{ppp}:{ppp_pin} '+
            #       f'PPPReg:{ppp_color} SBCReg:{sbc_sec} L:{length_c}')

            # find the correct line by xyz+lvr+lvr_ch
            # once found, check everything else is right
            count = 0
            for correct_line in correct_lines:
                if (x==correct_line.x and y==correct_line.y and
                    z==correct_line.z and lvr==correct_line.lvr and
                    lvr_ch==correct_line.lvr_ch):
                    count += 1
                    split_label_lvr = ((correct_line.lvr_label).split(' | '))[0]
                    split_label_lvr = split_label_lvr.split()
                    correct_line_lvr_pin = split_label_lvr[2]+'_'+split_label_lvr[4]
                    correct_line_sbc = split_label_lvr[5]
                    if not (ppp==correct_line.ppp and
                            ppp_pin==correct_line.ppp_pin and
                            check_ppp_color(z, ppp, ppp_color) and
                            lvr_pin==correct_line_lvr_pin and
                            sbc_sec==correct_line_sbc and
                            length_c==correct_line.length_c):
                        print(f'rest of {x}{y}{z} LVR{lvr} ch{lvr_ch} line '+
                              'doesn\'t agree...')
                        print(f'Petr: LVR{lvr} ch{lvr_ch} (pin {lvr_pin}) '+
                              f'Pos{ppp}:{ppp_pin} PPPReg:{ppp_color} '+
                              f'SBCReg:{sbc_sec} L:{length_c}')
                        print(f'Me: LVR{correct_line.lvr} ch{correct_line.lvr_ch} '+
                              f' (pin {correct_line_lvr_pin}) '+
                              f'Pos{correct_line.ppp}:{correct_line.ppp_pin} '+
                              # f'PPPReg:{correct_line.ppp_color} '+
                              f'SBCReg:{correct_line_sbc} '+
                              f'L:{correct_line.length_c}')
            if count < 1: print(f'Couldn\'t find {x}{y}{z} LVR{lvr} ch{lvr_ch}')
            elif count > 1: print(f'Found multiple {x}{y}{z} LVR{lvr} ch{lvr_ch}')
            else: mf.write(txt_line)

# parse and check Petr's PPP labels; actually, skip this
def parse_check_petr_ppp(file, correct_lines):
    return

# parse the cavern sense table, outputting a list of senseline objects for
# each twisted pair (in 1 SB)
def parse_cavern_sense(file, power_map):
    xlsx = pandas.ExcelFile(file)
    sheets = xlsx.sheet_names
    senselines = []
    sheet = sheets[0] # only 1 sheet
    df = pandas.read_excel(xlsx, sheet, usecols='A:J')
    for lvr in range(1,68):
        for con in ['J10', 'J16']:
            for twistpair_out in ['1-2', '4-5', '3-6', '7-8']:
                if senseline_used(lvr, con, twistpair_out, power_map):
                    # found sense line that should exist, so trace it through
                    # the cavern sense (layout) map
                    # throughout, take advantage of the formatting choices
                    # in underground_LVsense_layout_table.xlsx
                    out_row = df.loc[df['LVR Port'] == f'{lvr}_{con}']
                    crate = out_row['Crate Number (of LVR)'].item()
                    slot = out_row['Crate Slot Number'].item()
                    spltr = out_row['Splitter/Cable Label'].item()
                    spltr_out = out_row['Splitter/Cable Output'].item()
                    spltr_type = get_spltr_type(spltr)
                    spltr_in_pair = ['-', twistpair_out] # direct cable
                    in_row = out_row # direct cable
                    if spltr_type=='1':
                        spltr_in_pair = spltr1(spltr_out, twistpair_out)
                        in_row = df.loc[df['Splitter/Cable Label'] == spltr]
                        in_row = in_row.loc[in_row['Splitter/Cable Output'] == 'a']
                    elif spltr_type=='2':
                        spltr_in_pair = spltr2(spltr_out, twistpair_out)
                        in_row = df.loc[df['Splitter/Cable Input'] == f'{spltr}_{spltr_in_pair[0]}']
                    elif spltr_type=='3':
                        spltr_in_pair = spltr3(spltr_out, twistpair_out)
                        in_row = df.loc[df['Splitter/Cable Input'] == f'{spltr}_{spltr_in_pair[0]}']
                    elif spltr_type=='4':
                        spltr_in_pair = spltr4(spltr_out, twistpair_out)
                        in_row = df.loc[df['Splitter/Cable Label'] == spltr]
                        in_row = in_row.loc[in_row['Splitter/Cable Output'] == 'a']
                    elif spltr_type=='6':
                        spltr_in_pair = spltr6(spltr_out, twistpair_out)
                        in_row = df.loc[df['Splitter/Cable Input'] == f'{spltr}_{spltr_in_pair[0]}']
                    else:
                        if not spltr_type=='direct': print('dont recognize spltr...')
                    if spltr_in_pair==False: print('\n\nThis shouldnt happen...\n\n')
                    label_in = in_row['Sense Line Label'].item()
                    ppp_true = in_row['True PPP RJ45 Coupler'].item()
                    ppp_mir = in_row['Mirror PPP RJ45 Coupler'].item()
                    tbb_con = in_row['tBB Port'].item()
                    senselines.append(senseline(crate, slot, str(lvr), con,
                    twistpair_out, spltr, spltr_out, spltr_in_pair[0],
                    spltr_in_pair[1], label_in, ppp_true, ppp_mir, tbb_con))
    return senselines

# parse tBB netlist, returning a map from tBB connector+twisted pair to sensed
# load
def parse_tbb(netlist):
    tbb_con_load = {}
    with open(netlist) as nl:
        con = 'J0' # not a real connector on tBB
        cons = [f'J{num}' for num in range(1,22)]
        rows = nl.readlines()
        for ind in range(len(rows)):
            row = rows[ind]
            row_split = row.split()
            # if you find a connector, list all the lines
            if (len(row_split)>=2) and (row_split[-2] in cons):
                con = row_split[-2]
                for con_ind in [ind+1, ind+4, ind+3, ind+7]: # only P
                    con_row = rows[con_ind]
                    if not f'Net{con}' in con_row: # active sense line
                        con_row_split = (con_row.split()[2]).split('_')
                        flex = 'n/a' # dcb
                        load = 'n/a' # set explicitly for hybrids and dcbs
                        if 'JP' in con_row_split[0]: # hybrid
                            flex = bp_con_JP_to_alt(con_row_split[0], True)
                            load = con_row_split[2]
                            if 'WEST' in con_row: load += 'W'
                            if 'EAST' in con_row: load += 'E'
                        else: # dcb
                            load = f'{con_row_split[0][2:]}' # 1V5
                            # may as well take advantage of 'flex' to store redundant info
                            flex = con_row_split[2]
                            if '2V5' in con_row:
                                load = f'{con_row_split[0][2:]}-{con_row_split[1]}'
                                flex = con_row_split[3]
                        src = con_ind - ind
                        tbb_con_load[f'{con}_{src}-{twisted_ret(src)}'] = f'{flex}_{load}'
    return tbb_con_load


# Associate files with a function for parsing
parse_func = {}
files = [nominal, cavern, swap_pos, schem_ip, schem_mag, cavern_sense, tbb_schem] + compare
for file in files:
    if 'surface_LV_power_tests' in file: parse_func[file]=parse_surface
    elif 'LVR_PPP_Underground' in file: parse_func[file]=parse_cavern
    elif 'lvr_testing' in file: parse_func[file]=parse_cable_test # does nothing
    elif 'swap_positronic' in file: parse_func[file]=parse_swap_pos
    elif 'PEPI_' in file: parse_func[file]=parse_netlist
    elif 'CBM_LVR' in file: parse_func[file]=parse_check_petr_lvr
    elif 'CBM_PPP' in file: parse_func[file]=parse_check_petr_ppp # also nada
    elif 'CBI_LVR' in file: parse_func[file]=parse_check_petr_lvr
    elif 'CTM_LVR' in file: parse_func[file]=parse_check_petr_lvr
    elif 'CTI_LVR' in file: parse_func[file]=parse_check_petr_lvr
    elif 'underground_LVsense_layout' in file: parse_func[file]=parse_cavern_sense
    elif 'TelemetryBB_Mirror_FINAL_mpeco' in file: parse_func[file]=parse_tbb
    else:
        if file=='NA': continue
        print(f'\n\nDON\'T RECONGNIZE {file} FORMAT\n\n')


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

# outputs a list of rows to be printed for cctb testing tables. one sheet per
# PEPI. for A-side, use the comparable C-side PEPI lines
# order by BP first (gamma, beta, alpha), then by DCB/X hyb/S hyb, then by Pos
def organize_cctb_table(ppp_corrected_cavern_lines, z, truemir):
    lines = []
    for cl in ppp_corrected_cavern_lines:
        yz = z_truemir_to_y_z(z, truemir) # function is assuming C-side
        if cl.y == yz[0] and cl.z==yz[1]: lines.append(cl)
    rows = []
    for cl in lines:
        rows.append([cl.ppp, cl.ppp_pin, ppp_ret_pin(cl.ppp_pin), cl.bp,
                     cl.bp_con, cl.flex, cl.load, cl.get_lvr_pins(), cl.lvr,
                     cl.lvr_ch, cl.msa, '', '', '', '', ''])
    # order rows, add headers
    rows = sorted(rows, key=lambda r: r[1])
    rows_with_pos_refcol = [r+[int(r[0][1:])] for r in rows]
    rows = sorted(rows_with_pos_refcol, key=lambda r: r[-1])
    rows = [r[:-1] for r in rows] # drop refcol
    rows_with_flex_refcol = [r+[r[5][0].lower()] for r in rows]
    rows = sorted(rows_with_flex_refcol, key=lambda r: r[-1])
    rows = [r[:-1] for r in rows] # drop refcol
    rows = sorted(rows, key=lambda r: r[3], reverse=True)
    rows.insert(0, ['PPP Label', 'Pos. Src', 'Pos. Ret', 'Backplane',
                    'BP Con.', 'Flex Name', '4ASIC-group/DCB power', 'SBC Label',
                    'LVR Logical ID', 'LVR ch.', 'M/S/A', 'Connector on CCTB',
                    'Measured Voltage', 'Measured Current', 'Result', 'Comments'])
    return rows

# outputs a list of rows to be printed for cctb sense line testing tables. one
# sheet per true/mir PEPI type and per mag/IP (so 4 sheets total)
# order individual sheets by PPP connector
def organize_cctb_sense_table(senselines, truemir):
    rows_mag = []
    rows_ip = []
    res = []
    for sl in senselines:
        row = []
        if truemir=='True': ppp = sl.ppp_true
        elif truemir=='Mirror': ppp = sl.ppp_mir
        else: print('you formatted truemir wrong')
        row = [ppp, f' {sl.in_twistpair}', sl.in_spltr_lab(), sl.out_spltr_lab(),
               sl.lvr_con, sl.lvr,
               lvr_twistpair_to_ch(sl.lvr_con, sl.lvr_twistpair), '', '', '',
               '', '']
        if int(sl.lvr) <= 36: rows_mag.append(row)
        else: rows_ip.append(row)
    for rows in [rows_mag, rows_ip]:
        # order rows based on PPP connector (and twisted pair)
        rows = sorted(rows, key=lambda r: r[1])
        rows_with_ppp_refcol = [r+[int(r[0][1:])] for r in rows]
        rows = sorted(rows_with_ppp_refcol, key=lambda r: r[-1])
        rows = [r[:-1] for r in rows] # drop refcol
        rows_with_ppp_refcol = [r+[r[0][:1]] for r in rows]
        rows = sorted(rows_with_ppp_refcol, key=lambda r: r[-1])
        rows = [r[:-1] for r in rows] # drop refcol
        rows.insert(0, ['PPP Label', 'PPP Twisted Pair', 'Splitter Input',
                        'Splitter Output', 'LVR Con.', 'LVR Number', 'LVR ch.',
                        'M/S/A', 'Connector on CCTB', 'Measured Voltage',
                        'Result', 'Comments'])
        res.append(rows)
    return res

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
            print(f'\nFound cavern line with wrong PPP! '+
                  f'{cav_line.x+cav_line.y+cav_line.z+cav_line.bp}'+
                  f'{cav_line.bp_con+cav_line.ibbp2b2+cav_line.flex}'+
                  f'{cav_line.load+cav_line.msa}\n')
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
                print(f'\nFound (moved) cavern line with wrong PPP! '+
                      f'{cav_line.x+cav_line.y+cav_line.z+cav_line.bp}'+
                      f'{cav_line.bp_con+cav_line.ibbp2b2+cav_line.flex}'+
                      f'{cav_line.load+cav_line.msa}\n')
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
        # TODO comparison to LVR testing sheet. can skip Petr's PPP sorted sheet
        if comp in ['compare/lvr_testing.csv', 'compare/CBM_PPP_new.txt'] or \
            'alex' in comp: continue
        print(f'\n\nChecking {nominal} vs {comp}...\n\n')
        # For Petr comparison, he separates spliced lines into different
        # entries, so create a list with "unspliced" lines
        compare_lines = []
        for correct_line in ppp_corrected_cavern_lines.values():
            lvr_ch_split = correct_line.lvr_ch.split(' Y ')
            lvr_label_split = correct_line.lvr_label.split('   Y   ')
            for i in range(len(lvr_ch_split)):
                # make sure the ch number is an int string, not float...
                ch = str(int(float(lvr_ch_split[i])))
                split_correct_line = copy.deepcopy(correct_line)
                split_correct_line.set_lvr(correct_line.lvr, ch)
                split_correct_line.set_labels(correct_line.ppp_label,
                                              lvr_label_split[i])
                compare_lines.append(split_correct_line)
        # the actual checking will occur inside the respective parsing
        # functions! any errors will be printed there
        parse_func[comp](comp, compare_lines)


    # print(f'\n\nWriting (unformatted) fixed cavern mapping...\n\n')
    # fixed = 'fixed-'+cavern

    if ((args.doCheckLines).lower()=='true'):
        print(f'\n\nChecking {cavern} LVR<->load vs {schem_ip} and {schem_mag}...\n\n')
        # cavern_lines is what you want to compare
        # store erroroneous cav map/lvr schem line info, lvr info
        map_errors = [['Mag/IP', 'True/Mir', 'BP', 'M/S/A', 'LVR + Src Pin',
                      'Cav. Map. Load', 'LV Schem. Load']]
        ip_map_lvr_load = parse_func[schem_ip](schem_ip)
        mag_map_lvr_load = parse_func[schem_mag](schem_mag)
        map_lvr_load = {}
        for ip_lvr_load in ip_map_lvr_load:
            map_lvr_load[ip_lvr_load] = ip_map_lvr_load[ip_lvr_load]
        for mag_lvr_load in mag_map_lvr_load:
            map_lvr_load[mag_lvr_load] = mag_map_lvr_load[mag_lvr_load]
        # print(map_lvr_load)
        # go through the cavern_lines (including both lvrs for splices) and check
        # that lvr is mapping to the right load!
        for cav_line in cavern_lines:
            # by construction, you can just check that the lvr label is right
            lvr_labels = (cav_line.lvr_label).split('   Y   ')
            for lvr_label in lvr_labels:
                lvr = (lvr_label.split(' | '))[0][:-5] # ignore SBC and ret pin
                lvr = lvr.replace(' - ', '_') # reformat like netlist
                load = (lvr_label.split(' | '))[1]
                if not lvr in map_lvr_load:
                    print(f'Cannot find {lvr}!')
                    map_lvr_load[lvr] = 'Depopulated in LV Schem.!'
                if not load == map_lvr_load[lvr]:
                    # print(f'Cav. Map. {load} not equal to LV Schem. '+
                    #       f'{map_lvr_load[lvr]}: line '+
                    #       f'is {cav_line.x+cav_line.y}'+
                    #       f'{cav_line.z+cav_line.bp+cav_line.bp_con}'+
                    #       f'{cav_line.ibbp2b2+cav_line.flex+cav_line.load}'+
                    #       f'{cav_line.msa}, from LVR {lvr}')
                    map_errors.append([cav_line.z, true_mirror(cav_line.x,
                                       cav_line.y, cav_line.z),
                                       cav_line.bp, cav_line.msa, lvr, load,
                                       map_lvr_load[lvr]])
        fixme_map = open('fixme/lvr_load_mapping_errors.csv', 'w')
        writer_map = csv.writer(fixme_map)
        for row in map_errors: writer_map.writerow(row)
        fixme_map.close()

    # go through ppp_corrected_cavern_lines.value() and print the
    # columns Federico wants to a list; organize into different sheets for
    # each PEPI, then organize by BPs (gamma, beta, then alpha), then organize
    # DCBs then hybrids (straight then stereo), then order positronics small to
    # large
    for x in ['C', 'A']:
        for y in ['top', 'bot']:
            for z in ['ip', 'mag']:
                truemir = true_mirror(x,y,z)
                cctb_rows = organize_cctb_table(ppp_corrected_cavern_lines.values(),
                                                z, truemir)
                print(f'\nPrinting CCTB {x}-{y}-{z} power table...\n')
                cctb = open(f'output/{x}_{y}_{z}_{truemir}_LVpower_cctb.csv', 'w')
                cctb_writer = csv.writer(cctb)
                for row in cctb_rows: cctb_writer.writerow(row)
                cctb.close()

def cavern_sense_check():
    ip_map_lvr_load = parse_func[schem_ip](schem_ip)
    mag_map_lvr_load = parse_func[schem_mag](schem_mag)
    map_lvr_load = {}
    for ip_lvr_load in ip_map_lvr_load:
        map_lvr_load[ip_lvr_load] = ip_map_lvr_load[ip_lvr_load]
    for mag_lvr_load in mag_map_lvr_load:
        map_lvr_load[mag_lvr_load] = mag_map_lvr_load[mag_lvr_load]
    # print(map_lvr_load)
    tbb_map = parse_func[tbb_schem](tbb_schem)
    # print(tbb_map)
    senselines = parse_func[cavern_sense](cavern_sense, map_lvr_load)
    # for each sense line, check that the line the tBB claims is being sensed is
    # indeed the power line in Phoebe's map for the associated channel
    for sl in senselines:
        tbb_line = f'{sl.tbb_con}_{sl.in_twistpair}'
        bp = ((sl.in_label).split('_'))[0]
        if not tbb_line in tbb_map:
            print('\n\n!!! Couldnt find tBB {tbb_line}...\n\n')
            continue
        lvr_line = f'{sl.lvr}_{lvr_ch_to_pin(lvr_twistpair_to_ch(sl.lvr_con, sl.lvr_twistpair))}'
        if not lvr_line in map_lvr_load:
            print('\n\n!!! Couldnt find LVR {lvr_line}...\n\n')
            continue
        # print(f'My map+tBB: {bp}_{tbb_map[tbb_line]}, Phoebe power map: {map_lvr_load[lvr_line]}')
        lvr_load = load_label_phoebe_to_me(map_lvr_load[lvr_line])
        sense_load = f'{bp}_{tbb_map[tbb_line]}'
        if not lvr_load==sense_load:
            print(f'\nOn {lvr_line}, (Power) {lvr_load} != (Sense) {sense_load}\n')
    for x in ['C', 'A']:
        for y in ['top', 'bot']:
            for z in ['ip', 'mag']:
                truemir = true_mirror(x,y,z)
                cctb_rows = organize_cctb_sense_table(senselines, truemir)
                print(f'\nPrinting CCTB {x}-{y}-{z} sense table...\n')
                cctb = open(f'output/{x}_{y}_{z}_{truemir}_LVsense_cctb.csv', 'w')
                cctb_writer = csv.writer(cctb)
                if z=='mag': cctb_rows = cctb_rows[0]
                elif z=='ip': cctb_rows = cctb_rows[1]
                else: print('...')
                for row in cctb_rows: cctb_writer.writerow(row)
                cctb.close()

if cavern_sense == 'NA':
    no_typos = cavern_typo_check()
    if no_typos: cavern_check_fix()
else:
    cavern_sense_check()
