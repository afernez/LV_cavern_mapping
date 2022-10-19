"""
Parse the input cavern mapping
"""
import pandas, os, fnmatch
import warnings # pandas FutureWarnings are annoying...
warnings.simplefilter(action='ignore', category=FutureWarning)
from argparse import ArgumentParser

### grab command line args
parser = ArgumentParser(description='Produce computer-readable cavern mappings')
parser.add_argument('mapping', help='specify cavern mapping to be used as input')
args = parser.parse_args()

### global variables in script
dfDCBs = {}
dfHybrids = {}

pepiType = {}
pepiType['IP/CB'] = 'True'
pepiType['IP/CT'] = 'Mirror'
pepiType['Mag/CB'] = 'Mirror'
pepiType['Mag/CT'] = 'True'
pepiType['IP/AB'] = 'Mirror'
pepiType['IP/AT'] = 'True'
pepiType['Mag/AB'] = 'True'
pepiType['Mag/AT'] = 'Mirror'

# pandas.set_option('max_rows', 2000 ) # deprecated
pandas.set_option("expand_frame_repr", False)
## suppresses the following warning:
## A value is trying to be set on a copy of a slice from a DataFrame.
pandas.options.mode.chained_assignment = None

# set input mapping file
fileIn = args.mapping
xls = pandas.ExcelFile( fileIn )
sheets = xls.sheet_names
#sheets.sort(reverse=True)

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



#dfDCBs = {}
#dfDCBs['DCB - IP - True'] = parseDCBs( xls, 'DCB - IP - True' )
#dfDCBs['DCB - IP - Mirror'] = parseDCBs( xls, 'DCB - IP - Mirror' )
#dfDCBs['DCB - Mag - True'] = parseDCBs( xls, 'DCB - Mag - True' )
#dfDCBs['DCB - Mag - Mirror'] = parseDCBs( xls, 'DCB - Mag - Mirror' )

dcbs = fnmatch.filter( sheets, "DCB - *" )
for dcb in dcbs:
    print( dcb )
    dfDCBs[dcb] = parseDCBs( xls, dcb )


hybrids = fnmatch.filter( sheets, "Hybrid - *" )
for hybrid in hybrids:
    print(hybrid)
    dfHybrids[hybrid] = parseHybrids( xls, hybrid )

#
# Uses the global maps which is not ideal...
#
def merge( dcb, hybridX, hybridS ):
    print( "Merging...", dcb, hybridX, hybridS )
    dfMergeHybrid = dfHybrids[hybridX]
    dfMergeHybrid = dfMergeHybrid.append( dfHybrids[hybridS] )
    dfMergeHybrid['Pos'] = dfMergeHybrid['BP Connector'].str.replace( 'JP', '' ).astype(int)
    dfMergeHybrid = dfMergeHybrid.sort_values( by=['Pos', '4-asic group' ] )#, 'LVR Name'] )
    cols = list(dfMergeHybrid.columns)
    cols[cols.index('4-asic group')] = '4-asic group / DCB power'
    dfMergeHybrid.columns = cols
    del dfMergeHybrid['Pos']
    dfOut = dfDCBs[dcb][['LVR Name', 'PPP Name', 'BP Connector', 'iBB/P2B2 Connector', 'SBC FLEX NAME', 'Voltage', 'M/S/A', 'PPP Positronic', 'PPP Src/Ret', 'LVR ID', 'LVR Channel' ]]
    cols = list(dfOut.columns)
    cols[cols.index('Voltage')] = '4-asic group / DCB power'
    dfOut.columns = cols
    dfOut = dfOut.append( dfMergeHybrid[['LVR Name', 'PPP Name', 'BP Connector', 'iBB/P2B2 Connector', 'SBC FLEX NAME', '4-asic group / DCB power', 'M/S/A', 'PPP Positronic', 'PPP Src/Ret', 'LVR ID', 'LVR Channel' ]] )
    dfOut['LVR Crate'] = 'TBD'
    return dfOut

def mergePEPI( pepi ):
    ds = pepi.split('/')
    station = ds[0]
    quadrant = ds[1]
    type = pepiType[pepi]
    dcb = 'DCB - ' + station + ' - ' + type
    hybridX = 'Hybrid - ' + station + ' - ' + type + ' - Straight'
    hybridS = 'Hybrid - ' + station + ' - ' + type + ' - Stereo'
    print( 'Merging: ' + pepi, dcb, hybridX, hybridS )
    return merge( dcb, hybridX, hybridS )

## merge files and write output

# TODO: should edit directory structure so that these files get saved to an "output" folder
writerC = pandas.ExcelWriter('output/Cavern_LV_Mapping_MT_Formatting_C_side.xlsx', engine='xlsxwriter')
writerA = pandas.ExcelWriter('output/Cavern_LV_Mapping_MT_Formatting_A_side.xlsx', engine='xlsxwriter')

#dfPEPI = {}
for pepi in sorted(pepiType):
    dfPEPI = mergePEPI( pepi )
    for backplane in [ 'alpha', 'beta', 'gamma' ]:
        if pepi.find( '/C') > -1:
            dfPEPI[dfPEPI['LVR Name'].str.contains(backplane)].to_excel( writerC, sheet_name=pepi.replace( "/", '-' ) + '-' + backplane, index=False )
        elif pepi.find( '/A') > -1:
            dfPEPI[dfPEPI['LVR Name'].str.contains(backplane)].to_excel( writerA, sheet_name=pepi.replace( "/", '-' ) + '-' + backplane, index=False )

def format_columns(writer, hide ):
    cell_format = writer.book.add_format()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    for name in writer.sheets:
        worksheet = writer.sheets[name]
        ## centre cells
        worksheet.set_column( "A:B", 34, cell_format, {'hidden':hide} )
        worksheet.set_column( "C:C", 14, cell_format )
        worksheet.set_column( "D:D", 20, cell_format )
        worksheet.set_column( "E:E", 14, cell_format )
        worksheet.set_column( "F:F", 25, cell_format )
        worksheet.set_column( "G:G", 8, cell_format )
        worksheet.set_column( "H:I", 15, cell_format )
        worksheet.set_column( "J:J", 8, cell_format )
        worksheet.set_column( "K:K", 14, cell_format )

    return writer

writerC = format_columns( writerC, False )
writerC.save()
writerA = format_columns( writerA, False )
writerA.save()
