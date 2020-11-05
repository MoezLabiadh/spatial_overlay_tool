#-------------------------------------------------------------------------------
# Name:     FN referral tool

# Purpose: This tools assists BCTS Planners with FN referrals.
#          it generates an excel report showing overlay results
#          of blocks/roads and FN consultative areas.
#
# Author:  Moez Labiadh, BCTS-TKO
#
# Created: 21-10-2020
#-------------------------------------------------------------------------------
arcpy.AddMessage ('Importing modules...')


import os
import arcpy
import numpy as np
import pandas as pd
import xlsxwriter
from datetime import date

def initialize_tool (feature_type, features, ID_field):
    """Checks user inputs and cleans-up any temporary files
       from previous session"""
    arcpy.AddMessage ('Initializing the tool...')

    # Delete any temporary files from previous session
    arcpy.Delete_management("in_memory")
    if arcpy.Exists ('FN_areas_lyr'):
        arcpy.Delete_management("FN_areas_lyr")

    # Check user inputs and raise errors if any discrepancies.
    arcpy.AddMessage ('Checking user inputs...')
    desc = arcpy.Describe(features)
    geometryType = desc.shapeType
    spatialRef = desc.spatialReference

     #make sure the geometry type is correct
    if feature_type == 'Block':
        if not geometryType == 'Polygon':
            raise Exception ('Block features must be Polygons!')

    elif feature_type == 'Road':
        if not geometryType == 'Polyline':
            raise Exception ('Road features must be Lines!')
    else:
         raise Exception ('Input features must be either Blocks or Roads!')

    #make sure the layer is in projected coord system. for area/length calculations later on
    if feature_type == 'Block':
        if not spatialRef.linearUnitName == 'Meter':
            raise Exception ('Input layer must be in a projected coordinate system!')
    else:
        pass

    #make sure the input layer is not empty
    feat_count = arcpy.GetCount_management(features).getOutput(0)

    if int(feat_count) < 1:
        raise Exception ('Your input features layer is empty!')
    else:
        pass

    #make sure the ID/name field exists
    fields = [f.name for f in arcpy.ListFields(features)]

    if not ID_field in fields:
        raise Exception ('{} is not a field of {}' .format(ID_field, features))
    else:
        pass


def get_FN_areas ():
    """Returns FN Consultative areas intersecting with a BA boundaries
       This function runs ONLY inside the TKO Drilldown MXD"""
    arcpy.AddMessage ('Getting FN consultative areas...')
    # Set to current MXD
    mxd =  arcpy.mapping.MapDocument('CURRENT')
    # Get the TKO boundaries and FN terriotories layers from the MXD
    for lyr in arcpy.mapping.ListLayers(mxd):
        if lyr.name == 'LegalAreas':
            BA_layer = lyr
        elif lyr.name == 'FN Consultative Areas':
            FN_layer = lyr
        else:
            pass
    # Make a layer that has only FN territories in TKO
    FN_areas = 'FN_areas_lyr'
    arcpy.MakeFeatureLayer_management(FN_layer, FN_areas)
    arcpy.SelectLayerByLocation_management(FN_areas, 'intersect', BA_layer)

    return FN_areas


def Feature_FN_overlay (FN_areas, feature_type, features, field_team, op_area, ID_field, dem):
    """Returns a dictionnary containing the intersection results
       of blocks/roads against FN consultative areas"""
    arcpy.AddMessage ('Performing the analysis...')
    # Count of features in the input layer
    feat_count = arcpy.GetCount_management(features).getOutput(0)

    # Create a Dict that will hold the data
    val_dict = {}
    val_dict ['Type'] = []
    val_dict ['Field Team'] = []
    val_dict ['Op Area'] = []
    val_dict ['Name'] = []
    val_dict ['Elevation']= []

    if feature_type == 'Block':
        measure = 'Area (ha)'
        val_dict ['Type'].extend('Block' for i in range(int(feat_count)))
    elif feature_type == 'Road':
        measure = 'Length (m)'
        val_dict ['Type'].extend('Road' for i in range(int(feat_count)))

    val_dict [measure]= []

    # Get Field Team info
    arcpy.SpatialJoin_analysis(features, field_team, 'in_memory\st_inter', 'JOIN_ONE_TO_ONE')
    val_dict ['Field Team'] = [row[0] for row in arcpy.da.SearchCursor('in_memory\st_inter', ['FIELD_TEAM'])]

    # Get Op Area info
    arcpy.SpatialJoin_analysis(features, op_area, 'in_memory\op_area', 'JOIN_ONE_TO_ONE')
    val_dict ['Op Area'] = [row[0] for row in arcpy.da.SearchCursor('in_memory\op_area', ['OPAREA_NAM'])]

    # Initialize a counter
    proc_count = 1

    # Add the rest of data to the dict
    fields = [ID_field, "SHAPE@AREA", "SHAPE@Length", "SHAPE@XY"]
    sr = arcpy.Describe(dem).spatialReference
    cursor = arcpy.da.SearchCursor(features,fields,'', sr)
    for row in cursor:
        arcpy.AddMessage ('Processing feature {} of {}' .format (proc_count, feat_count))
        proc_count += 1
        # add Name, Type and Area/Length data for each block/road
        val_dict ['Name'].append (str(row[0]))
        if feature_type == 'Block':
            val_dict [measure].append ((round (row[1]/10000, 2)))
        elif feature_type == 'Road':
            val_dict [measure].append ((int(row[2])))

        # add Elevation data for each feature
          #get feature centroid pt
        pnt = arcpy.Point(row[3][0],row[3][1])
        ptGeometry = arcpy.PointGeometry(pnt)
          #clip TRIM DEM based on buffer around the centroid pt
        arcpy.Buffer_analysis(ptGeometry, 'in_memory\ptBuf', 100)
        extent = arcpy.Describe('in_memory\ptBuf').extent
        arcpy.Clip_management(dem, str(extent), 'in_memory\dem')
          #convert the dem to numpy array
        rast = arcpy.Raster('in_memory\dem')
        desc = arcpy.Describe(rast)
        ulx = desc.Extent.XMin
        uly = desc.Extent.YMax
        rstArray = arcpy.RasterToNumPyArray(rast)
         #get the row/col position of the centroid pt
        deltaX = pnt.X - ulx
        deltaY = uly- pnt.Y
        arow = int(deltaY/rast.meanCellHeight)
        acol = int(deltaX/rast.meanCellWidth)
         #extract the elevation from array and add it to the Dict
        elevation = rstArray[arow,acol]
        val_dict ['Elevation'].append(elevation)

        arcpy.Delete_management("in_memory")

    # Spatial Join of feaures and FN territories
    arcpy.AddMessage ('Populating FN overlay results...')
    intersect = 'in_memory\intersect'
    arcpy.SpatialJoin_analysis(FN_areas, features, intersect, 'JOIN_ONE_TO_MANY')

    # Add referral requirement of each feature to the dict based on the spatial overlay
    for row in arcpy.da.SearchCursor(intersect, ['CONTACT_ORGANIZATION_NAME', ID_field]):
      if str(row[0]) not in val_dict:
          val_dict[str(row[0])] = []
          val_dict[str(row[0])].extend('n/r' for i in range(int(feat_count)))

      for i in range(0, int(feat_count)):
           if str(row[1]) == val_dict ['Name'][i]:
                val_dict[str(row[0])][i] = 'required'
           else:
                pass

    #print ({k:len(v) for k, v in val_dict.items()})

    return val_dict


def make_excel_report (feature_type, val_dict, out_excel):
    """Outputs an Excel report based on the overlay results Dict"""
    arcpy.AddMessage ('Generating the Excel report...')
    # Convert the dictionnary to a pandas dataframe
    df = pd.DataFrame.from_dict(val_dict)

    # Make sure the columns appear in the desired order
    if feature_type == 'Block':
        first_cols = ['Type', 'Field Team', 'Op Area', 'Name', 'Area (ha)', 'Elevation']
    elif feature_type == 'Road':
        first_cols = ['Type', 'Field Team', 'Op Area', 'Name', 'Length (m)', 'Elevation']

    all_cols = first_cols + (df.columns.drop(first_cols).tolist())
    df = df[all_cols]

    # Sort entries by name
    df.sort_values(by=['Name'], inplace=True)
    df = df.reset_index(drop=True)

    # Make the index starts at 1 instead of 0
    df.index = df.index + 1
    df.index.name = '#'

    #print (df.head())

    # remove the default Pandas header formatting
    pd.formats.format.header_style = None

    # Export to excel
    writer = pd.ExcelWriter(out_excel, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='FN_ref_report')

    # Make the excel look nice!
    workbook = writer.book
    worksheet = writer.sheets['FN_ref_report']
    worksheet.set_zoom(90)

    #get the number of rows and colums
    rows = df.shape[0]
    cols = df.shape[1]

    #create formats
    format_all = workbook.add_format({'border':1, 'text_wrap': True})
    format_header =  workbook.add_format({'bold': True, 'text_wrap': True})
    format_header.set_align('vcenter')
    format_header.set_align('center')
    format_Y = workbook.add_format({'bg_color':'#F0FFF0'})
    format_N = workbook.add_format({'bg_color':'#FFE6E6'})
    format_YN = workbook.add_format()
    format_YN.set_align('center')


    worksheet.conditional_format( 0,0,rows,0,{'type':'no_blanks',
                                              'format':format_header})

    worksheet.conditional_format( 0,0,rows,cols,{'type':'no_blanks',
                                                 'format':format_all,})
    worksheet.conditional_format( 0,0,rows,cols,{'type':'blanks',
                                                 'format':format_all})

    worksheet.conditional_format( 0,0,rows,cols,{'type':'cell',
                                                 'criteria': 'equal to',
                                                 'value': '"required"',
                                                 'format':format_Y})
    worksheet.conditional_format( 0,0,rows,cols,{'type':'cell',
                                                 'criteria': 'equal to',
                                                 'value': '"n/r"',
                                                 'format':format_N})

    # set column/row width and format
    worksheet.set_row (0, None, format_header)
    worksheet.set_column (0,0, None, format_header)
    worksheet.set_column(0, 1, 9)
    worksheet.set_column(2, 4, 14)
    worksheet.set_column(5, 6, 12)
    worksheet.set_column (7,cols, 14, format_YN)

    # Add report date
    today = date.today().strftime("%B %d, %Y")
    worksheet.write(rows+3, 1, 'Report generated on: {}'.format(today))

    # Add a 'Notes" sheet
    notes = workbook.add_worksheet('Notes')
    notes.set_column(0, 0, 110)
    format_title = workbook.add_format({'bold': True, 'font_size':16})
    format_text = workbook.add_format({'text_wrap': True})
    notes.write(0, 0, 'The FN referral tool', format_title )

    text_1 = ('This tool is intended to assist Planners with FN referrals. \n'
            'The spatial overlays of input features (blocks or roads) and FN consultative areas are reported in the first sheet. \n')

    text_2=  ('Please take note of the following when using the tool: \n'
              '1- The tool must be executed inside the Drilldown tool MXD \n'
              '2- The Profiles of Indigenous Peoples (PIP) layer is used as FN linework in this analysis \n'
              '3- Block centroid points (mid-points for roads) are used to derive Elevation data \n'
              '4- Features with empty Name/ID field are showns as "None" in the report \n'
              '5- Features outside of BCTS TKO operating Areas will have empty "Op Area" fields')

    notes.write(1, 0, text_1, format_text)
    notes.write(2, 0, text_2, format_text)

    writer.save()
    print ('Excel report saved at: {}' .format(out_excel))


def main():
    """ Runs the tool"""
    arcpy.overwriteOutput = True

    # User inputs
    feature_type = arcpy.GetParameterAsText(0)
    features = arcpy.GetParameterAsText(1)
    ID_field = arcpy.GetParameterAsText(2)
    out_excel = arcpy.GetParameterAsText(3)

    # Additonal layers needed for the analysis
    tko_data_loc = r'\...\Business_Area'
    field_team = os.path.join(tko_data_loc, 'BusinessArea.shp')
    op_area = os.path.join(tko_data_loc, 'BCTS_TKO_OA2020.shp')
    dem = r'\...\bc_elevation_25m_bcalb.tif'

    # Run the functions
    initialize_tool (feature_type, features, ID_field)
    FN_areas = get_FN_areas ()
    val_dict = Feature_FN_overlay(FN_areas,feature_type, features, field_team, op_area, ID_field, dem)
    make_excel_report (feature_type, val_dict, out_excel)

    # Delete temporary files
    arcpy.Delete_management("in_memory")
    if arcpy.Exists ('FN_areas_lyr'):
        arcpy.Delete_management("FN_areas_lyr")

    arcpy.AddMessage ('Processing Completed!')

main()
