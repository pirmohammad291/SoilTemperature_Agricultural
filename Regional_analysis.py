import arcpy
import os
from arcpy.sa import*
from arcpy import env
from xlwt import Workbook, Formula
import xlrd
import numpy as np

book = Workbook()
sheet1 = book.add_sheet("Raw")
##sheet1.write(0,0, "File Name")
##sheet1.write(0,1, 'Min') 
##sheet1.write(0,2, 'Max')
##sheet1.write(0,1, 'Mean')
##sheet1.write(0,4, 'STD')
arcpy.CheckOutExtension("Spatial")
arcpy.env.overwriteOutput = True
InputFolder = r"C:\Users\pirmo\OneDrive - The Hong Kong Polytechnic University\LUMIP\land-noLu\tas\Mean_Model\\"
arcpy.env.workspace = InputFolder
OutputFolder= InputFolder+"\Regions\\"
if not os.path.isdir(OutputFolder):
    os.mkdir(OutputFolder)              # If output folder doesnt exist, create       
CitySquareShape = r"C:\Users\pirmo\OneDrive - The Hong Kong Polytechnic University\LUMIP\Region21\Regions\\"

cities = []
for dir_entry in os.listdir(CitySquareShape):
    cities.append(os.path.join(CitySquareShape, dir_entry))
rasters = []
for dir_entry in os.listdir(InputFolder):
    rasters.append(os.path.join(InputFolder, dir_entry))
    
cities = [x for x in cities if ".shp" in x[-4:]]
filenames = [x.split("\\")[-1].split('.')[0] for x in cities]
rasters = [x for x in rasters if ".tif" in x[-4:]]

col=1
for city, filename in zip(cities, filenames):
    cityOutputPath = os.path.join(OutputFolder, filename)   # Current folder where results should be saved   
    if not os.path.isdir(cityOutputPath):
        os.mkdir(cityOutputPath)              # If output folder doesnt exist, create        
##    extent = arcpy.Describe(city).extent
##    west = extent.XMin
##    south = extent.YMin
##    east = extent.XMax
##    north = extent.YMax
    print ("[+] City: "+filename)
    sheet1.write(0,col, filename)
    ex_row = 1
    for raster in rasters:
            rasterName = raster.split("\\")[-1]
            if not os.path.exists(os.path.join(cityOutputPath, filename+"_"+rasterName)):
                cityOutput = os.path.join(cityOutputPath, rasterName)
##              nodata = arcpy.Describe(arcpy.Raster(raster)).noDataValue
                print ("[+] Clipping: " + rasterName)
##              arcpy.Clip_management(raster, "{} {} {} {}".format(west, south, east, north) , cityOutput, city, nodata, "ClippingGeometry", "NO_MAINTAIN_EXTENT")
                arcpy.gp.ExtractByMask_sa(raster, city, cityOutput)
##              sheet1.write(ex_row, 0, rasterName)
##              sheet1.write(ex_row, 1, arcpy.GetRasterProperties_management(raster, "MINIMUM", "").getOutput(0))
##              sheet1.write(ex_row, 2, arcpy.GetRasterProperties_management(raster, "MAXIMUM", "").getOutput(0))
                sheet1.write(ex_row, col, arcpy.GetRasterProperties_management(cityOutput, "MEAN", "").getOutput(0))
##              sheet1.write(ex_row, 4, arcpy.GetRasterProperties_management(raster, "STD", "").getOutput(0))
            ex_row = ex_row + 1
    col = col + 1          
book.save (OutputFolder+"Stats.xls")

