# Spine-Density-HD

These macros are meant to be run on files generated by NeuronStudio. Master data file is meant to start with only one page.  
  
Raw data file naming convention: ##\_(2-character animal ID)\_(3-character brain region)(1-character slice number)\_C(1-2-digit cell number)\_(P or D is optional)(segment number)(optional \_(L or R))  
  
AverageDensityAndHD will generate Average Density and Head Diameter (HD) for Thin and Mushroom spines.  
  
AllSpinesHDPart1 will pull all of the HD's for each spine from the raw data files into one sheet.  
AllSpinesHDPart2 will compile/merge the HD's for each cell once you've processed all of the raw data with AllSpinesHDPart1.  
  
ProcessingMasterSplitSpines will separate the data in the master sheet for the Mushroom and Thin Spines by cell and by region.
