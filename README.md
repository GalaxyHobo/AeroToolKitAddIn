# AeroToolKitAddIn
Excel add-in with functions to perform common aeronautical engineering calculations, including standard atmosphere, airspeed, aircraft performance and unit conversion calculations. Includes non-standard day calculations (ISA deviation). The add-in is written in VBA embedded in AeroToolKit.xlam.  

Atmospheric calculations are based on the ICAO standard atmosphere (reference ICAO doc 7488, dated 1993).  

This add-in derives from an earlier add-in by the same original author, stdAtmo20140121. This new add-in standardizes naming conventions and adds functionality.  

To install the .xlam file, go to File | Options | Add-ins, and navigate to the location at which you wish to keep the add-in on your file system.   NOTE! There seems to be an issue with Excel "remembering" add-ins that it has loaded, particularly with Excel 2019. If this is a problem, go to the Add-ins menu (File | Options | Add-ins), select Manage Excel Add-ins, and uncheck the Aerotoolkit add-in and click Ok. Then return to the menu and re-check the add-in. 

The .xlam add-in consists of a single module, AeroToolKit, which is exported from the .xlam to the text file AeroToolKit.bas and maintained in the repo. 

Other essential elements of the repo include the function list with descriptions, AeroToolKitFunctionList.xlsx, and a full test of all functions (including expected results), AeroToolKitTest.xlsm.  

For configuration management, observe the following procedures:  
1) Changes to the functions in the AeroToolKit.xlam file must be exported to the text AeroToolKit.bas file.  
2) New functions or changes to existing functions must be accounted in an update to AeroToolKitFunctionList.xlsx. Export a copy of the file to the equivalent AeroToolKitFunctionList.csv in Excel. 
3) Create new tests of the function in AeroToolKitTest.xlsm.  

Users concerned with potential security risks associated of downloading and using an .xlam file from the internet can create their own add-in and import AeroToolKit.bas. Particularly nervous users can review the source for AeroToolKit.bas online and create their own text .bas file via copy and paste.
