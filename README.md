# AeroToolKitAddIn

Excel add-in with common aeronautical engineering functions, including standard atmosphere, QNH, QFE, altimetry, obstacle clearance, airspeeds, climb gradients, turns and common aeronautical unit conversions. The atmospheric model is based on the ICAO standard (as documented in ICAO doc 7488, dated 1993). Altitude inputs up to the stratopause (51 km or 167,000 ft) are permitted. Non-standard atmospheric characteristics are available via ISA deviations.

For reference, the ICAO standard atmosphere, the ISA standard atmosphere, and the US 1962 and 1976 standard atmospheres are identical to 32 km (105,000 ft).

The add-in uses an intuitive function naming convention. All function names start with the "aero" prefix, and are followed by a description of the function, starting with the output of interest. Underbars ("_") are used in function names to highlight the units output from functions. The lower-case "f" is used to denote "function of" and is followed by one or more independent variables. As an example, the function AeroQ_lbfPerFt2_fHpKtasIsaDevCelsius computes dynamic pressure ("Q") in units of pounds per square foot ("lbfPerFt2") as output, using true airspeed in knots ("Ktas") and temperature deviation from the standard atmosphere in degrees Celsius ("IsaDevCelsius") as input. Similarly, functions that perform unit conversion have function names that begin with "aeroConv", followed by a "before" unit and an "after" unit separated by "To". As an example, the conversion function "aeroConvFtToMeter" converts feet as input to meters as output.

The naming convention takes advantage of Excel's "intellisense", which narrows function choices as the user types. For example, typing "=aero" causes a list of all aeroToolKit functions to appear. Further typing narrows the list of available functions. For instance, typing "=aeroConv" limits the list to the conversion functions. With a little typing/guessing of function names, and use of arrow keys to navigate the filtered function lists, users can narrow-in on functions of interest quickly. This eliminates looking-up, memorizing or typing entire function names (however, for reference, a complete list of functions is provided (AeroToolKitFunctionList.xlsx), organized by type, inputs, outputs, and description). 

To keep function names compact, units are sometimes omitted. For reference the default unit for length is feet. Refer to the list of functions for a complete description.

To install the .xlam file:
Click on "File" | "Options" | "Add-ins", and then "Go..." (next to "Manage: Excel Add-ins"). Click on "Browse" and navigate to the location of the add-in on your file system. Make sure this is a stable location where you intend to keep the add-in.

NOTE! Excel may have a problem "remembering" add-ins that it has loaded upon each re-start. To prevent this, do the following two things after installation:
1) Locate the add-in file in Windows Explorer, right-click on it and scroll down to select "Properties". On the bottom of the “Properties” window, check "Unblock" (if such an option is available).
2) In Excel, click on "File" | "Options". On the left-hand side of the window, click on "Trust Center" and then "Trust Center Settings". In the “Trust Center” window, click on “Add new location” or "Trusted Locations" and select the folder containing the add-in.

If Excel "forgets" the add-in anyway, go to the add-ins menu ("File" | "Options" | "Add-ins"), select "Manage Excel Add-ins", uncheck the Aerotoolkit add-in and click "Ok". Then follow the same steps to re-check and activate the add-in. 

The AeroToolKit.xlam add-in, which is a binary file, consists of a single module, AeroToolKit.bas, which is a text file of VBA source code. 

Other essential elements of the repo include the function list with descriptions (AeroToolKitFunctionList.xlsx), and a full test of all functions with expected results (AeroToolKitTest.xlsm).

Please observe the following procedures when updating the code:  
1) Follow a workflow in which the AeroToolKit.bas file is updated (via text editor) and imported to the AeroToolKit.xlam file. If changes are made to AeroToolKit from the Development environment in Excel (i.e., via the Visual Basic editor) and saved directly to the .xlam without export to the .bas file (a plausible workflow), there will be no record of changes. A workflow that exports the module from Excel would result in a trackable text file, but does not force the discipline of creating the text file, and invites a disconnect between the .xlam and text file. It is not recommended. 
2) The function list and descriptions in AeroToolKitFunctionList.xlsx must reflect all changes. The .xlsx file is binary, so for traceability, use Excel to export a copy of the updated list to AeroToolKitFunctionList.csv. 
3) Create and evaluate new tests of the function, as needed, in AeroToolKitTest.xlsm.

Users concerned with potential security risks associated with an .xlam file from the internet can create their own add-in and import the plain-text AeroToolKit.bas. Particularly nervous users can review the source for AeroToolKit.bas online and create their own text .bas file using copy and paste.

This add-in derives from an earlier add-in by the same original author, stdAtmo20140121. This new add-in standardizes naming conventions and adds functionality.