# Common-Excel-VBP-Unused-Public-Items-Service
 Supplements MZ-Tools by displaying all unused Public items in a selected Workbook's VB-Project - which is not covered by MZ-Tools's dead code analysis.
 
 ## Usage
 1. Download and open the [VBPunusedPublic.xlsb][1] Workbook
 2. Select a Workbook (will be opened when not already open)
 3. Decide from the list of displayed VBComponents which to exclude or include in the analysis
 4. Start the analysis
 
 Since the analysis may take a couple of seconds, the progress is displayed in the Application.StatusBar

## The service considers, recognizes, copes with:
- Code lines continued
- Code lines with multiple sub-lines (separated by ': ')
- Project public, VBComponent global and Procedure local Class Instances
- Public items:
  - Constants
  - Variables
  - Sub-Procedures
  - Functions
  - Properties
  - Methods (Function, Sub in Class-Modules)

 [1]:https://gitcdn.link/cdn/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service/master/VBPunusedPublic.xlsb
