# Common-Excel-VBP-Unused-Public-Items-Service
 Supplements MZ-Tools by displaying all unused Public items in a selected Workbook's VB-Project - which is not covered by MZ-Tools's dead code analysis.
 
 ## Usage
 1. Download and open the [VBPunusedPublic.xlsb][1] Workbook
 2. Select a Workbook (will be opened when not already open)
 3. Decide from the list of displayed VBComponents which to exclude or include in the analysis
 4. In any component add (the specified excluded components and code lines are just an example) 
 
 ```vb
 Private Sub CheckForUnusedPublicItems()
' ----------------------------------------------------------------
' Attention! The service requires the VBPUnusedPublic.xlsb
'            Workbook open. When not open the service terminates
'            without notice.
' ----------------------------------------------------------------
    Const COMPS_EXCLUDED = "clsQ,mRng,fMsg,mBasic,mDct,mErH,mFso,mMsg,mNme,mReg,mTrc,mWbk,mWsh"
    Const LINES_EXCLUDED = "Select Case ErrMsg(ErrSrc(PROC))" & vbLf & _
                           "Case vbResume:*Stop: Resume" & vbLf & _
                           "Case Else:*GoTo xt"
    On Error Resume Next
    '~~ Providing all (optional) arguments saves the Workbook selection dialog and the VBComponents selection dialog
    Application.Run "VBPUnusedPublic.xlsb!mUnused.Unused", ThisWorkbook, COMPS_EXCLUDED, LINES_EXCLUDED

End Sub
```
 
 Since the analysis may take a couple of seconds, the progress is displayed in the Application.StatusBar

## The service covers the Public items
- Constants
- Variables
- Sub-Procedures
- Functions
- Class Instances
- Properties (Get, Let, Set)
- Methods (Function, Sub in Class-Modules)
  
## The service considers, recognizes, copes with:
- Code lines continued
- Code lines with multiple sub-lines (separated by ': ')
- Class Instances (Project public, VBComponent global, and Procedure local)
- OnAction in Worksheet Controls
- Nested With End With instructions for Class Instances

## Performance
Though the performance shouldn't really matter for a service used only occasionally:
```
 12 Thousand analyzed code lines (2 thousand skipped) in
692 Procedures had been analyzed for 
277 Public declared items in
 10 Seconds
```
The result was 3 unused and 274 used Public declared items.

 [1]:https://gitcdn.link/cdn/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service/master/VBPunusedPublic.xlsb
