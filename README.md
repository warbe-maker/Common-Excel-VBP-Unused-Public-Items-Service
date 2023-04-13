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
Since the service will be used only occasionally the performance shouldn't really matter. The performance mainly depends on the number of detected unused items and appears acceptable. <br>
### Example 
```
~12 Thousand analyzed code lines (~2 thousand skipped) in
687 Procedures had been analyzed for 
283 Public declared items within
  5 Seconds for a result of
  0 of 283 items detected unused
```
See also the below execution trace.
```
23-04-12 17:13:09 Execution trace by 'Common VBA Execution Trace Service' (https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service)
23-04-12 17:13:09                 >> Begin execution trace 
23-04-12 17:13:09 00,0000         |  >> mUnusedPublicTest.Test_UnusedPublic
23-04-12 17:13:09 00,0011         |  |  >> mUnused.Unused
23-04-12 17:13:09 00,0026         |  |  |  >> mComps.Collect
23-04-12 17:13:09 00,0036 00,0009 |  |  |  << mComps.Collect
23-04-12 17:13:09 00,0040         |  |  |  >> mProc.Collect
23-04-12 17:13:09 00,1672 00,1633 |  |  |  << mProc.Collect
23-04-12 17:13:09 00,1742         |  |  |  >> mClass.CollectInstncsCompGlobal
23-04-12 17:13:09 00,1813 00,0071 |  |  |  << mClass.CollectInstncsCompGlobal
23-04-12 17:13:09 00,1824         |  |  |  >> mClass.CollectInstncsProcLocal
23-04-12 17:13:10 00,3409 00,1586 |  |  |  << mClass.CollectInstncsProcLocal
23-04-12 17:13:10 00,3424         |  |  |  >> mItems.CollectPublicItems
23-04-12 17:13:10 00,3458 00,0035 |  |  |  << mItems.CollectPublicItems
23-04-12 17:13:10 00,3461         |  |  |  >> mItems.CollectPublicUsage
23-04-12 17:13:14 04,8192 04,4730 |  |  |  << mItems.CollectPublicUsage
23-04-12 17:13:14 04,9265 04,9254 |  |  << mUnused.Unused
23-04-12 17:13:14 04,9270 04,9270 |  << mUnusedPublicTest.Test_UnusedPublic
23-04-12 17:13:14 04,9270 04,9270 << End execution trace 
23-04-12 17:13:14                 Impact on the overall performance (caused by the trace itself): 00,0005 seconds!
23-04-12 17:13:14 Execution trace by 'Common VBA Execution Trace Service' (https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service)
```
### Approach
All required information is collected first which 200 to 300 msec. Subsequently each code line is checked once for any of the public declared items being used. Any used item detected is removed from the public items and moved to a used collection. I.e. that the loop for a code line over all (remaining) public items becomes faster with each one detected used.

 [1]:https://gitcdn.link/cdn/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service/master/VBPunusedPublic.xlsb
