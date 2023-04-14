# Common-Excel-VBP-Unused-Public-Items-Service
 Supplements MZ-Tools by displaying all unused Public items in a selected Workbook's VB-Project - which is not covered by MZ-Tools's dead code analysis.
 
 ## Usage
 1. Download and open the [VBPunusedPublic.xlsb][1] Workbook
 2. Copy the below code into any of the to-be-analyzed Workbooks components and execute it directly from within the IDE
 
 ```vb
 Private Sub UnusedPublicItems()
' ----------------------------------------------------------------
' Please note:  The service displays the result by means of 
'               ShellRun. In case no application is linked with
'               the file extension .txt a dialog to determine
'               which application to use for the open will be
'               displayed.
'
' W. Rauschenberger, Berlin Apr 2023
' ----------------------------------------------------------------
    Const UNUSED_SERVICE As String = "VBPunusedPublic.xlsb!mUnused.Unused"  ' must not be altered
    Const COMPS_EXCLUDED As String = vbNullString                           ' Example: "mBasic,mDct,mErH,mObject,mTrc"
    Const LINES_EXCLUDED As String = "Select Case*ErrMsg(ErrSrc(PROC))" & vbCrLf & _
                                        "Case vbResume:*Stop:*Resume" & vbCrLf & _
                                        "Case Else:*GoTo xt"
    
    
    
    '~~ Check if the servicing Workbook is open and terminate of not.
    Dim wbk As Workbook
    On Error Resume Next
    Set wbk = Application.Workbooks("VBPunusedPublic.xlsb")
    If Err.Number <> 0 Then
        MsgBox Title:="The Workbook VBPunusedPublic.xlsb is not open!", Prompt:="The Workbook needs to be opened before this procedure is re-executed." & vbLf & vbLf & _
                      "The Workbook may be downloaded from the link provided in the 'Immediate Window'. Use the download button on the displayed webpage."
        Debug.Print "https://github.com/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service/blob/main/VBPunusedPublic.xlsb?raw=true"
        Exit Sub
    End If
    
    Application.Run UNUSED_SERVICE, ThisWorkbook , COMPS_EXCLUDED, LINES_EXCLUDED

End Sub
```

## Application.Run arguments

| Argument        | Description |
|-----------------|-------------|
|_UNUSED\_SERVICE_| Name of the servicing Workbook and the called procedure. Must not be altered except when the Workbook's name and/or the procedure is changed. |
|_ThisWorkbook_   | Optional, a Workbook expression. ThisWorkbook when the procedure is dedicated for executing the service for the Workbook. When the argument is omitted (`Application.Run UNUSED_SERVICE, , COMPS_EXCLUDED, LINES_EXCLUDED`, the procedure will become universal in the sense that the service will display a dialog for the selection of a Workbook - which will be opened when not already open. |
|_COMPS\_EXCLUDED_| Optional, a comma delimited string specifying VBComponents for being excluded from the analysis. Excluded may be (for example) _Common Components_ of which by nature only a few public services will be used. When no VBComponents are to be excluded a _vbNullString_ must be provided. When the argument is omitted a dialog for the selection of included/excluded VBComponents is displayed. |
|_LINES\_EXCLUDED_| Optional, string expression, lines delimited by _vbCrLf_. The provision of project specific or standard code lines which for sure will not contain any Public item may increase the performance. |
 
## Public items covered/recognized by the analysis
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
- Class Instances (Project scope, VBComponent scope, and Procedure scope)
- OnAction in Worksheet Controls (Command-Buttons)
- Nested `With` ... `End With` instructions for Class Instances

## Performance
The service removes a Public item from the list when it is detected used. This reduces the list of the to-be-checked Public items along with the analysis and thus the performance mainly depends on the number of the remaining unused Public items. Because the service may take a couple of (likely less than 10) seconds, the progress is displayed in the _Application.StatusBar_. Since just used occasionally, the performance of the service shouldn't really matter. However ...  <br>
**Example:** 
```
~12 Thousand analyzed code lines (~2 thousand skipped) in
687 Procedures had been analyzed for 
283 Public declared items within
 ~5 Seconds for a result of
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

 [1]:https://github.com/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service/blob/main/VBPunusedPublic.xlsb?raw=true
 
