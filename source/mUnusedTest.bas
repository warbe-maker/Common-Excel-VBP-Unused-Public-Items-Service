Attribute VB_Name = "mUnusedTest"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mUnusedTest: Test of all services of the module.
'
' ----------------------------------------------------------------

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mUnusedPublicTest" & "." & e_proc
End Function

Private Sub Test_UnusedPublic()
' ----------------------------------------------------------------
' For testing the service analyzes its own VBProject by means of
' the proposed procedure
' See https://github.com/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service
' ----------------------------------------------------------------
    Const COMPS_EXCLUDED As String = "fMsg,mBasic,mErH,mMsg,mTrc,mCompManClient"
    Const LINES_EXCLUDED As String = "Select Case*ErrMsg(ErrSrc(PROC))" & vbCrLf & _
                                     "Case vbResume:*Stop:*Resume" & vbCrLf & _
                                     "Case Else:*GoTo xt"
    Const UNUSED_SERVICE As String = "VBPunusedPublic.xlsb!mUnused.Unused" ' must not be altered
      
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
    
    Application.Run UNUSED_SERVICE, ThisWorkbook, COMPS_EXCLUDED, LINES_EXCLUDED

End Sub
