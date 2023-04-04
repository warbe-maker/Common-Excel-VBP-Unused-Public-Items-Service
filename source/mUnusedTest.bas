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
' The test procedure is also a usage example.
' ----------------------------------------------------------------
    Const PROC              As String = "Test_UnusedPublic"
    Const COMPS_EXCLUDED    As String = "clsQ,mRng,fMsg,mBasic,mDct,mErH,mFso,mMsg,mNme,mReg,mTrc,mWbk,mWsh"
    Const LINES_EXCLUDED    As String = "Select Case ErrMsg(ErrSrc(PROC))" & vbCrLf & _
                                        "Case vbResume:*Stop:*Resume" & vbCrLf & _
                                        "Case Else:*GoTo xt"
        
    mBasic.BoP ErrSrc(PROC)
    '~~ Providing the Workbook argments saves the Workbook selection dialog
    '~~ Providing the specification of the excluded VBComponents saves the selection dialog
    '~~ Providing excluded lines may improve the overall performance
    mUnused.Unused Application.Workbooks("CompMan.xlsb"), COMPS_EXCLUDED, LINES_EXCLUDED
    mBasic.EoP ErrSrc(PROC)
    
    mTrc.Dsply

End Sub

