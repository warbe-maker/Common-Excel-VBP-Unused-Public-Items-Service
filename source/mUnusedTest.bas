Attribute VB_Name = "mUnusedTest"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mUnusedTest: Test of all services of the module.
'
' ----------------------------------------------------------------
Private Const COMPS_EXCLUDED = "clsQ,mRng,fMsg,mBasic,mDct,mErH,mFso,mMsg,mNme,mReg,mTrc,mWbk,mWsh"

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mUnusedPublicTest" & "." & e_proc
End Function

Private Sub Test_UnusedPublic()
    Const PROC = "Test_UnusedPublic"
        
    mBasic.BoP ErrSrc(PROC)
    '~~ Prior specification of the Workbook saved the Workbook selection dialog
    Set wbkServiced = Application.Workbooks("CompMan.xlsb")
    
    '~~ Prior spcification of the excluded VBComponents saves the selection dialog
    mUnused.Excluded = COMPS_EXCLUDED
    
    '~~ Specifying Workbook/VB-Project specific code lines which are a kind of standard and should be excluded
    mUnused.ExcludedCodeLines = "Select Case ErrMsg(ErrSrc(PROC))" & vbCrLf & _
                                "Case vbResume:  Stop: Resume" & vbCrLf & _
                                "Case Else:      GoTo xt"
    mUnused.Unused
    mBasic.EoP ErrSrc(PROC)
    
    mTrc.Dsply

End Sub

