Attribute VB_Name = "mUnusedPublicTest"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mFsoTest: Test of all services of the module.
'
' ----------------------------------------------------------------
Private Const COMPS_EXCLUDED = "clsQ,mRng,fMsg,mBasic,mDct,mErH,mFso,mMsg,mNme,mReg,mTrc,mWbk,mWsh"

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mUnusedPublicTest" & "." & e_proc
End Function

Private Sub Test_AnalysePublicUnused()
    Const PROC = "Test_AnalysePublicUnused"
        
    mBasic.BoP ErrSrc(PROC)
    mUnusedPublic.PublicUnused
    mBasic.EoP ErrSrc(PROC)
    
    mTrc.Dsply

End Sub

