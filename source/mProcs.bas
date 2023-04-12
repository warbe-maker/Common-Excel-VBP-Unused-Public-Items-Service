Attribute VB_Name = "mProcs"
Option Explicit

Public LinesTotal   As Long
Public ProcsTotal   As Long

Public Sub Collect()
' ------------------------------------------------------------------------------
' Assembles for each Procedure in non excluded VBComponents (dctComps)
' the first relevant code line (Sub, Function, Property), its line number and
' the Procedures last line number (End xxx) in a Collection and returns a
' Dictionary with these Collections as the item and <compname>.<procname> as the
' key.
' ------------------------------------------------------------------------------
    Const PROC = "Collect"
    
    On Error GoTo eh
    Dim dctCompProcs    As Dictionary
    Dim cllCompProc     As Collection
    Dim cllComp         As Collection
    Dim i               As Long
    Dim KoP             As vbext_ProcKind
    Dim sKey            As String
    Dim sLine           As String
    Dim v               As Variant
    Dim vbc             As VBComponent
    Dim vbcm            As CodeModule
    Dim sProc           As String
    Dim sComp           As String
    Dim lStop           As Long
    Dim lStartsAt       As Long
    Dim lFrom           As Long
    Dim lLines          As Long
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Collect all First-Of-Proc lines
    LinesTotal = 0
    ProcsTotal = 0
    Set dctProcs = New Dictionary
        
    '~~ Collect all Procedures
    Set dctProcs = New Dictionary
    For Each v In dctComps
        Set dctCompProcs = New Dictionary
        sComp = v
        Set cllComp = dctComps(v)
        Set vbc = CompCollVBC(cllComp)
        Set vbcm = vbc.CodeModule
        With vbcm
            i = .CountOfDeclarationLines
            lStop = 0
            '~~ Collect all Procedures in the CodeModule
            sLine = mLine.NextLine(vbcm, i, lStartsAt)
            While sLine <> vbNullString And i <= .CountOfLines
                If mLine.IsFirstOfProc(.ProcOfLine(i, KoP), sProc, KoP) Then
                    lFrom = i
                    sKey = Split(sProc, ".")(0)
                    Set cllCompProc = New Collection
                    cllCompProc.Add vbcm
                    cllCompProc.Add sLine
                    cllCompProc.Add lFrom
                End If
                If mLine.IsEndProc(sLine) Then
                    ProcsTotal = ProcsTotal + 1
                    cllCompProc.Add i
                    lLines = (cllCompProc(cllCompProc.Count) - cllCompProc(cllCompProc.Count - 1)) + 1
                    LinesTotal = LinesTotal + lLines
                    If Not dctCompProcs.Exists(sKey) Then
                        dctCompProcs.Add sKey, cllCompProc
                        Set cllCompProc = Nothing
                    End If
                End If
                sLine = mLine.NextLine(vbcm, i, lStartsAt)
            Wend
        End With
        
        If Not dctProcs.Exists(sComp) Then
            dctProcs.Add sComp, dctCompProcs
            Set dctCompProcs = Nothing
        End If
        
        '~~ By the way collect all components which do have an equally named procedure
        If Not dctPublicItemsUnique.Exists(sProc) Then
            Set cllCompProc = New Collection
            cllCompProc.Add sComp
            dctPublicItemsUnique.Add sProc, cllCompProc
            Set cllCompProc = Nothing
        ElseIf dctPublicItemsUnique.Exists(sProc) Then
            Set cllCompProc = dctPublicItemsUnique(sProc)
            cllCompProc.Add sComp
            dctPublicItemsUnique.Remove sProc
            dctPublicItemsUnique.Add sProc, cllCompProc
            Set cllCompProc = Nothing
        End If
    
    Next v
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function IsLocal(ByVal i_comp As String, _
                        ByVal i_proc As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the procedure (i_proc) is a component's (i_comp) procedure.
' ------------------------------------------------------------------------------
    IsLocal = dctProcs(i_comp).Exists(i_proc)
End Function



Public Function CollLine(ByVal cll As Collection) As String:       CollLine = cll(2):      End Function

Public Function CollLineFrom(ByVal cll As Collection) As String:   CollLineFrom = cll(3):  End Function

Public Function CollLineTo(ByVal cll As Collection) As String:     CollLineTo = cll(4):    End Function

Public Function CollVBCM(ByVal cll As Collection) As CodeModule:   Set CollVBCM = cll(1):  End Function

Private Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mProc" & "." & e_proc
End Function

