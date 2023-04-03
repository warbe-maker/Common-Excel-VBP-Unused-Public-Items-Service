Attribute VB_Name = "mProc"
Option Explicit

Public LinesTotal   As Long
Public ProcsTotal   As Long

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

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
    
    BoP ErrSrc(PROC)
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
                        AddAscByKey dctCompProcs, sKey, cllCompProc
                        Set cllCompProc = Nothing
                    End If
                End If
                sLine = mLine.NextLine(vbcm, i, lStartsAt)
            Wend
        End With
        
        If Not dctProcs.Exists(sComp) Then
            AddAscByKey dctProcs, sComp, dctCompProcs
            Set dctCompProcs = Nothing
        End If
        
        '~~ By the way collect all components which do have an equally named procedure
        If Not dctPublicItemsUnique.Exists(sProc) Then
            Set cllCompProc = New Collection
            cllCompProc.Add sComp
            AddAscByKey dctPublicItemsUnique, sProc, cllCompProc
            Set cllCompProc = Nothing
        ElseIf dctPublicItemsUnique.Exists(sProc) Then
            Set cllCompProc = dctPublicItemsUnique(sProc)
            cllCompProc.Add sComp
            dctPublicItemsUnique.Remove sProc
            AddAscByKey dctPublicItemsUnique, sProc, cllCompProc
            Set cllCompProc = Nothing
        End If
    
    Next v
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) the
'          function is used to turn the positive number into a negative one.
'          The error message will regard a negative error number as an
'          'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
' - ErrSrc The caller provides the (name of the) source of the error through
'          the module specific function ErrSrc(PROC) which adds the module
'          name to the procedure name.
'
' W. Rauschenberger Berlin, May 2022
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

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

