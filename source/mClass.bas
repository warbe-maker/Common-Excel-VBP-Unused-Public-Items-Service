Attribute VB_Name = "mClass"
Option Explicit

Private dctInstncsProcLocal     As Dictionary ' collection of Procedure local declared class instances
Private dctInstncsCompGlobal    As Dictionary ' collection of VBComponent class instances
Private dctInstncsVBPrjctGlobal As Dictionary
Private dctClassModules         As Dictionary

Public Sub CollectInstncsCompGlobal()
' ------------------------------------------------------------------------------------
' Collects for all collected VBComponents (dctComps) the global Class Instances
' in a Directory whereby the key is <comp> and the item is the Directory of
' instances with key = <instance-name> and item = <class-module-name>.
' Note: For DataModules = Worksheet the instance and the class name are the Worksheets
'       ModuleName.
' ------------------------------------------------------------------------------------
    Const PROC = "CollectInstncsCompGlobal"
    
    On Error GoTo eh
    Dim i           As Long
    Dim lStartsAt   As Long
    Dim lStopLoop   As Long
    Dim sAs         As String
    Dim sComp       As String
    Dim sItem       As String
    Dim sLine       As String
    Dim vbc         As VBComponent
    Dim vbcm        As CodeModule
    Dim vComp       As Variant
    Dim dct         As Dictionary
    Dim cllComp     As Collection
    
    BoP ErrSrc(PROC)
    Set dctInstncsCompGlobal = New Dictionary
    For Each vComp In dctComps
        sComp = vComp
        Set cllComp = dctComps(vComp)
        Set vbc = CompCollVBC(cllComp)
        Set vbcm = vbc.CodeModule
        Set dct = New Dictionary
        
        i = 1
        While i <= vbcm.CountOfDeclarationLines
            sLine = mLine.NextLine(vbcm, i, lStartsAt)
            lStopLoop = 0
            While mLine.DeclaresInstanceGlobal(sLine, sItem, sAs)
                lStopLoop = lStopLoop + 1
                If lStopLoop > 50 Then Stop
                If mClass.IsModule(sAs) Then
                    If Not dctInstncsCompGlobal.Exists(sItem) Then
                        dct.Add sItem, sAs
                    End If
                    sLine = Trim(Replace(Replace(sLine, " As " & sAs, vbNullString, 1, 1), sItem, vbNullString, 1, 1))
                End If
            Wend
            If lStartsAt = 0 Then i = i + 1
        Wend
        
        If Not dctInstncsCompGlobal.Exists(sComp) Then
            dctInstncsCompGlobal.Add sComp, dct
            Set dct = Nothing
        End If
    Next vComp

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

                                            
Public Sub CollectInstncsVBPGlobal()
    
    Dim cll             As Collection
    Dim dct             As Dictionary
    Dim wsh             As Worksheet
    Dim v               As Variant
    Dim sComp           As String
    Dim vbc             As VBComponent
    Dim vbcm            As CodeModule
    Dim i               As Long
    Dim sItem           As String
    Dim sAs             As String
    Dim sLine           As String
    Dim enKoItem        As enKindOfItem
    Dim enKoComp        As enKindOfComponent
    Dim lNextSubLine    As Long
    
    Set dctInstncsVBPrjctGlobal = New Dictionary
    
    '~~ Collect per-se VBP-global Class instances (Workbook and the Worksheets)
    '~~ in the dctInstncsVBPrjctGlobal Directory under a vbNullString key !
    Set dct = New Dictionary
    AddAscByKey dct, wbkServiced.CodeName, wbkServiced.CodeName
    For Each wsh In wbkServiced.Worksheets
        AddAscByKey dct, wsh.CodeName, wsh.CodeName
    Next wsh
    AddAscByKey dctInstncsVBPrjctGlobal, vbNullString, dct
    Set dct = Nothing
    
    '~~ Collect VBP-global Class instance in VBComponents (those declared Public)
    For Each v In dctComps
        Set dct = New Dictionary
        sComp = v
        Set cll = dctComps(v)
        Set vbc = CompCollVBC(cll)
        enKoComp = CompCollKind(cll)
        Set vbcm = vbc.CodeModule
        i = 0
        mLine.NextLine vbcm, i, lNextSubLine
        While i <= vbcm.CountOfDeclarationLines And sLine <> vbNullString
            If sLine Like "Public *" Then
                If DeclaresPublicItem(i, sLine, sItem, sAs, vbcm, enKoComp, enKoItem) Then
                    If mClass.IsModule(sAs, vbc) Then
                        AddAscByKey dct, sItem, sAs
                    Else
                        ItemCollect sComp, sItem, i, sLine, enKoComp, enKoItem
                    End If
                    
                End If
            End If
            sLine = mLine.NextLine(vbcm, i, lNextSubLine)
        Wend
        AddAscByKey dctInstncsVBPrjctGlobal, sComp, dct
        Set dct = Nothing
    Next v
            
End Sub

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mClass" & "." & e_proc
End Function

'Public Function IsInstanceLocal(ByVal c_comp As String, _
'                                ByVal c_proc As String, _
'                                ByVal c_instance As String, _
'                                ByRef c_class As String) As Boolean
'' ------------------------------------------------------------------------------
'' When the instance (c_instance) exists in the <comp>.<proc> the function
'' returns TRUE and the name of the Class-Module of the instance (c_class).
'' ------------------------------------------------------------------------------
'    Const PROC = "IsInstanceLocal"
'
'    On Error GoTo eh
'    Dim dct     As Dictionary
'    Dim sKey    As String
'
'    sKey = c_comp & "." & c_proc
'    If dctInstncsProcLocal.Exists(sKey) Then
'        Set dct = dctInstncsProcLocal(sKey)
'        If dct.Exists(c_instance) Then
'            IsInstanceLocal = True
'            c_class = dct(c_instance)
'        End If
'        Set dct = Nothing
'    End If
'
'xt: Exit Function
'
'eh: Select Case ErrMsg(ErrSrc(PROC))
'        Case vbResume:  Stop: Resume
'        Case Else:      GoTo xt
'    End Select
'End Function

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

Public Sub Init()
    Set dctInstncsProcLocal = Nothing   ' All Procedure local declared class instances
    Set dctInstncsCompGlobal = Nothing  ' All VBComponent global declared class instances
End Sub

Public Sub CollectInstncsProcLocal()
' ------------------------------------------------------------------------------------
' Collects for all collected procedures (dctProcs) the local Class Instances in
' Directory whereby the key is <comp>.<proc> and the item is the Directory of
' instances with key = <instance-name> and item = <class-module-name>.
' ------------------------------------------------------------------------------------
    Const PROC = "CollectInstncsProcLocal"
    
    On Error GoTo eh
    Dim cllProc         As Collection
    Dim i               As Long
    Dim lFrom           As Long
    Dim lStartsAt       As Long
    Dim lTo             As Long
    Dim sAs             As String
    Dim sComp           As String
    Dim sItem           As String
    Dim sLine           As String
    Dim sProc           As String
    Dim vbcm            As CodeModule
    Dim vProc           As Variant
    Dim vComp           As Variant
    Dim dctCompProcs    As Dictionary
    Dim dctInstComp     As Dictionary
    Dim dctInstProc     As Dictionary
    
    BoP ErrSrc(PROC)
    If dctInstncsProcLocal Is Nothing _
    Then Set dctInstncsProcLocal = New Dictionary
    
    For Each vComp In dctProcs
        Set dctInstComp = New Dictionary
        sComp = vComp
        Set dctCompProcs = dctProcs(sComp)
        For Each vProc In dctCompProcs
            sProc = vProc
            Set cllProc = dctCompProcs(vProc)
            Set vbcm = mProc.CollVBCM(cllProc)
            lFrom = mProc.CollLineFrom(cllProc)
            lTo = mProc.CollLineTo(cllProc)
            Set cllProc = Nothing
            Set dctInstProc = New Dictionary
            i = lFrom - 1
'            If sComp = "mExport" And sProc = "All" Then Stop
            sLine = mLine.NextLine(vbcm, i, lStartsAt)
            While i <= lTo And sLine <> vbNullString
'                If InStr(sLine, "ByRef a_stats As clsStats = Nothing") <> 0 Then Stop
                Select Case True
                    Case sLine Like "* As New *"
                        sItem = Split(Trim(Split(sLine, " As New ")(0)), " ")(UBound(Split(Trim(Split(sLine, " As ")(0)), " ")))
                        sAs = Split(Trim(Split(sLine, " As New ")(1)), " ")(0)
                        If mClass.IsModule(sAs) Then
                            dctInstProc.Add sItem, sAs
                        End If
                    Case sLine Like "* As *"
                        sItem = Split(Trim(Split(sLine, " As ")(0)), " ")(UBound(Split(Trim(Split(sLine, " As ")(0)), " ")))
                        sAs = Split(Trim(Split(sLine, " As ")(1)), " ")(0)
                        If mClass.IsModule(sAs) Then
                            If Not dctInstProc.Exists(sItem) Then
                                dctInstProc.Add sItem, sAs
                            End If
'                            Debug.Print "Local instance '" & sItem & "' of Class '" & sAs & "' (" & sLine & ")"
                        End If
                End Select
                sLine = mLine.NextLine(vbcm, i, lStartsAt)
            Wend
        
            If Not dctInstComp.Exists(sProc) And dctInstProc.Count <> 0 Then
                AddAscByKey dctInstComp, sProc, dctInstProc
                Set dctInstProc = Nothing
            End If
        Next vProc
    
        If Not dctInstncsProcLocal.Exists(sComp) And dctInstComp.Count <> 0 Then
            AddAscByKey dctInstncsProcLocal, sComp, dctInstComp
            Set dctInstComp = Nothing
        End If
    Next vComp
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function IsInstance(ByVal i_comp_name As String, _
                           ByVal i_instance_name As String, _
                           ByRef i_class_name As String, _
                  Optional ByVal i_proc_name As String = vbNullString) As Boolean
' ------------------------------------------------------------------------------
' When the instance (i_instance_name) is a known Class instance the function
' returns TRUE and the corresponding class' name (i_class_name)
' ------------------------------------------------------------------------------
    Const PROC = "IsInstance"
    
    On Error GoTo eh
    Dim dct As Dictionary
    Dim v   As Variant
    
    If i_proc_name <> vbNullString Then
        If dctInstncsProcLocal.Exists(i_comp_name) Then
            Set dct = dctInstncsProcLocal(i_comp_name)
            If dct.Exists(i_proc_name) Then
                Set dct = dct(i_proc_name)
                If dct.Exists(i_instance_name) Then
                    i_class_name = dct(i_instance_name)
                    IsInstance = True
                    GoTo xt
                End If
            End If
        End If
    End If
    
    '~~ When the i_instance_name is not known as a local class instance
    '~~ it still may be known as a VBComponent global one
    If dctInstncsCompGlobal.Exists(i_comp_name) Then
        Set dct = dctInstncsCompGlobal(i_comp_name)
        If dct.Exists(i_instance_name) Then
            i_class_name = dct(i_instance_name)
            IsInstance = True
            GoTo xt
        End If
    End If
    
    '~~ When the i_instance is not known as a VBComponent global class instance
    '~~ it may still be a VBProject global class instance declared in a VBComponent
    If dctInstncsVBPrjctGlobal.Exists(i_comp_name) Then
        Set dct = dctInstncsVBPrjctGlobal(i_comp_name)
        If dct.Exists(i_instance_name) Then
            i_class_name = dct(i_instance_name)
            IsInstance = True
            GoTo xt
        End If
    End If
            
    For Each v In dctInstncsVBPrjctGlobal
        Set dct = dctInstncsVBPrjctGlobal(v)
        If dct.Exists(i_instance_name) Then
            i_class_name = dct(i_instance_name)
            IsInstance = True
            GoTo xt
        End If
    Next v
    
    '~~ When the i_instance is not known as a VBComponent global class instance declared in a VBComponent
    '~~ it may still be a class instance like the Workbook itself or any of its Worksheets
    If dctInstncsVBPrjctGlobal.Exists(vbNullString) Then
        Set dct = dctInstncsVBPrjctGlobal(vbNullString)
        If dct.Exists(i_instance_name) Then
            i_class_name = dct(i_instance_name)
            IsInstance = True
            GoTo xt
        End If
    End If
            
            
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Property Get IsModule(Optional ByVal i_name As String, _
                             Optional ByVal i_vbc As VBComponent) As Boolean
    Set i_vbc = i_vbc
    If Not dctClassModules Is Nothing Then
        IsModule = dctClassModules.Exists(i_name)
    End If
End Property

Public Property Let IsModule(Optional ByVal i_name As String, _
                             Optional ByVal i_vbc As VBComponent, _
                                      ByVal i_is As Boolean)
                                            
    If dctClassModules Is Nothing Then Set dctClassModules = New Dictionary
    If i_is Then
        If Not dctClassModules.Exists(i_name) Then
            AddAscByKey dctClassModules, i_name, i_vbc
        End If
    End If
    
End Property



