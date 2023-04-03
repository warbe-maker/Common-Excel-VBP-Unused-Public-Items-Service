Attribute VB_Name = "mLine"
Option Explicit

Public Function LineExcluded(ByVal l_line As String) As Boolean

    Static vExcludedCodeLines   As Variant
    Static bInitialized         As Boolean
    Dim v                       As Variant
    
    If mUnused.ExcludedCodeLines = vbNullString Then Exit Function
    If Not bInitialized Then
        vExcludedCodeLines = Split(mUnused.ExcludedCodeLines, vbCrLf)
        bInitialized = True
    End If
    
    For Each v In vExcludedCodeLines
        If l_line Like "*" & v & "*" Then
            LineExcluded = True
            Exit Function
        End If
    Next v
    
End Function

Private Sub DropLineNumber(ByRef c_line As String)
' ------------------------------------------------------------------------------------
' Returns a code line (c_line) with a line-number removed.
' ------------------------------------------------------------------------------------
    Dim v   As Variant
    
    If Len(c_line) > 1 Then
        v = Split(c_line, " ")
        If IsNumeric(Split(c_line, " ")(0)) Then
            c_line = Trim(Replace(c_line, v(0), vbNullString))
        End If
    End If

End Sub

Public Function ForBeingParsed(ByVal c_line_to_parse As String, _
                               ByRef c_line_for_parsing As String, _
                               ByVal c_comp_to_parse As String, _
                               ByVal c_proc_to_parse As String, _
                               ByRef c_delim As String, _
                               ByVal c_ignore As String) As String
' ------------------------------------------------------------------------------------
' Returns a code line (c_line_to_parse) prepared for being explored (c_line_for_parsing) which means:
' - all potential strings are enclosed/demimited by a spspecified and returned
'   delimiter (c_delim)
' - A " ." is replaced by the top WithStack item - which may be the corresponding
'   class modules's name when the preceeding With line identified a known class
'   instance
' - Any .xxx is replaced by the corresponding class modules name when xxx is a known
'   instance of a class module
' ------------------------------------------------------------------------------------
    Const PROC = "ForBeingParsed"
    
    On Error GoTo eh
    Dim sClass      As String
    Dim i           As Long
    Dim vElements   As Variant
    Dim sItem       As String
    Dim vElement    As Variant
    Dim vItems      As Variant
    Dim sComp       As String
    
    c_delim = " | "
    
    '~~ Drop what is to be ignored
    c_line_for_parsing = Replace(c_line_to_parse, " " & c_ignore, " ")
    
    '~~ Replaces any structuring string into a c_delim string and enclos the line in spaces
    mLine.ReplaceByDelimiter c_line_for_parsing, " | ", "(", ")", ", ", ":=", " = ", "!", """"
    c_line_for_parsing = " " & Trim(c_line_for_parsing) & " "
    
    '~~ Replace a "Me." by the component's name
    c_line_for_parsing = Replace(c_line_for_parsing, " Me.", " " & c_comp_to_parse & ".")
    
    '~~ Replace a default object (" .") with the corresponding With item
    If InStr(c_line_for_parsing, " .") <> 0 Then
        c_line_for_parsing = Replace(c_line_for_parsing, " .", " " & mWithStack.Top() & ".")
    End If
    
    '~~ Replace any Public string Constant by its value
'    For Each v In dctPublicItems
'        Set cll = dctPublicItems(v)
'        If PublicItemCollKindOfItem(cll) = enConstant Then
'            sItem = Split(v, ".")(1)
'            If InStr(c_line_for_parsing, sItem) <> 0 Then
'                Stop
'            End If
'        End If
'    Next v
    
    '~~       c_line_to_parse
    '~~ sElements
    '~~ sItems
    If InStr(c_line_for_parsing, ".") <> 0 Then
        vElements = Split(c_line_for_parsing, " ")
        For Each vElement In vElements
            If InStr(vElement, ".") <> 0 Then
                vItems = Split(Trim(vElement), ".")
                For i = LBound(vItems) To UBound(vItems)
                    sItem = vItems(i)
'                    If sItem = "Raw" Then Stop
                    If Not i = LBound(vItems) And Not i = UBound(vItems) _
                    Then sComp = vItems(i - 1) _
                    Else sComp = c_comp_to_parse
                    If mClass.IsInstance(sComp, sItem, sClass, c_proc_to_parse) Then
                        Select Case i
                            Case LBound(vItems)
                                c_line_for_parsing = Replace(c_line_for_parsing, sItem & ".", sClass & ".")
                            Case UBound(vItems)
                                c_line_for_parsing = Replace(c_line_for_parsing, "." & sItem, "." & sClass)
                            Case Else
                                c_line_for_parsing = Replace(c_line_for_parsing, "." & sItem & ".", "." & sClass & ".")
                        End Select
                    End If
                Next i
            End If
        Next vElement
    End If

xt: ForBeingParsed = c_line_for_parsing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

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

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mLine" & "." & e_proc
End Function

Public Function Ignore(ByVal c_line As String) As String
' ------------------------------------------------------------------------------------
' Returns the string of the code line's item. Examples:
' ... Const XXX = "abc" returns XXX
' ... Property Get XXX ... returns XXX
' ... Function XXX() As .. returns XXX
' ------------------------------------------------------------------------------------
    Const PROC = "Item"
    Dim sItem   As String
    
    If UBound(Split(" " & c_line & " ", " Property Get ")) > 0 Then
        sItem = Split(" " & c_line & " ", " Property Get ")(1)
        sItem = Split(sItem, "(")(0)
    ElseIf UBound(Split(" " & c_line & " ", " Function ")) > 0 Then
        sItem = Split(" " & c_line & " ", " Function ")(1)
        sItem = Split(sItem, "(")(0)
    End If
    Ignore = sItem & " = "
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsEndProc(ByVal c_line As String) As Boolean
    IsEndProc = c_line Like "*End Property*" _
                     Or c_line Like "*End Function*" _
                     Or c_line Like "*End Sub*"

End Function

Public Function IsFirstOfProc(ByVal c_proc_of_line As String, _
                              ByRef c_proc As String, _
                              ByVal c_kind_of_proc As vbext_ProcKind) As Boolean
    Static sProc As String
    
    If c_proc_of_line & "." & KoPstring(c_kind_of_proc) <> sProc Then
        IsFirstOfProc = True
        sProc = c_proc_of_line & "." & KoPstring(c_kind_of_proc)
        c_proc = sProc
    End If
    
End Function

Public Function NextLine(ByVal n_vbcm As CodeModule, _
                         ByRef n_as_of_line_no As Long, _
                Optional ByRef n_line_starts_at As Long) As String
' ------------------------------------------------------------------------------------
' Returns the next code line (sLine) as of a start line (n_as_of_line_no) by:
' - skipping empty lines
' - kipping comment lines
' - unstripping comments
' - replacing constants with ""
' - concatenating continuation lines
' - splitting lines with ': ' into separate lines.
' When the line has no sub-lines (delimited by ": ") the returned next sub-line is 0
' else the next one due.
' ------------------------------------------------------------------------------------
    Const PROC = "NextLine"
    
    On Error GoTo eh
    Static lStartsAt    As Long
    Static lNextSubLine As Long ' when > 1 a sub-line (: ) of the sLine is to be returned
    Static sLine        As String
    Dim v               As Variant
    Dim lStopLoop       As Long
    Dim vItems          As Variant
    Dim i               As Long
    Dim lLineNo         As Long
    
    If lNextSubLine <> 0 Then
        '~~ When lNextSubLine is not 0 there is another one still to be returned from the sLine.
        '~~ When the returned sub-line is the last one of the multiple lines in sLine the
        '~~ returned lNextSubLine is 0
        NextLine = NextSubLine(sLine, lNextSubLine)
        GoTo xt
    End If
    
    n_as_of_line_no = n_as_of_line_no + 1
    
    sLine = Trim(n_vbcm.Lines(n_as_of_line_no, 1))
    DropLineNumber sLine
    SkipLinesEmptyOrComment n_as_of_line_no, sLine, n_vbcm
    If Len(sLine) = 0 Then Exit Function ' no line with a content found
    
    n_line_starts_at = n_as_of_line_no
    ConcatenateAllContinuationLines n_as_of_line_no, sLine, n_vbcm
    UnstripCommentIfAny sLine
    
    '~~ When the line does not have multiple code lines the reuturned line is the line as it is
    '~~ Else the returned line is the first of the multiple lines in the code line and the
    '~~ lNextSubLine is incremented by one for the next one being returned with the subsequent
    '~~ NextLine call
    NextLine = NextSubLine(sLine, lNextSubLine)
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function NextSubLine(ByRef c_line As String, _
                             ByRef c_next_sub_line As Long) As String
    Dim v   As Variant
    
    v = Split(c_line, ": ")
    NextSubLine = Trim(v(c_next_sub_line))
    
    If UBound(v) > c_next_sub_line Then
        c_next_sub_line = c_next_sub_line + 1
    Else
        c_next_sub_line = 0
    End If

End Function

Private Sub UnstripCommentIfAny(ByRef c_line As String)
' ------------------------------------------------------------------------------------
' Returns a code line (c_line) with a possible comment unstripped by ignoring any
' comment indicated by a space followed by a '.
' ------------------------------------------------------------------------------------
    Dim i As Long
    Dim v As Variant
    
    v = Split(c_line, """")
    For i = LBound(v) To UBound(v)
        Select Case i
            Case 0, 2, 4, 6, 8, 10, 12, 14
                If InStr(v(i), " '") <> 0 Then
                    c_line = Trim(Split(c_line, " '")(0))
                    Exit For
                End If
        End Select
    Next i
End Sub

Private Sub SkipLinesEmptyOrComment(ByRef c_i As Long, _
                                            ByRef c_line As String, _
                                            ByVal c_vbcm As CodeModule)
' ------------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------------

    c_line = Trim(c_line)
    With c_vbcm
        While (Len(c_line) = 0 Or c_line Like "'*") And c_i <= .CountOfLines
            c_i = c_i + 1
            c_line = Trim(.Lines(c_i, 1))
        Wend
    End With
    
End Sub


Private Sub ConcatenateAllContinuationLines(ByVal c_line_starts_at As Long, _
                                            ByRef c_line As String, _
                                            ByVal c_vbcm As CodeModule)
' ------------------------------------------------------------------------------------
' Returns a code line (c_line) with all continuation lines connected by preserving the
' start of the code line (c_line_starts_at).
' ------------------------------------------------------------------------------------
    Dim iLine   As Long
    
    iLine = c_line_starts_at
    With c_vbcm
        While Right(Trim(c_line), 1) = "_" And iLine <= .CountOfLines
            iLine = iLine + 1
            c_line = Left(c_line, Len(c_line) - 1) & Trim(.Lines(iLine, 1))
            c_line = Replace(c_line, " ,", ",")
        Wend
    End With
    
End Sub

Private Sub RemoveConstantStrings(ByRef c_line As String)
' ------------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------------
    Dim v       As Variant
    Dim lStop   As Long
    
    If Not c_line Like "*Application.Run*" Then
        v = Split(Trim(c_line), """")
        lStop = 0
        While UBound(v) >= 1
            lStop = lStop + 1
            If lStop > 50 Then Stop
            c_line = Replace(c_line, """" & v(1) & """", vbNullString) & """"
            c_line = Replace(c_line, """", vbNullString)
            v = Split(c_line, """")
        Wend
    End If

End Sub

Public Function RefersToPublicItem(ByVal c_line_to_explore As String, _
                                   ByVal c_public_item As String, _
                                   ByVal c_using_vbc As VBComponent, _
                                   ByVal c_dot_class As String) As Boolean
' ------------------------------------------------------------------------------------
' Returns TRUE when the code line (c_line_to_explore) contains the 'Public Item' (c_public_item).
' ------------------------------------------------------------------------------------
    Const PROC = "RefersToPublicItem"
    
    On Error GoTo eh
    Dim sItem           As String
    Dim sPublicItemComp As String
    Dim sPublicItemItem As String
    Dim i               As Long
    Dim lStopLoop       As Long
    Dim sPublicItem     As String
    
    sPublicItemComp = Split(c_public_item, ".")(0)
    sPublicItemItem = Split(c_public_item, ".")(1)
    sPublicItem = sPublicItemComp & "." & sPublicItemItem
    
    If InStr(c_line_to_explore, " " & sPublicItem & " ") _
    Or InStr(c_line_to_explore, "." & sPublicItem & " ") <> 0 Then
        '~~ Fully qualified public item
        RefersToPublicItem = True
        GoTo xt
    End If
             
    If InStr(c_line_to_explore, "." & sPublicItemItem & " ") <> 0 Then
        If mPublic.IsUniqueItem(sPublicItemItem) Then
            '~~ When unique this is the public item checked
            RefersToPublicItem = True
            GoTo xt
        End If
    End If
    
    If InStr(c_line_to_explore, " " & sPublicItemItem & " ") <> 0 Then
        If mPublic.IsUniqueItem(sPublicItemItem) Then
            '~~ when unique even unqualified is a clear usage indication
            RefersToPublicItem = True
            GoTo xt
        End If
    End If
        
    '~~ 3. Check if the code line contains the <proc> unqualified
'    If c_line_to_explore Like "* " & sPublicItemItem & " *" _
'    Or c_line_to_explore Like "*." & sPublicItemItem & " *" Then Stop
    
    Select Case True
        Case c_line_to_explore Like "* " & sPublicItemItem & " *"
            '~~ This is an unqualified call of a Public item
            If mPublic.IsUniqueItem(sPublicItemItem) Then
                '~~ Since the name is unique the unqualified call concerns to the Public item
                RefersToPublicItem = True
                GoTo xt
            ElseIf mProc.IsLocal(c_using_vbc.name, sPublicItemItem) Then
                '~~ The unqualified call concerns the local procedure
                GoTo xt
            End If
        Case c_line_to_explore Like "*." & sPublicItemItem & " ", _
             c_line_to_explore Like "*." & sPublicItemItem & "."
            '~~ This is a qualified call of a public item - which may still not be the one meant!
            sItem = Split(c_line_to_explore, ".")(1)
            sItem = c_dot_class & "." & sItem
            If sItem = c_public_item Then
                '~~ The qualified call concerns definitely the Public item explored
                RefersToPublicItem = True
                GoTo xt
            End If
        Case c_line_to_explore Like "*.*"
            i = 0
            
            lStopLoop = 0
'            c_line_to_explore = Replace(c_line_to_explore, " Me.", " " & c_using_vbc.name & ".")
            While InStr(c_line_to_explore, ".") <> 0 And i < 10 ' 10 iterations = 10 dots in the line !!
                lStopLoop = lStopLoop + 1
                If lStopLoop > 50 Then Stop
                i = i + 1
                On Error Resume Next
                
                sItem = Split(Split(c_line_to_explore, ".")(0), " ")(UBound(Split(Split(c_line_to_explore, ".")(0), " "))) & "." & Split(Split(c_line_to_explore, ".")(1), " ")(0)
                If Err.Number <> 0 Then
                    Debug.Print "Extracting an item failed with line '" & c_line_to_explore & "'"
                    Stop
                End If
                On Error GoTo eh
                If sItem = c_public_item Then
                    RefersToPublicItem = True
                    GoTo xt
                End If
                c_line_to_explore = Replace(c_line_to_explore, sItem, vbNullString) '
            Wend
            If i = 10 Then
                Debug.Print "Extracting an item '" & sItem & "' failed with line '" & c_line_to_explore & "'"
            End If
'            If sItem = "." Then Stop
            GoTo xt
    End Select
        
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function DeclaresInstanceGlobal(ByRef c_line As String, _
                                                ByRef c_item As String, _
                                                ByRef c_as As String) As Boolean
' ------------------------------------------------------------------------------------
' When the code line (c_line) declares an item where "As" names a known Class Module
' the function returns:
' - TRUE
' - The declared item (c_item)
' - The As string (c_as)
' ------------------------------------------------------------------------------------
    Const PROC = "DeclaresInstanceGlobal"
    
    On Error GoTo eh
    
    If c_line Like "* As New *" Then
'        Debug.Print c_vbcm.Parent.name & ", Line number=" & c_line_no & ", Line= >" & c_line & "<"
        mPublic.DeclaredAs c_line, c_as
        DeclaresInstanceGlobal = mClass.IsModule(c_as)
        If Not DeclaresInstanceGlobal Then GoTo xt
        c_item = Trim(Split(c_line, " As New ")(0))
        c_item = Trim(Split(c_item, " ")(UBound(Split(c_item, " "))))
    ElseIf c_line Like "* As *" Then
'        Debug.Print c_vbcm.Parent.name & ", Line number=" & c_line_no & ", Line= >" & c_line & "<"
        mPublic.DeclaredAs c_line, c_as
        DeclaresInstanceGlobal = mClass.IsModule(c_as)
        If Not DeclaresInstanceGlobal Then GoTo xt
        c_item = Trim(Split(c_line, " As ")(0))
        c_item = Trim(Split(c_item, " ")(UBound(Split(c_item, " "))))
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function DeclaresPublicItem(ByRef c_line_no As Long, _
                                   ByRef c_line As String, _
                                   ByRef c_item As String, _
                                   ByRef c_as As String, _
                                   ByVal c_vbcm As CodeModule, _
                                   ByVal c_kind_of_comp As enKindOfComponent, _
                                   ByRef c_kind_of_item As enKindOfItem) As Boolean
' ------------------------------------------------------------------------------------
' When the code line (c_line) declares a "Public " item, i.e. the line starts with
' "Public ", the function returns:
' - TRUE
' - The declared item (c_item)
' - The As string (c_as)
' ------------------------------------------------------------------------------------
    Const PROC = "DeclaresPublicItem"
    
    On Error GoTo eh
    Dim sComp           As String
    Dim sItem           As String
    Dim lLoopStop       As Long
    Dim lStartsAt       As Long
    Dim sLine           As String
    
    If Not c_line Like "Public *" Then GoTo xt
    DeclaresPublicItem = True
'    Debug.Print c_vbcm.Parent.name & ", Line number=" & c_line_no & ", Line= >" & c_line & "<"
    
    Select Case True
        Case c_line Like "Public Type *"
            sComp = c_vbcm.Parent.name
            lLoopStop = 0
            sLine = mLine.NextLine(c_vbcm, c_line_no, lStartsAt)
            While Not sLine = "End Type" And Not sLine <> vbNullString
                lLoopStop = lLoopStop + 1
                If lLoopStop > 100 Then Stop
                c_line = c_line & " "
                sItem = Split(c_line, " ")(0)
                ItemCollect sComp, sItem, c_line_no, c_line, c_kind_of_comp, enUserDefinedType
                sLine = mLine.NextLine(c_vbcm, c_line_no, lStartsAt)
            Wend
            DeclaresPublicItem = False
            GoTo xt

        Case c_line Like "Public Enum *"
            sComp = c_vbcm.Parent.name
            lLoopStop = 0
            sLine = mLine.NextLine(c_vbcm, c_line_no, lStartsAt)
            While Not sLine Like "End Enum" And Not sLine <> vbNullString
                lLoopStop = lLoopStop + 1
                If lLoopStop > 500 Then Stop
                c_line = c_line & " "
                c_line = Split(c_line, " ")(0)
                ItemCollect sComp, sItem, c_line_no, c_line, c_kind_of_comp, enEnumeration
                sLine = mLine.NextLine(c_vbcm, c_line_no, lStartsAt)
            Wend
            DeclaresPublicItem = False
            GoTo xt
            
        Case c_line Like "Public Const *":          Item "Public Const ", c_line, c_item:                 c_kind_of_item = enConstant
        Case c_line Like "Public Sub *"
                                                    If sComp = "mExport" Then Stop
                                                    mPublic.Item "Public Sub ", c_line, c_item:                   c_kind_of_item = enSub
        Case c_line Like "Public Function *":       mPublic.ItemAs "Public Function ", c_line, c_item, c_as:      c_kind_of_item = enFunction
        Case c_line Like "Public Property Get *":   mPublic.ItemAs "Public Property Get ", c_line, c_item, c_as:  c_kind_of_item = enPropertyGet
        Case c_line Like "Public Property Let *":   mPublic.Item "Public Property Let ", c_line, c_item:          c_kind_of_item = enPropertyGet
        Case c_line Like "Public Property Set *":   mPublic.Item "Public Property Set ", c_line, c_item:          c_kind_of_item = enPropertySet
        Case c_line Like "Friend Property Get *":   mPublic.ItemAs "Friend Property Get ", c_line, c_item, c_as:  c_kind_of_item = enPropertyGet
        Case c_line Like "Friend Property Let *":   mPublic.Item "Friend Property Let ", c_line, c_item:          c_kind_of_item = enPropertyLet
        Case c_line Like "Friend Property Set *":   mPublic.Item "Friend Property Set ", c_line, c_item:          c_kind_of_item = enPropertySet
        Case c_line Like "Public *":                mPublic.Variable "Public ", c_line, c_item, c_as:             c_kind_of_item = enVariable
    End Select
            
    If c_item = vbNullString Then
        Debug.Print c_vbcm.Parent.name & " Line " & c_line_no & " ???? no item"
        Stop
    Else
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub ReplaceByDelimiter(ByRef r_line As String, _
                              ByVal r_delim As String, _
                              ParamArray r_elements() As Variant)
' ------------------------------------------------------------------------------------
' Replaces all elements (r_elements) in a string (r-line) by a delimiter (r_delim).
' ------------------------------------------------------------------------------------
    Dim v As Variant
    For Each v In r_elements
        r_line = Replace(r_line, v, r_delim)
    Next v
    
End Sub

Public Sub StopIfLike(ByVal c_line As String, _
                               ByVal c_like As String, _
                      Optional ByVal c_comp_using As String = vbNullString, _
                      Optional ByVal c_comp As String)
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    If c_like <> vbNullString Then
        If c_comp_using <> vbNullString Then
            If c_line Like "*" & c_like & "*" _
            And c_comp_using = c_comp Then
                Stop
            End If
        ElseIf c_line Like "*" & c_like & "*" Then
            Stop
        End If
    End If
End Sub


