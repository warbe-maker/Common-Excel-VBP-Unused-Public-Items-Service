Attribute VB_Name = "mLine"
Option Explicit
' ------------------------------------------------------------------------------------
' Standard Module mLine:
' ======================
' Public services:
' - DeclaresGlobalClassInstance
' - ForBeingParsed
' - Ignore
' - IsEndProc
' - IsFirstOfProc
' - LineExcluded
' - NextLine
' - RefersToPublicItem          Returns TRUE when a code line contains the 'Public'
'                               item.
' - IsLike                  For debuging purpose only
' ------------------------------------------------------------------------------------

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

Public Function DeclaresGlobalClassInstance(ByRef c_line As String, _
                                            ByRef c_item As String, _
                                            ByRef c_as As String) As Boolean
' ------------------------------------------------------------------------------------
' When the code line (c_line) declares an item where "As" is a known Class Module
' the function returns:
' - TRUE
' - The declared item (c_item)
' - The As string (c_as)
' ------------------------------------------------------------------------------------
    Const PROC = "DeclaresGlobalClassInstance"
    
    On Error GoTo eh
    
    If c_line Like "* As New *" Then
        mItems.DeclaredAs c_line, c_as
        DeclaresGlobalClassInstance = mClass.IsClassModule(c_as)
        If Not DeclaresGlobalClassInstance Then GoTo xt
        c_item = Trim(Split(c_line, " As New ")(0))
        c_item = Trim(Split(c_item, " ")(UBound(Split(c_item, " "))))
    ElseIf c_line Like "* As *" Then
        mItems.DeclaredAs c_line, c_as
        DeclaresGlobalClassInstance = mClass.IsClassModule(c_as)
        If Not DeclaresGlobalClassInstance Then GoTo xt
        c_item = Trim(Split(c_line, " As ")(0))
        c_item = Trim(Split(c_item, " ")(UBound(Split(c_item, " "))))
    End If
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
' When the code line (c_line) declares a "Public " item the function returns:
' - TRUE
' - The declared item (c_item)
' - The As string (c_as)
' Note: For a "Public Type" or "Public Enum" all elements are regarded Public items.
' ------------------------------------------------------------------------------------
    Const PROC = "DeclaresPublicItem"
    
    On Error GoTo eh
    Dim sComp   As String
    Dim sItem   As String
    Dim sLine   As String
    
    If Not c_line Like "Public *" Then GoTo xt
    DeclaresPublicItem = True
    
    Select Case True
        Case c_line Like "Public Type *"
            sComp = c_vbcm.Parent.name
            sLine = mLine.NextLine(c_vbcm, c_line_no)
            While Not sLine = "End Type" And Not sLine <> vbNullString And c_line_no <= c_vbcm.CountOfLines
                '~~ Register all elements as Public items
                c_line = c_line & " "
                sItem = Split(c_line, " ")(0)
                CollectPublicItem sComp, sItem, c_line_no, c_line, c_kind_of_comp, enUserDefinedType
                sLine = mLine.NextLine(c_vbcm, c_line_no)
            Wend
            DeclaresPublicItem = False
            GoTo xt

        Case c_line Like "Public Enum *"
            sComp = c_vbcm.Parent.name
            sLine = mLine.NextLine(c_vbcm, c_line_no)
            While Not sLine Like "End Enum" And Not sLine <> vbNullString And c_line_no <= c_vbcm.CountOfLines
                '~~ Register all elements as Public items
                c_line = c_line & " "
                c_line = Split(c_line, " ")(0)
                CollectPublicItem sComp, sItem, c_line_no, c_line, c_kind_of_comp, enEnumeration
                sLine = mLine.NextLine(c_vbcm, c_line_no)
            Wend
            DeclaresPublicItem = False
            GoTo xt
            
        Case c_line Like "Public Const *":          mItems.Item "Public Const ", c_line, c_item:                   c_kind_of_item = enConstant
        Case c_line Like "Public Sub *":            mItems.Item "Public Sub ", c_line, c_item:                     c_kind_of_item = enSub
        Case c_line Like "Public Function *":       mItems.ItemAs "Public Function ", c_line, c_item, c_as:        c_kind_of_item = enFunction
        Case c_line Like "Public Property Get *":   mItems.ItemAs "Public Property Get ", c_line, c_item, c_as:    c_kind_of_item = enPropertyGet
        Case c_line Like "Public Property Let *":   mItems.Item "Public Property Let ", c_line, c_item:            c_kind_of_item = enPropertyGet
        Case c_line Like "Public Property Set *":   mItems.Item "Public Property Set ", c_line, c_item:            c_kind_of_item = enPropertySet
        Case c_line Like "Friend Property Get *":   mItems.ItemAs "Friend Property Get ", c_line, c_item, c_as:    c_kind_of_item = enPropertyGet
        Case c_line Like "Friend Property Let *":   mItems.Item "Friend Property Let ", c_line, c_item:            c_kind_of_item = enPropertyLet
        Case c_line Like "Friend Property Set *":   mItems.Item "Friend Property Set ", c_line, c_item:            c_kind_of_item = enPropertySet
        Case c_line Like "Public *":                mItems.Variable "Public ", c_line, c_item, c_as:               c_kind_of_item = enVariable
    End Select
            
    If c_item = vbNullString Then
        Debug.Print c_vbcm.Parent.name & " Line " & c_line_no & " ???? no item"
        Stop
    Else
    End If
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mLine" & "." & e_proc
End Function

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
        c_line_for_parsing = Replace(c_line_for_parsing, " .", " " & mStack.Top() & ".")
    End If
    
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsEndProc(ByVal c_line As String) As Boolean
    Dim s As String
    s = Trim(c_line)
    IsEndProc = s Like "End Property*" _
             Or s Like "End Function*" _
             Or s Like "End Sub*"

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

Public Function LineExcluded(ByVal l_line As String) As Boolean
    
    Dim v As Variant
    
    If VarType(vExcludedCodeLines) = vbArray Then
        For Each v In vExcludedCodeLines
            If l_line Like "*" & v & "*" Then
                LineExcluded = True
                Exit Function
            End If
        Next v
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
    SkipLinesEmptyOrComment n_as_of_line_no, sLine, n_vbcm
    If Len(sLine) = 0 Then Exit Function ' no line with a content found
    
    n_line_starts_at = n_as_of_line_no
    DropLineNumber sLine
    ConcatenateAllContinuationLines n_as_of_line_no, sLine, n_vbcm
    UnstripCommentIfAny sLine
    
    '~~ When the line does not have multiple code lines the reuturned line is the line as it is
    '~~ Else the returned line is the first of the multiple lines in the code line and the
    '~~ lNextSubLine is incremented by one for the next one being returned with the subsequent
    '~~ NextLine call
    NextLine = NextSubLine(sLine, lNextSubLine)
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
        If mItems.IsUniqueItem(sPublicItemItem) Then
            '~~ When unique this is the public item checked
            RefersToPublicItem = True
            GoTo xt
        End If
    End If
    
    If InStr(c_line_to_explore, " " & sPublicItemItem & " ") <> 0 Then
        If mItems.IsUniqueItem(sPublicItemItem) Then
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
            If mItems.IsUniqueItem(sPublicItemItem) Then
                '~~ Since the name is unique the unqualified call concerns to the Public item
                RefersToPublicItem = True
                GoTo xt
            ElseIf mProcs.IsLocal(c_using_vbc.name, sPublicItemItem) Then
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
            GoTo xt
    End Select
        
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

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

Private Sub ReplaceByDelimiter(ByRef r_line As String, _
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

Public Function IsLike(ByVal c_line As String, _
                      ByVal c_like As String, _
             Optional ByVal c_comp_using As String = vbNullString, _
             Optional ByVal c_comp As String) As Boolean
' ----------------------------------------------------------------------------
' For debuging only!
' ----------------------------------------------------------------------------
    If c_like <> vbNullString Then
        If c_comp_using <> vbNullString Then
            If c_line Like "*" & c_like & "*" _
            And c_comp_using = c_comp Then
                Stop
            End If
        Else
            IsLike = c_line Like "*" & c_like & "*"
        End If
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

