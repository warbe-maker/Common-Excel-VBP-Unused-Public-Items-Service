Attribute VB_Name = "mUnusedPublic"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mUnusedPublic
' -----------------------------
' The 'PublicUnused' service analyses the code of a Workbook for any unused
' public items by considering:
' - Public Constants
' - Public Sub, Function, Property
' - Methods and Properties of Class-Modules used through Class-Instances
' - Nested With xxx .... End With for a default Class-Instance object
' - Nested class instance calls x.y.z
' - Public prodedures used via an OnAction property
' - Public procedures called via Application.Run (provided the called
'   procedure is not 'hidden' in a Constant.
'
' The service supplements MZ-Tools' dead code analyses which excludes Public
' items.
'
' W. Rauschenberger
' ----------------------------------------------------------------------------
Public Enum enKindOfComponent
    enA
    enStandardModule
    enClassModule
    enWorkbook
    enWorksheet
    enUserForm
    enZ
End Enum

Public Enum enKindOfItem
    enA
    enClassInstance
    enConstant
    enEnumeration
    enFunction
    enMethod
    enPropertyGet
    enPropertyLet
    enPropertySet
    enSub
    enUserDefinedType
    enVariable
    enZ
End Enum

Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

'***App Window Constants***
Private Const WIN_NORMAL = 1         'Open Normal

'***Error Codes***
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const vbResume                  As Long = 6 ' return value (equates to vbYes)

Private WithStack               As Collection
Private lLenPublicItems         As Long
Private lLinesTotal             As Long
Private vbpServiced             As VBProject
Private sFile                   As String
Private lLinesExplored          As Long
Private lProcsTotal             As Long
Private sCompPublic             As String
Private sProcParsing            As String
Private sCompParsing            As String
Private sItemPublic             As String
Private dctClassModules         As Dictionary
Private dctClassStack           As Dictionary
Private dctCodeRefPublicItem    As Dictionary
Private dctExcluded             As Dictionary
Private dctInstncsCompGlobal    As Dictionary ' collection of VBComponent class instances
Private dctInstncsProcLocal     As Dictionary ' collection of Procedure local declared class instances
Private dctInstncsVBPrjctGlobal As Dictionary
Private dctKindOfItem           As Dictionary
Private dctOnActions            As Dictionary
Private dctPublicItemsUsed      As Dictionary   ' All Public items used
Private dctPublicItemsUnique    As Dictionary ' Collection of all those public items with a unique name

Public dctComps                 As Dictionary
Public dctProcLines             As Dictionary   ' All component's procedures with theit start and end line
Public dctProcs                 As Dictionary
Public dctPublicItems           As Dictionary   ' All Public ... and Friend ... - finally only those not used
Public dctUsed                  As Dictionary
Public dctUnused                As Dictionary
Public Excluded                 As String
Public wbkServiced              As Workbook

Private Sub GetOpen(ByVal g_wbk_full_name As String, _
                    ByRef g_wbk As Workbook)
    Dim wbk     As Workbook
    Dim fso     As New FileSystemObject
    Dim sName   As String
    
    sName = fso.GetFileName(g_wbk_full_name)
    For Each wbk In Application.Workbooks
        If wbk.name = sName Then
            Set g_wbk = wbk
            Exit For
        End If
    Next wbk

    If g_wbk Is Nothing Then
        Set g_wbk = Workbooks.Open(g_wbk_full_name)
    End If

End Sub

Private Property Get ClassStack(Optional ByVal c_class_mod_name As String) As clsStk
    If dctClassStack.Exists(c_class_mod_name) Then
        Set ClassStack = dctClassStack(c_class_mod_name)
    Else
        Set ClassStack = New clsStk
    End If
End Property

Private Property Let ClassStack(Optional ByVal c_class_mod_name As String, _
                                         ByVal c_class_stack As clsStk)
    If dctClassStack Is Nothing Then Set dctClassStack = New Dictionary
    If dctClassStack.Exists(c_class_mod_name) Then
        dctClassStack.Remove c_class_mod_name
    End If
    dctClassStack.Add c_class_mod_name, c_class_stack
End Property

Private Property Get FileTemp(Optional ByVal tmp_path As String = vbNullString, _
                              Optional ByVal tmp_extension As String = ".tmp") As String
' ----------------------------------------------------------------------------
' Returns the full file name of a temporary randomly named file, when tmp_path
' is omitted in the CurDir path.
' ----------------------------------------------------------------------------
    Dim fso     As New FileSystemObject
    Dim sTemp   As String
    
    If VBA.Left$(tmp_extension, 1) <> "." Then tmp_extension = "." & tmp_extension
    sTemp = Replace(fso.GetTempName, ".tmp", tmp_extension)
    If tmp_path = vbNullString Then tmp_path = ThisWorkbook.Path
    On Error Resume Next
    fso.DeleteFile tmp_path & "\rad*" & tmp_extension
    
    sTemp = VBA.Replace(tmp_path & "\" & sTemp, "\\", "\")
    FileTemp = sTemp
    Set fso = Nothing
    
End Property

Private Property Let FileText(Optional ByVal ft_file As Variant, _
                              Optional ByVal ft_append As Boolean = True, _
                              Optional ByRef ft_split As String, _
                                       ByVal ft_string As String)
' ----------------------------------------------------------------------------
' Writes the string (ft_string) into the file (ft_file) which might be a file
' object or a file's full name.
' Note: ft_split is not used but specified to comply with Property Get.
' ----------------------------------------------------------------------------
    Const PROC = "FileText-Let"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim ts  As TextStream
    Dim sFl As String
   
    ft_split = ft_split ' not used! just for coincidence with Get
    With fso
        If TypeName(ft_file) = "File" Then
            sFl = ft_file.Path
        Else
            '~~ ft_file is regarded a file's full name, created if not existing
            sFl = ft_file
            If Not .FileExists(sFl) Then .CreateTextFile sFl
        End If
        
        If ft_append _
        Then Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForAppending) _
        Else Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForWriting)
    End With
    
    ts.WriteLine ft_string

xt: ts.Close
    Set fso = Nothing
    Set ts = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get IsClassModule(Optional ByVal i_name As String, _
                                   Optional ByVal i_vbc As VBComponent) As Boolean
    Set i_vbc = i_vbc
    If Not dctClassModules Is Nothing Then
        IsClassModule = dctClassModules.Exists(i_name)
    End If
End Property

Private Property Let IsClassModule(Optional ByVal i_name As String, _
                                   Optional ByVal i_vbc As VBComponent, _
                                            ByVal i_is As Boolean)
                                            
    If dctClassModules Is Nothing Then Set dctClassModules = New Dictionary
    If i_is Then
        If Not dctClassModules.Exists(i_name) Then
            AddAscByKey dctClassModules, i_name, i_vbc
        End If
    End If
    
End Property
                             
Public Sub AddAscByKey(ByRef add_dct As Dictionary, _
                       ByVal add_key As Variant, _
                       ByVal add_item As Variant)
' ------------------------------------------------------------------------------------
' Adds to the Dictionary (add_dct) an item (add_item) in ascending order by the key
' (add_key). When the key is an object with no Name property an error is raisede.
'
' Note: This is a copy of the DctAdd procedure with fixed options which may be copied
'       into any VBProject's module in order to have it independant from this
'       Common Component.
'
' W. Rauschenberger, Berlin Jan 2022
' ------------------------------------------------------------------------------------
    Const PROC = "AddAscByKey"
    
    On Error GoTo eh
    Dim bDone           As Boolean
    Dim dctTemp         As Dictionary
    Dim vItemExisting   As Variant
    Dim vKeyExisting    As Variant
    Dim vValueExisting  As Variant ' the entry's add_key/add_item value for the comparison with the vValueNew
    Dim vValueNew       As Variant ' the argument add_key's/add_item's value
    Dim bStayWithFirst  As Boolean
    Dim bOrderByItem    As Boolean
    Dim bOrderByKey     As Boolean
    Dim bSeqAscending   As Boolean
    Dim bCaseIgnored    As Boolean
    Dim bCaseSensitive  As Boolean
    Dim bEntrySequence  As Boolean
    
    If add_dct Is Nothing Then Set add_dct = New Dictionary
    
    '~~ Plausibility checks
    bOrderByItem = False
    bOrderByKey = True
    bSeqAscending = True
    bCaseIgnored = False
    bCaseSensitive = True
    bStayWithFirst = True
    bEntrySequence = False
    
    With add_dct
        '~~ When it is the very first add_item or the add_order option
        '~~ is entry sequence the add_item will just be added
        If .Count = 0 Or bEntrySequence Then
            .Add add_key, add_item
            GoTo xt
        End If
        
        '~~ When the add_order is by add_key and not stay with first entry added
        '~~ and the add_key already exists the add_item is updated
        If bOrderByKey And Not bStayWithFirst Then
            If .Exists(add_key) Then
                If IsObject(add_item) Then Set .Item(add_key) = add_item Else .Item(add_key) = add_item
                GoTo xt
            End If
        End If
    End With
        
    '~~ When the add_order argument is an object but does not have a name property raise an error
    If bOrderByKey Then
        If IsObject(add_key) Then
            On Error Resume Next
            add_key.name = add_key.name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(7), ErrSrc(PROC), "The add_order option is by add_key, the add_key is an object but does not have a name property!"
        End If
    ElseIf bOrderByItem Then
        If IsObject(add_item) Then
            On Error Resume Next
            add_item.name = add_item.name
            If Err.Number <> 0 _
            Then Err.Raise AppErr(8), ErrSrc(PROC), "The add_order option is by add_item, the add_item is an object but does not have a name property!"
        End If
    End If
    
    vValueNew = AddAscByKeyValue(add_key)
    
    With add_dct
        '~~ Get the last entry's add_order value
        vValueExisting = AddAscByKeyValue(.Keys()(.Count - 1))
        
        '~~ When the add_order mode is ascending and the last entry's add_key or add_item
        '~~ is less than the add_order argument just add it and exit
        If bSeqAscending And vValueNew > vValueExisting Then
            .Add add_key, add_item
            GoTo xt
        End If
    End With
        
    '~~ Since the new add_key/add_item couldn't simply be added to the Dictionary it will
    '~~ be inserted before or after the add_key/add_item as specified.
    Set dctTemp = New Dictionary
    bDone = False
    
    For Each vKeyExisting In add_dct
        
        If IsObject(add_dct.Item(vKeyExisting)) _
        Then Set vItemExisting = add_dct.Item(vKeyExisting) _
        Else vItemExisting = add_dct.Item(vKeyExisting)
        
        With dctTemp
            If bDone Then
                '~~ All remaining items just transfer
                .Add vKeyExisting, vItemExisting
            Else
                vValueExisting = AddAscByKeyValue(vKeyExisting)
            
                If vValueExisting = vValueNew And bOrderByItem And bSeqAscending And Not .Exists(add_key) Then
                    If bStayWithFirst Then
                        .Add vKeyExisting, vItemExisting:   bDone = True ' not added
                    Else
                        '~~ The add_item already exists. When the add_key doesn't exist and bStayWithFirst is False the add_item is added
                        .Add vKeyExisting, vItemExisting:   .Add add_key, add_item:                     bDone = True
                    End If
                ElseIf bSeqAscending And vValueExisting > vValueNew Then
                    .Add add_key, add_item:                 .Add vKeyExisting, vItemExisting:   bDone = True
                Else
                    .Add vKeyExisting, vItemExisting ' transfer existing add_item, wait for the one which fits within sequence
                End If
            End If
        End With ' dctTemp
    Next vKeyExisting
    
    '~~ Return the temporary dictionary with the new add_item added and all exiting items in add_dct transfered to it
    Set add_dct = dctTemp
    Set dctTemp = Nothing

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AddAscByKeyValue(ByVal add_key As Variant) As Variant
' ----------------------------------------------------------------------------
' When add_key is an object its name becomes the sort order value else the
' the value is returned as is.
' ----------------------------------------------------------------------------
    If VarType(add_key) = vbObject Then
        On Error Resume Next ' the object may not have a Name property
        AddAscByKeyValue = add_key.name
        If Err.Number <> 0 Then Set AddAscByKeyValue = add_key
    Else
        AddAscByKeyValue = add_key
    End If
End Function

Private Function Align(ByVal align_s As String, _
                       ByVal align_lngth As Long, _
              Optional ByVal align_mode As StringAlign = AlignLeft, _
              Optional ByVal align_margin As String = vbNullString, _
              Optional ByVal align_fill As String = " ") As String
' ---------------------------------------------------------
' Returns a string (align_s) with a lenght (align_lngth)
' aligned (aligned) filled with characters (align_fill).
' ---------------------------------------------------------
    Dim SpaceLeft       As Long
    
    Select Case align_mode
        Case AlignLeft
            If Len(align_s & align_margin) >= align_lngth _
            Then Align = VBA.Left$(align_s & align_margin, align_lngth) _
            Else Align = align_s & align_margin & VBA.String$(align_lngth - (Len(align_s & align_margin)), align_fill)
        Case AlignRight
            If Len(align_margin & align_s) >= align_lngth _
            Then Align = VBA.Left$(align_margin & align_s, align_lngth) _
            Else Align = VBA.String$(align_lngth - (Len(align_margin & align_s)), align_fill) & align_margin & align_s
        Case AlignCentered
            If Len(align_margin & align_s & align_margin) >= align_lngth Then
                Align = align_margin & Left$(align_s, align_lngth - (2 * Len(align_margin))) & align_margin
            Else
                SpaceLeft = Max(1, ((align_lngth - Len(align_s) - (2 * Len(align_margin))) / 2))
                Align = VBA.String$(SpaceLeft, align_fill) & align_margin & align_s & align_margin & VBA.String$(SpaceLeft, align_fill)
                Align = VBA.Right$(Align, align_lngth)
            End If
    End Select

End Function

Private Function WbkSelect() As String
    Dim fDialog As FileDialog
    Dim result  As Integer

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    'Optional: FileDialog properties
    fDialog.AllowMultiSelect = False
    fDialog.Title = "Select a Workbook (will be opened when not open)"
    fDialog.InitialFileName = "C:\"
    'Optional: Add filters
    fDialog.Filters.Clear
    fDialog.Filters.Add "Excel files", "*.xls*"
    fDialog.Filters.Add "All files", "*.*"
 
    'Show the dialog. -1 means success!
    If fDialog.Show = -1 Then
       WbkSelect = fDialog.SelectedItems(1)
    End If
End Function

Public Sub PublicUnused()
' ------------------------------------------------------------------------------
' - Select a Workbook and decide on which VBComponents to exclude or include,
' - Collect all unused Public items
' - Display a list of all unused/used Public items
' ------------------------------------------------------------------------------
    Const PROC  As String = "AnalysePublicUnused"
    
    On Error GoTo eh
    Dim cll                 As Collection
    Dim lMaxCompProcKind    As Long
    Dim lMaxKindOfComp      As Long
    Dim lMaxKindOfItem      As Long
    Dim lMaxLenItems        As Long
    Dim lMaxPublic          As Long
    Dim lMaxUsing           As Long
    Dim s                   As String
    Dim sComp               As String
    Dim sProc               As String
    Dim vPublic             As Variant
    Dim sWbk                As String
    
    BoP ErrSrc(PROC)
    sWbk = WbkSelect
    If sWbk = vbNullString Then GoTo xt
    GetOpen sWbk, wbkServiced
    If wbkService Is Nothing Then GoTo xt
    fExcludeInclude.Show ' assembles in Excluded the ignored VBComponents
        
    PublicItemsUsageCollect
    
    lMaxKindOfComp = MaxKindOfComp
    lMaxKindOfItem = MaxKindOfItem
    lMaxLenItems = MaxLenItems(dctPublicItems)
    
    lMaxCompProcKind = lMaxKindOfComp + lMaxKindOfItem + 3
    
    s = "The following " & dctPublicItems.Count & " Public declared items are  u n u s e d ! *)"
    WriteToFile s
    WriteToFile String(Len(s), "-")
    s = Align("Kind of Component.Item", lMaxCompProcKind, AlignCentered, , "-") & _
        " " & _
        Align("Public item (component.item)", lMaxLenItems, AlignCentered, , "-")
    WriteToFile s
    s = String(lMaxCompProcKind, "-") & " " & String(lMaxLenItems, "-")
    
    For Each vPublic In dctPublicItems
        Set cll = dctPublicItems(vPublic)
        sComp = Split(vPublic, ".")(0)
        sProc = Split(vPublic, ".")(1)
        s = "(" & PublicItemCollKindOfCompItem(cll) & ")"
        WriteToFile Align(s, lMaxCompProcKind, , " ") & vPublic
    Next vPublic
    
    lMaxLenItems = MaxLenItems(dctUsed)
    WriteToFile vbNullString
    WriteToFile "*) Public items are not analysed in their own component."
    s = "   I.e. an unused Public item may still be used within its own Component."
    WriteToFile s
    WriteToFile "   In case the Public item should rather be turned into Private!"
    WriteToFile String(Len(s), "=")
    WriteToFile vbLf
    s = "The following " & dctUsed.Count & " Public declared items had been found in at least one code line:"
    WriteToFile s
    WriteToFile String(Len(s), "-")
    
    For Each vPublic In dctUsed
        Set cll = dctUsed(vPublic)
        lMaxPublic = Max(lMaxPublic, Len(vPublic))
        lMaxUsing = Max(lMaxUsing, Len(cll(5)))
    Next vPublic
    
    WriteToFile Align("Public item used", lMaxPublic + 1, AlignLeft, , ".") & Align("Used by (example)", lMaxUsing + 1, AlignLeft, , ".") & ": " & "Code line"
    
    For Each vPublic In dctUsed
        Set cll = dctUsed(vPublic)
        sComp = Split(vPublic, ".")(0)
        sProc = Split(vPublic, ".")(1)
        WriteToFile Align(vPublic, lMaxPublic + 1, AlignLeft, , ".") & Align(cll(5), lMaxUsing + 1, AlignLeft, , ".") & ": " & cll(7)
    Next vPublic
        
    OpenUrlEtc sFile, WIN_NORMAL
    
xt: EoP ErrSrc(PROC)

#If ExecTrace = 1 Then
    mTrc.Dsply
#End If
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error.
' - Returns a given positive 'Application Error' number (app_err_no) into a
'   negative by adding the system constant vbObjectError
' - Returns the original 'Application Error' number when called with a negative
'   error number.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Sub BoC(ByVal boc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(C)ode with id (boc_id) trace. Procedure to be copied as Private
' into any module potentially using the Common VBA Execution Trace Service. Has
' no effect when Conditional Compile Argument is 0 or not set at all.
' Note: The begin id (boc_id) has to be identical with the paired EoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.BoC boc_id, s
#End If
End Sub

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

Private Function CodeLineDeclaresInstanceGlobal(ByRef c_line As String, _
                                                ByRef c_item As String, _
                                                ByRef c_as As String) As Boolean
' ------------------------------------------------------------------------------------
' When the code line (c_line) declares an item where "As" names a known Class Module
' the function returns:
' - TRUE
' - The declared item (c_item)
' - The As string (c_as)
' ------------------------------------------------------------------------------------
    Const PROC = "CodeLineDeclaresInstanceGlobal"
    
    On Error GoTo eh
    
    If c_line Like "* As New *" Then
'        Debug.Print c_vbcm.Parent.name & ", Line number=" & c_line_no & ", Line= >" & c_line & "<"
        ItemAs c_line, c_as
        CodeLineDeclaresInstanceGlobal = IsClassModule(c_as)
        If Not CodeLineDeclaresInstanceGlobal Then GoTo xt
        c_item = Trim(Split(c_line, " As New ")(0))
        c_item = Trim(Split(c_item, " ")(UBound(Split(c_item, " "))))
    ElseIf c_line Like "* As *" Then
'        Debug.Print c_vbcm.Parent.name & ", Line number=" & c_line_no & ", Line= >" & c_line & "<"
        ItemAs c_line, c_as
        CodeLineDeclaresInstanceGlobal = IsClassModule(c_as)
        If Not CodeLineDeclaresInstanceGlobal Then GoTo xt
        c_item = Trim(Split(c_line, " As ")(0))
        c_item = Trim(Split(c_item, " ")(UBound(Split(c_item, " "))))
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CodeLineDeclaresPublicItem(ByRef c_line_no As Long, _
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
    Const PROC = "CodeLineDeclaresPublicItem"
    
    On Error GoTo eh
    Dim sLine       As String
    Dim sComp       As String
    Dim sItem       As String
    Dim lLoopStop   As Long
    Dim lStartsAt   As Long
    If Not c_line Like "Public *" Then GoTo xt
    CodeLineDeclaresPublicItem = True
'    Debug.Print c_vbcm.Parent.name & ", Line number=" & c_line_no & ", Line= >" & c_line & "<"
    
    Select Case True
        Case c_line Like "Public Type *"
            c_line_no = c_line_no + 1
            CodeLineNext c_line_no, c_vbcm, sLine, lStartsAt
            sComp = c_vbcm.Parent.name
            lLoopStop = 0
            While Not sLine Like "End Type*"
                lLoopStop = lLoopStop + 1
                If lLoopStop > 100 Then Stop
                sLine = sLine & " "
                sItem = Split(sLine, " ")(0)
                PublicItemCollect sComp, sItem, c_line_no, sLine, c_kind_of_comp, enUserDefinedType
                c_line_no = c_line_no + 1
                CodeLineNext c_line_no, c_vbcm, sLine, lStartsAt
            Wend
            CodeLineDeclaresPublicItem = False
            GoTo xt

        Case c_line Like "Public Enum *"
            c_line_no = c_line_no + 1
            CodeLineNext c_line_no, c_vbcm, sLine, lStartsAt
            sComp = c_vbcm.Parent.name
            lLoopStop = 0
            While Not sLine Like "End Enum*"
                lLoopStop = lLoopStop + 1
                If lLoopStop > 500 Then Stop
                sLine = sLine & " "
                sItem = Split(sLine, " ")(0)
                PublicItemCollect sComp, sItem, c_line_no, sLine, c_kind_of_comp, enEnumeration
                c_line_no = c_line_no + 1
                CodeLineNext c_line_no, c_vbcm, sLine, lStartsAt
            Wend
            CodeLineDeclaresPublicItem = False
            GoTo xt
            
        Case c_line Like "Public Const *":          PublicItem "Public Const ", c_line, c_item:                 c_kind_of_item = enConstant
        Case c_line Like "Public Sub *"
                                                    If sComp = "mExport" Then Stop
                                                    PublicItem "Public Sub ", c_line, c_item:                   c_kind_of_item = enSub
        Case c_line Like "Public Function *":       PublicItemAs "Public Function ", c_line, c_item, c_as:      c_kind_of_item = enFunction
        Case c_line Like "Public Property Get *":   PublicItemAs "Public Property Get ", c_line, c_item, c_as:  c_kind_of_item = enPropertyGet
        Case c_line Like "Public Property Let *":   PublicItem "Public Property Let ", c_line, c_item:          c_kind_of_item = enPropertyGet
        Case c_line Like "Public Property Set *":   PublicItem "Public Property Set ", c_line, c_item:          c_kind_of_item = enPropertySet
        Case c_line Like "Friend Property Get *":   PublicItemAs "Friend Property Get ", c_line, c_item, c_as:  c_kind_of_item = enPropertyGet
        Case c_line Like "Friend Property Let *":   PublicItem "Friend Property Let ", c_line, c_item:          c_kind_of_item = enPropertyLet
        Case c_line Like "Friend Property Set *":   PublicItem "Friend Property Set ", c_line, c_item:          c_kind_of_item = enPropertySet
        Case c_line Like "Public *":                PublicVariable "Public ", c_line, c_item, c_as:             c_kind_of_item = enVariable
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

Private Sub CodeLineDropLineNumber(ByRef c_line As String)
        
    Dim v As Variant
    If Len(c_line) > 1 Then
        v = Split(c_line, " ")
        If IsNumeric(v(0)) Then
            c_line = Trim(Replace(c_line, v(0), vbNullString))
        End If
    End If

End Sub

Private Function CodeLineForBeingParsed(ByVal c_line_to_parse As String, _
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
    Const PROC = "CodeLineForBeingParsed"
    
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
    CodeLineReplaceByDelimiter c_line_for_parsing, " | ", "(", ")", ", ", ":=", " = ", "!", """"
    c_line_for_parsing = " " & Trim(c_line_for_parsing) & " "
    
    '~~ Replace a "Me." by the component's name
    c_line_for_parsing = Replace(c_line_for_parsing, " Me.", " " & c_comp_to_parse & ".")
    
    '~~ Replace a default object (" .") with the corresponding With item
    If InStr(c_line_for_parsing, " .") <> 0 Then
        c_line_for_parsing = Replace(c_line_for_parsing, " .", " " & WithStackTop() & ".")
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
                    If IsClassInstance(sComp, sItem, sClass, c_proc_to_parse) Then
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

xt: CodeLineForBeingParsed = c_line_for_parsing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CodeLineIgnore(ByVal c_line As String) As String
' ------------------------------------------------------------------------------------
' Returns the string of the code line's item. Examples:
' ... Const XXX = "abc" returns XXX
' ... Property Get XXX ... returns XXX
' ... Function XXX() As .. returns XXX
' ------------------------------------------------------------------------------------
    Const PROC = "CodeLineItem"
    Dim sItem   As String
    
    If UBound(Split(" " & c_line & " ", " Property Get ")) > 0 Then
        sItem = Split(" " & c_line & " ", " Property Get ")(1)
        sItem = Split(sItem, "(")(0)
    ElseIf UBound(Split(" " & c_line & " ", " Function ")) > 0 Then
        sItem = Split(" " & c_line & " ", " Function ")(1)
        sItem = Split(sItem, "(")(0)
    End If
    CodeLineIgnore = sItem & " = "
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CodeLineIsEndProc(ByVal c_line As String) As Boolean
    CodeLineIsEndProc = c_line Like "*End Property*" _
                     Or c_line Like "*End Function*" _
                     Or c_line Like "*End Sub*"

End Function

Private Function CodeLineIsFirstOfProc(ByVal c_proc_of_line As String, _
                                       ByRef c_proc As String, _
                                       ByVal c_kind_of_proc As vbext_ProcKind) As Boolean
    Static sProc As String
    
    If c_proc_of_line & "." & KoPstring(c_kind_of_proc) <> sProc Then
        CodeLineIsFirstOfProc = True
        sProc = c_proc_of_line & "." & KoPstring(c_kind_of_proc)
        c_proc = sProc
    End If
    
End Function

Private Function CodeLineNext(ByRef c_i As Long, _
                              ByVal c_vbcm As CodeModule, _
                              ByRef c_line As String, _
                     Optional ByRef c_line_starts_at As Long) As Boolean
' ------------------------------------------------------------------------------------
' Returns the next code line (c_line) as of a start line (c_i) by:
' - skipping empty lines
' - kipping comment lines
' - unstripping comments
' - replacing constants with ""
' - concatenating continuation lines.
' ------------------------------------------------------------------------------------
    Const PROC = "CodeLineNext"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim lStopLoop   As Long
    Dim sComp       As String
    Dim vItems      As Variant
    Dim i           As Long
    
    sComp = c_vbcm.Parent.name
    With c_vbcm
        c_line = Trim(.Lines(c_i, 1))

        CodeLineDropLineNumber c_line
        If Len(c_line) > 1 Then
            v = Split(c_line, " ")
            If IsNumeric(v(0)) Then
                c_line = Trim(Replace(c_line, v(0), vbNullString))
            End If
        End If
        
        lStopLoop = 0
        While (Len(c_line) = 0 Or c_line Like "'*") And c_i <= .CountOfLines
            lStopLoop = lStopLoop + 1
            If lStopLoop > 100 Then Stop
            c_i = c_i + 1
            c_line = Trim(.Lines(c_i, 1))
        Wend
        If Len(c_line) = 0 Then Exit Function ' no line with a content found
        
        '~~ Concatenate continuation lines
        lStopLoop = 0
        c_line_starts_at = c_i
        While Right(Trim(c_line), 1) = "_" And c_i <= .CountOfLines
            lStopLoop = lStopLoop + 1
            If lStopLoop > 50 Then Stop
            c_i = c_i + 1   ' concatenate the following continuation line
            c_line = Left(c_line, Len(c_line) - 1) & Trim(.Lines(c_i, 1))
            c_line = Replace(c_line, " ,", ",")
        Wend
        
        '~~ Unstrip comment (any space followed by a ' which is not within a "")
        vItems = Split(c_line, """")
        For i = LBound(vItems) To UBound(vItems)
            Select Case i
                Case 0, 2, 4, 6, 8, 10, 12, 14
                    If InStr(vItems(i), " '") <> 0 Then
                        c_line = Trim(Split(c_line, " '")(0))
                        Exit For
                    End If
            End Select
        Next i
            
        '~~ Eliminate strings in "" except the line is an "Application.Run..." one
        If Not c_line Like "*Application.Run*" Then
            v = Split(Trim(c_line), """")
            lStopLoop = 0
            While UBound(v) >= 1
                lStopLoop = lStopLoop + 1
                If lStopLoop > 50 Then Stop
                c_line = Replace(c_line, """" & v(1) & """", vbNullString) & """"
                c_line = Replace(c_line, """", vbNullString)
                v = Split(c_line, """")
            Wend
        Else
'            Stop
        End If
    End With
    
    CodeLineNext = c_line <> vbNullString

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function CodeLineRefersToPublicItem(ByVal c_line_to_explore As String, _
                                            ByVal c_public_item As String, _
                                            ByVal c_using_vbc As VBComponent, _
                                            ByVal c_dot_class As String) As Boolean
' ------------------------------------------------------------------------------------
' Returns TRUE when the code line (c_line_to_explore) contains the 'Public Item' (c_public_item).
' ------------------------------------------------------------------------------------
    Const PROC = "CodeLineRefersToPublicItem"
    
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
        CodeLineRefersToPublicItem = True
        GoTo xt
    End If
             
    If InStr(c_line_to_explore, "." & sPublicItemItem & " ") <> 0 Then
        If IsUniquePublicItem(sPublicItemItem) Then
            '~~ When unique this is the public item checked
            CodeLineRefersToPublicItem = True
            GoTo xt
        End If
    End If
    
    If InStr(c_line_to_explore, " " & sPublicItemItem & " ") <> 0 Then
        If IsUniquePublicItem(sPublicItemItem) Then
            '~~ when unique even unqualified is a clear usage indication
            CodeLineRefersToPublicItem = True
            GoTo xt
        End If
    End If
        
    '~~ 3. Check if the code line contains the <proc> unqualified
'    If c_line_to_explore Like "* " & sPublicItemItem & " *" _
'    Or c_line_to_explore Like "*." & sPublicItemItem & " *" Then Stop
    
    Select Case True
        Case c_line_to_explore Like "* " & sPublicItemItem & " *"
            '~~ This is an unqualified call of a Public item
            If IsUniquePublicItem(sPublicItemItem) Then
                '~~ Since the name is unique the unqualified call concerns to the Public item
                CodeLineRefersToPublicItem = True
                GoTo xt
            ElseIf IsLocalProcedure(c_using_vbc.name, sPublicItemItem) Then
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
                CodeLineRefersToPublicItem = True
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
                    CodeLineRefersToPublicItem = True
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

Private Sub CodeLineReplaceByDelimiter(ByRef r_line As String, _
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

Private Sub CodeLineStopIfLike(ByVal c_line As String, _
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

Public Sub Collect()
' ------------------------------------------------------------------------------
' Collects (in Dictionary with comp-name.item-name as key):
' 1. All those VBComponenKoItemts not explicitely excluded (dctComponenKoItemts)
'    1.1 All Class Modules (dctClassModules)
' 2. All items declared Public (dctPublicItems)
' 3. All public item's Kind (dctKindOfItem)
' 4. All public items indicating unique True or False (dctPublicItemsUnique)
' 5. All VBComponenKoItemt's Procedures with their start and enKoItemd line (dctProcLines)
' ------------------------------------------------------------------------------
    Const PROC  As String = "Collect"
    
    On Error GoTo eh
    Dim cllComp         As Collection
    Dim cllProc         As Collection
    Dim i               As Long
    Dim sAs             As String
    Dim sComp           As String
    Dim sItem           As String
    Dim sLine           As String
    Dim sProc           As String
    Dim vComp           As Variant
    Dim vbc             As VBComponent
    Dim vbcm            As CodeModule
    Dim dctCompProcs    As Dictionary
    Dim vProc           As Variant
    Dim enKoItem        As enKindOfItem
    Dim enKoComp        As enKindOfComponent
    
    BoP ErrSrc(PROC)
    CompsCollect Excluded
    ProcsCollect           ' Collect all procedures in not exluded VBComponenKoItemts
    CollectInstncsVBPGlobal   ' Collect all class instance which are VB-Project global
    CollectInstncsCompGlobal  ' Collect all class instances which are ComponenKoItemt global
    CollectInstncsProcLocal   ' Collect all class instances in Procedures
    CollectOnActions
    
    For Each vComp In dctComps
        sComp = vComp
        Set cllComp = dctComps(sComp)
        Set vbc = CompCollVBC(cllComp)
        enKoComp = CompCollKind(cllComp)
        Set vbcm = vbc.CodeModule
        
        Set dctCompProcs = dctProcs(sComp)
        For Each vProc In dctCompProcs
            sProc = vProc
            Set cllProc = dctCompProcs(vProc)
            sLine = ProcCollLine(cllProc)
            i = ProcCollLineFrom(cllProc)
            If CodeLineDeclaresPublicItem(i, sLine, sItem, sAs, vbcm, enKoComp, enKoItem) Then
                PublicItemCollect sComp, sItem, i, sLine, enKoComp, enKoItem
            End If
        Next vProc
    Next vComp
        
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub CollectExcluded(ByVal s As String)
' ------------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------------
    Dim v As Variant
    Dim dct As New Dictionary
    
    dct.Add "mUnusedPublic", vbNullString ' The component excludes itself
    If s <> vbNullString Then
        For Each v In Split(s, ",")
            dct.Add Trim(v), vbNullString
        Next v
    End If
    Set dctExcluded = dct
    Set dct = Nothing

End Sub

Private Sub CollectInstncsCompGlobal()
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
        
        For i = 1 To vbcm.CountOfDeclarationLines
            CodeLineNext i, vbcm, sLine, lStartsAt
            lStopLoop = 0
            While CodeLineDeclaresInstanceGlobal(sLine, sItem, sAs)
                lStopLoop = lStopLoop + 1
                If lStopLoop > 50 Then Stop
                If IsClassModule(sAs) Then
                    If Not dctInstncsCompGlobal.Exists(sItem) Then
                        dct.Add sItem, sAs
                    End If
                    sLine = Trim(Replace(Replace(sLine, " As " & sAs, vbNullString, 1, 1), sItem, vbNullString, 1, 1))
                End If
            Wend
        Next i
        
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

Private Sub CollectInstncsProcLocal()
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
            Set vbcm = ProcCollVBCM(cllProc)
            lFrom = ProcCollLineFrom(cllProc)
            lTo = ProcCollLineTo(cllProc)
            Set cllProc = Nothing
            Set dctInstProc = New Dictionary
            With vbcm
                For i = lFrom To lTo
                    CodeLineNext i, vbcm, sLine, lStartsAt
'                    If InStr(sLine, "ByRef a_stats As clsStats = Nothing") <> 0 Then Stop
                    Select Case True
                        Case sLine Like "* As New *"
                            sItem = Split(Trim(Split(sLine, " As New ")(0)), " ")(UBound(Split(Trim(Split(sLine, " As ")(0)), " ")))
                            sAs = Split(Trim(Split(sLine, " As New ")(1)), " ")(0)
                            If IsClassModule(sAs) Then
                                dctInstProc.Add sItem, sAs
                            End If
                        Case sLine Like "* As *"
                            sItem = Split(Trim(Split(sLine, " As ")(0)), " ")(UBound(Split(Trim(Split(sLine, " As ")(0)), " ")))
                            sAs = Split(Trim(Split(sLine, " As ")(1)), " ")(0)
                            If IsClassModule(sAs) Then
                                If Not dctInstProc.Exists(sItem) Then
                                    dctInstProc.Add sItem, sAs
                                End If
    '                            Debug.Print "Local instance '" & sItem & "' of Class '" & sAs & "' (" & sLine & ")"
                            End If
                    End Select
                Next i
            End With
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

                                            
Private Sub CollectInstncsVBPGlobal()
    
    Dim cll         As Collection
    Dim dct         As Dictionary
    Dim wsh         As Worksheet
    Dim v           As Variant
    Dim sComp       As String
    Dim vbc         As VBComponent
    Dim vbcm        As CodeModule
    Dim i           As Long
    Dim sItem       As String
    Dim sAs         As String
    Dim sLine       As String
    Dim enKoItem    As enKindOfItem
    Dim enKoComp    As enKindOfComponent
    
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
        With vbcm
            For i = 1 To .CountOfDeclarationLines
                CodeLineNext i, vbcm, sLine
                If sLine Like "Public *" Then
                    If CodeLineDeclaresPublicItem(i, sLine, sItem, sAs, vbcm, enKoComp, enKoItem) Then
                        If IsClassModule(sAs, vbc) Then
                            AddAscByKey dct, sItem, sAs
                        Else
                            PublicItemCollect sComp, sItem, i, sLine, enKoComp, enKoItem
                        End If
                        
                    End If
                End If
            Next i
        End With
        AddAscByKey dctInstncsVBPrjctGlobal, sComp, dct
        Set dct = Nothing
    Next v
            
End Sub

Private Sub CollectOnActions()
' ------------------------------------------------------------------------------------
' Collects all OnActions (which must be Public items) in a Dictionary (dctOnActions)
' with the OnAction as key.
' ------------------------------------------------------------------------------------
    
    Static dct      As Dictionary
    Dim shp         As Shape
    Dim sOnAction   As String
    Dim wsh         As Worksheet
    
    Set dct = New Dictionary
    For Each wsh In wbkServiced.Worksheets
        For Each shp In wsh.Shapes
            sOnAction = vbNullString
            On Error Resume Next
            sOnAction = shp.OnAction
            If sOnAction <> vbNullString Then
                sOnAction = Split(sOnAction, "!")(1)
                If Not dct.Exists(sOnAction) Then
                    AddAscByKey dct, sOnAction, vbNullString
                End If
            End If
        Next shp
    Next wsh
    
    Set dctOnActions = dct
    Set dct = Nothing
    
End Sub

Private Function CompCollKind(ByVal cll As Collection) As enKindOfComponent:    CompCollKind = cll(2):  End Function

Private Function CompCollVBC(ByVal cll As Collection) As VBComponent:   Set CompCollVBC = cll(1):   End Function

Public Sub CompsCollect(ByVal c_excluded As String)
' ------------------------------------------------------------------------------------
' Provides a Dictionary (dctComps) with all components not excluded
' XrefVBProject with the VBComponent's Name as the key and the VBComponent as item.
' ------------------------------------------------------------------------------------
    Const PROC  As String = "CompsCollect"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim enKind  As enKindOfComponent
    Dim sComp   As String
    Dim cll     As Collection
    
    BoP ErrSrc(PROC)
    Set dctComps = New Dictionary
    CollectExcluded c_excluded
    
    For Each vbc In wbkServiced.VBProject.VBComponents
        Set cll = New Collection
        cll.Add vbc
        With vbc
            sComp = .name
            Select Case .Type
                Case vbext_ct_ClassModule
                    enKind = enClassModule
                Case vbext_ct_Document
                    If IsSheet(sComp, wbkServiced) _
                    Then enKind = enWorksheet _
                    Else enKind = enWorkbook
                Case vbext_ct_MSForm:       enKind = enUserForm
                Case vbext_ct_StdModule:    enKind = enStandardModule
            End Select
        
            If Not IsExcluded(sComp) Then
                If Not dctComps.Exists(sComp) Then
                    cll.Add enKind
                    AddAscByKey dctComps, .name, cll
                    Set cll = Nothing
                End If
                If .Type = vbext_ct_ClassModule Then
                    IsClassModule(sComp, vbc) = True
                End If
            End If
        End With
    Next vbc

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub EoC(ByVal eoc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(C)ode id (eoc_id) trace. Procedure to be copied as Private into
' any module potentially using the Common VBA Execution Trace Service. Has no
' effect when the Conditional Compile Argument is 0 or not set at all.
' Note: The end id (eoc_id) has to be identical with the paired BoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.EoC eoc_id, s
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
    ErrSrc = "mUnusedPublic" & "." & e_proc
End Function

Private Sub Initialize()

    Set dctInstncsCompGlobal = Nothing  ' All VBComponent global declared class instances
    Set dctInstncsProcLocal = Nothing   ' All Procedure local declared class instances
    Set dctProcs = Nothing              ' All Procedures if non-excluded VBComponents
    Set dctPublicItemsUsed = Nothing    ' All Public items used
    Set dctPublicItemsUnique = Nothing  ' Collection of all those public items with a unique name
    Set dctComps = Nothing              ' All bot excluded VBComponents
    Set dctProcLines = Nothing          ' All component's procedures with theit start and end line
    Set dctPublicItems = Nothing        ' All Public ... and Friend ... - finally only those unused
    
    sFile = vbNullString
    Set dctPublicItems = New Dictionary
    Set dctKindOfItem = New Dictionary
    Set dctPublicItemsUnique = New Dictionary
    Set dctProcLines = New Dictionary
    
End Sub

Private Function IsClassInstance(ByVal i_comp_name As String, _
                                 ByVal i_instance_name As String, _
                                 ByRef i_class_name As String, _
                        Optional ByVal i_proc_name As String = vbNullString) As Boolean
' ------------------------------------------------------------------------------
' When the instance (i_instance_name) is a known Class instance the function
' returns TRUE and the corresponding class' name (i_class_name)
' ------------------------------------------------------------------------------
    Const PROC = "IsClassInstance"
    
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
                    IsClassInstance = True
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
            IsClassInstance = True
            GoTo xt
        End If
    End If
    
    '~~ When the i_instance is not known as a VBComponent global class instance
    '~~ it may still be a VBProject global class instance declared in a VBComponent
    If dctInstncsVBPrjctGlobal.Exists(i_comp_name) Then
        Set dct = dctInstncsVBPrjctGlobal(i_comp_name)
        If dct.Exists(i_instance_name) Then
            i_class_name = dct(i_instance_name)
            IsClassInstance = True
            GoTo xt
        End If
    End If
            
    For Each v In dctInstncsVBPrjctGlobal
        Set dct = dctInstncsVBPrjctGlobal(v)
        If dct.Exists(i_instance_name) Then
            i_class_name = dct(i_instance_name)
            IsClassInstance = True
            GoTo xt
        End If
    Next v
    
    '~~ When the i_instance is not known as a VBComponent global class instance declared in a VBComponent
    '~~ it may still be a class instance like the Workbook itself or any of its Worksheets
    If dctInstncsVBPrjctGlobal.Exists(vbNullString) Then
        Set dct = dctInstncsVBPrjctGlobal(vbNullString)
        If dct.Exists(i_instance_name) Then
            i_class_name = dct(i_instance_name)
            IsClassInstance = True
            GoTo xt
        End If
    End If
            
            
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsExcluded(ByVal i_comp_name As String) As Boolean
    IsExcluded = dctExcluded.Exists(i_comp_name)
End Function

Private Function IsInstanceLocal(ByVal c_comp As String, _
                                 ByVal c_proc As String, _
                                 ByVal c_instance As String, _
                                 ByRef c_class As String) As Boolean
' ------------------------------------------------------------------------------
' When the instance (c_instance) exists in the <comp>.<proc> the function
' returns TRUE and the name of the Class-Module of the instance (c_class).
' ------------------------------------------------------------------------------
    Const PROC = "IsInstanceLocal"
    
    On Error GoTo eh
    Dim dct     As Dictionary
    Dim sKey    As String
    
    sKey = c_comp & "." & c_proc
    If dctInstncsProcLocal.Exists(sKey) Then
        Set dct = dctInstncsProcLocal(sKey)
        If dct.Exists(c_instance) Then
            IsInstanceLocal = True
            c_class = dct(c_instance)
        End If
        Set dct = Nothing
    End If
    
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function IsLocalProcedure(ByVal i_comp As String, _
                                  ByVal i_proc As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the procedure (i_proc) is a component's (i_comp) procedure.
' ------------------------------------------------------------------------------
    IsLocalProcedure = dctProcs(i_comp).Exists(i_proc)
End Function

Private Function IsSheet(ByVal i_comp_name As String, _
                         ByVal i_wb As Workbook) As Boolean
' ------------------------------------------------------------------------
' Returns TRUE when the Component's name (i_comp_name) is a Worksheet's
' CodeName in the Workbook (i_wb).
' ------------------------------------------------------------------------
    Dim ws As Worksheet
    
    For Each ws In i_wb.Worksheets
        If ws.CodeName = i_comp_name Then
            IsSheet = True
            Exit For
        End If
    Next ws

End Function

Public Function IsSheetDocMod(ByVal i_vbc As VBComponent, _
                              ByVal i_wbk As Workbook, _
                     Optional ByRef i_wsh As Worksheet) As Boolean
' ------------------------------------------------------------------------------
' When the VBComponent (vbc) represents a Worksheet the function returns TRUE
' and the corresponding Worksheet (i_wsh).
' ------------------------------------------------------------------------------
    Dim wsh As Worksheet

    IsSheetDocMod = i_vbc.Type = vbext_ct_Document And i_vbc.name <> i_wbk.CodeName
    If IsSheetDocMod Then
'        Debug.Print "i_vbc.Name: " & i_vbc.name
        For Each wsh In i_wbk.Worksheets
'            Debug.Print "wsh.CodeName: " & wsh.CodeName
            If wsh.CodeName = i_vbc.name Then
                Set i_wsh = wsh
                Exit For
            End If
        Next wsh
    End If

End Function

Private Function IsUniquePublicItem(ByVal i_item As String) As Boolean
    If dctPublicItemsUnique.Exists(i_item) Then
        IsUniquePublicItem = dctPublicItemsUnique(i_item).Count = 1
    Else
'        Debug.Print "The procedure named '" & i_item & "' is a procedure name in more than one VBComponent!"
    End If
    
End Function

Private Sub ItemAs(ByVal c_line As String, _
                     ByRef c_as As String)
' ------------------------------------------------------------------------------------
' Returns from the code line (c_line) the Public " As " (c_as).
' ------------------------------------------------------------------------------------
    Const PROC = "ItemAs"
                       
    On Error GoTo eh
    Select Case True
        Case c_line Like "*) As New *": c_as = Split(c_line, ") As New ")(1)
        Case c_line Like "*) As *":     c_as = Split(Split(c_line, ") As ")(1), " ")(0)
        Case c_line Like "* As New *":  c_as = Split(c_line, " As New ")(1)
        Case c_line Like "* As *":      c_as = Split(Split(c_line, " As ")(1), " ")(0)
    End Select
    If c_as = vbNullString Then Stop
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function KindOfComponent(ByVal en As enKindOfComponent) As String
    Select Case en
        Case enStandardModule:  KindOfComponent = "Standard-Module"
        Case enClassModule:     KindOfComponent = "Class-Module"
        Case enWorkbook:        KindOfComponent = "Workbook"
        Case enWorksheet:       KindOfComponent = "WorkSheet"
        Case enUserForm:        KindOfComponent = "UserForm"
    End Select
End Function

Private Function KindOfItem(ByVal en As enKindOfItem) As String

    Select Case en
        Case enClassInstance:   KindOfItem = "Class-Instance"
        Case enConstant:        KindOfItem = "Constant"
        Case enEnumeration:     KindOfItem = "Enumeration"
        Case enFunction:        KindOfItem = "Function"
        Case enMethod:          KindOfItem = "Method"
        Case enPropertyGet:     KindOfItem = "Property-Get"
        Case enPropertyLet:     KindOfItem = "Property-Let"
        Case enPropertySet:     KindOfItem = "Property-Set"
        Case enSub:             KindOfItem = "Sub-Procedure"
        Case enUserDefinedType: KindOfItem = "User-Defined-Type"
        Case enVariable:        KindOfItem = "Variable"
    End Select
    
End Function

Private Function KoPstring(ByVal k_kop As vbext_ProcKind) As String
    Select Case k_kop
        Case vbext_pk_Get:  KoPstring = "Get"
        Case vbext_pk_Let:  KoPstring = "Let"
        Case vbext_pk_Proc: KoPstring = "Proc"
        Case vbext_pk_Set:  KoPstring = "Set"
    End Select
End Function

Private Function Max(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the maximum value of all values provided (va).
' --------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function MaxKindOfComp() As Long
    Dim en  As enKindOfComponent
    
    For en = enKindOfComponent.enA To enKindOfComponent.enZ
        MaxKindOfComp = Max(MaxKindOfComp, Len(KindOfComponent(en)))
    Next en
    
End Function

Private Function MaxKindOfItem() As Long
    Dim en  As enKindOfItem
    
    For en = enKindOfItem.enA To enKindOfItem.enZ
        MaxKindOfItem = Max(MaxKindOfItem, Len(KindOfItem(en)))
    Next en
    
End Function

Private Function MaxLenItems(ByVal dct As Dictionary) As Long
    Dim v As Variant
    For Each v In dct
        MaxLenItems = Max(MaxLenItems, Len(v))
    Next v
End Function

Private Function OpenUrlEtc(ByVal oue_string As String, _
                            ByVal oue_show_how As Long) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples
' - Open a folder:          OpenUrlEtc("C:\TEMP\",WIN_NORMAL)
' - Call Email app:         OpenUrlEtc("mailto:dash10@hotmail.com",WIN_NORMAL)
' - Open URL:               OpenUrlEtc("http://home.att.net/~dashish", WIN_NORMAL)
' - Handle Unknown extensions (call Open With Dialog):
'                           OpenUrlEtc("C:\TEMP\TestThis",Win_Normal)
' - Start Access instance:  OpenUrlEtc("I:\mdbs\CodeNStuff.mdb", Win_NORMAL)
'
' Copyright:
' This code was originally written by Dev Ashish. It is not to be altered or
' distributed, except as part of an application. You are free to use it in any
' application, provided the copyright notice is left unchanged.
'
' Code Courtesy of: Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, oue_string, vbNullString, vbNullString, oue_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Error: Out of Memory/Resources. Couldn't Execute!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Error: File not found.  Couldn't Execute!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Error: Path not found. Couldn't Execute!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Error:  Bad File Format. Couldn't Execute!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & oue_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:         lRet = -1
    End Select
    
    OpenUrlEtc = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Private Function ProcCollLine(ByVal cll As Collection) As String:       ProcCollLine = cll(2):      End Function

Private Function ProcCollLineFrom(ByVal cll As Collection) As String:   ProcCollLineFrom = cll(3):  End Function

Private Function ProcCollLineTo(ByVal cll As Collection) As String:     ProcCollLineTo = cll(4):    End Function

Private Function ProcCollVBCM(ByVal cll As Collection) As CodeModule:   Set ProcCollVBCM = cll(1):  End Function

Private Sub ProcsCollect()
' ------------------------------------------------------------------------------
' Assembles for each Procedure in non excluded VBComponents (dctComps)
' the first relevant code line (Sub, Function, Property), its line number and
' the Procedures last line number (End xxx) in a Collection and returns a
' Dictionary with these Collections as the item and <compname>.<procname> as the
' key.
' ------------------------------------------------------------------------------
    Const PROC = "ProcsCollect"
    
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
    lLinesTotal = 0
    lProcsTotal = 0
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
            i = .CountOfDeclarationLines + 1
            lStop = 0
            '~~ Collect all Procedures in the CodeModule
            While CodeLineNext(i, vbcm, sLine, lStartsAt) And i <= .CountOfLines
                If CodeLineIsFirstOfProc(.ProcOfLine(i, KoP), sProc, KoP) Then
                    lFrom = i
                    sKey = Split(sProc, ".")(0)
                    Set cllCompProc = New Collection
                    cllCompProc.Add vbcm
                    cllCompProc.Add sLine
                    cllCompProc.Add lFrom
                End If
                If CodeLineIsEndProc(sLine) Then
                    lProcsTotal = lProcsTotal + 1
                    cllCompProc.Add i
                    lLines = (cllCompProc(cllCompProc.Count) - cllCompProc(cllCompProc.Count - 1)) + 1
                    lLinesTotal = lLinesTotal + lLines
                    If Not dctCompProcs.Exists(sKey) Then
                        AddAscByKey dctCompProcs, sKey, cllCompProc
                        Set cllCompProc = Nothing
                    End If
                End If
                i = i + 1
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

Private Sub PublicItem(ByVal c_public As String, _
                       ByVal c_line As String, _
                       ByRef c_item As String)
' ------------------------------------------------------------------------------------
' Returns the name of the Public (c_public) as the public item (c_item).
' ------------------------------------------------------------------------------------
    Const PROC = "PublicItem"
                       
    On Error GoTo eh
    c_item = Trim(Split(c_line, c_public)(1))
    c_item = Trim(Split(c_item, " ")(0))
    c_item = Trim(Split(c_item, "(")(0))
'    If c_line Like "*Public Const *" Then
'        c_item_value = Trim(Split(c_line & " ", " = ")(1))
'        c_item_value = Split(c_item_value, " ")(0)
'    End If
    If c_item = vbNullString Then Stop
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PublicItemAs(ByVal c_public As String, _
                         ByVal c_line As String, _
                         ByRef c_item As String, _
                         ByRef c_as As String)
' ------------------------------------------------------------------------------------
' Returns the name of the Public (c_public) as the public item (c_item).
' ------------------------------------------------------------------------------------
    Const PROC = "PublicItemAs"
                       
    On Error GoTo eh
    PublicItem c_public, c_line, c_item
    If c_item = vbNullString Then Stop
    
    ItemAs c_line, c_as
    If c_as = vbNullString Then Stop
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function PublicItemCollCodeLine(ByVal cll As Collection) As String:             PublicItemCollCodeLine = cll(4):                                                        End Function

Private Function PublicItemCollCodeLineNo(ByVal cll As Collection) As String:           PublicItemCollCodeLineNo = cll(3):                                                      End Function

Private Sub PublicItemCollect(ByVal c_comp As String, _
                              ByVal c_item As String, _
                              ByVal c_line_no As Long, _
                              ByVal c_line As String, _
                              ByVal c_kind_of_comp As enKindOfComponent, _
                              ByVal c_kind_of_item As enKindOfItem)
' ------------------------------------------------------------------------------
' Collects a Public item and additionally collects all VBComponents with a
' same named procedure.
' ------------------------------------------------------------------------------
    Dim cll     As Collection
    Dim sKey    As String
    
    sKey = c_comp & "." & c_item
    lLenPublicItems = Max(lLenPublicItems, Len(sKey))
    If Not dctPublicItems.Exists(sKey) Then
        Set cll = New Collection
        cll.Add c_kind_of_comp
        cll.Add c_kind_of_item
        cll.Add c_line_no
        cll.Add c_line
        If Not dctPublicItems.Exists(sKey) Then
            AddAscByKey dctPublicItems, sKey, cll
            Set cll = Nothing
        End If
    End If
      
    '~~ By the way collect all equally named items
    If Not dctPublicItemsUnique.Exists(c_item) Then
        Set cll = New Collection
        cll.Add c_comp
        AddAscByKey dctPublicItemsUnique, c_item, cll
        Set cll = Nothing
    ElseIf dctPublicItemsUnique.Exists(c_item) Then
        Set cll = dctPublicItemsUnique(c_item)
        cll.Add c_comp
        dctPublicItemsUnique.Remove c_item
        AddAscByKey dctPublicItemsUnique, c_item, cll
        Set cll = Nothing
    End If
    
End Sub

Private Function PublicItemCollInCodeLine(ByVal cll As Collection) As String:               PublicItemCollInCodeLine = cll(9):                                                      End Function

Private Function PublicItemCollInCodeLineNo(ByVal cll As Collection) As String:             PublicItemCollInCodeLineNo = cll(8):                                                    End Function

Private Function PublicItemCollInKindOfComp(ByVal cll As Collection) As enKindOfComponent:  PublicItemCollInKindOfComp = cll(6):                                                    End Function

Private Function PublicItemCollInKindOfCompItem(ByVal cll As Collection) As String:         PublicItemCollInKindOfCompItem = KindOfComponent(cll(6)) & "." & KindOfItem(cll(7)):    End Function

Private Function PublicItemCollInKindOfItem(ByVal cll As Collection) As enKindOfItem:       PublicItemCollInKindOfItem = cll(7):                                                    End Function

Private Function PublicItemCollKindOfComp(ByVal cll As Collection) As enKindOfComponent:    PublicItemCollKindOfComp = cll(1):                                                      End Function

Private Function PublicItemCollKindOfCompItem(ByVal cll As Collection) As String:           PublicItemCollKindOfCompItem = KindOfComponent(cll(1)) & "." & KindOfItem(cll(2)):      End Function

Private Function PublicItemCollKindOfItem(ByVal cll As Collection) As enKindOfItem:         PublicItemCollKindOfItem = cll(2):                                                      End Function

Private Sub PublicItemsUsageCollect()
' ------------------------------------------------------------------------------------
' Loops through all collected Procedures (dctProcs) code lines and scans each
' for any of the collected Public items, removing those found from the collected
' public items. The finally remaing public items are listed as unused.
' ------------------------------------------------------------------------------------
    Const PROC  As String = "PublicItemsUsageCollect"
    
    On Error GoTo eh
    Dim cllUsed             As Collection
    Dim cllProc             As Collection
    Dim dct                 As Dictionary
    Dim dctCompProcs        As Dictionary
    Dim i                   As Long
    Dim lFrom               As Long
    Dim lItems              As Long
    Dim lStartsAt           As Long
    Dim lTo                 As Long
    Dim sClass              As String
    Dim sDelim              As String
    Dim sIgnore             As String
    Dim sLineToParse        As String
    Dim sLineProc           As String
    Dim v                   As Variant
    Dim vbcm                As CodeModule
    Dim vComp               As Variant
    Dim vProc               As Variant
    Dim vPublic             As Variant
    Dim lProcs              As Long
    Dim lProcsTotal         As Long
    Dim lLinesAnalysed      As Long
    Dim lLinesSkipped       As Long
    
    BoP ErrSrc(PROC)
    Initialize
    Collect
    
    Set dct = New Dictionary
    lItems = dctPublicItems.Count
    Set dctUsed = New Dictionary
    For Each vComp In dctProcs
        lProcsTotal = lProcsTotal + dctProcs(vComp).Count
    Next vComp
    
    lLinesExplored = 0
    For Each vComp In dctProcs
        sCompParsing = vComp
        Set dctCompProcs = dctProcs(vComp)
        For Each vProc In dctCompProcs
            lProcs = lProcs + 1
            sProcParsing = vProc
            If sProcParsing = "mCompManTest" Then Stop
            Set cllProc = dctCompProcs(vProc)
            Set vbcm = ProcCollVBCM(cllProc)
            sLineProc = ProcCollLine(cllProc)
            lFrom = ProcCollLineFrom(cllProc)
            lTo = ProcCollLineTo(cllProc)
'            If sCompParsing = "mSyncShapes" And sProcParsing = "AllDone" Then Stop
            For i = lFrom To lTo
                CodeLineNext i, vbcm, sLineProc, lStartsAt
                lLinesExplored = lLinesExplored + 1
                Select Case True
                    Case sLineProc Like "*Const *":         lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc Like "Sub *":            lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc Like "End If*":          lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc Like "Loop*":            lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc Like "Wend*":            lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc Like "* Property Let *": lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc Like "With *End With":   lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc Like "With *"
                        PushInstanceOnWithStack sCompParsing, sProcParsing, sLineProc
                    Case sLineProc Like "End With":         WithStackPop
                    Case sLineProc = "End Sub":             lLinesSkipped = lLinesSkipped + 1:  sIgnore = vbNullString
                    Case sLineProc = "End Function":        lLinesSkipped = lLinesSkipped + 1:  sIgnore = vbNullString
                    Case sLineProc = "End Property":        lLinesSkipped = lLinesSkipped + 1:  sIgnore = vbNullString
                    Case sLineProc = "End Select":          lLinesSkipped = lLinesSkipped + 1
                    Case Else
                        '~~ Analyse the line for any Public Items used
                        lLinesAnalysed = lLinesAnalysed + 1
                        If sLineProc Like "*Function *" _
                        Or sLineProc Like "* Property Get *" _
                        Or sLineProc Like "* Property Set *" Then sIgnore = CodeLineIgnore(sLineProc)
'                        CodeLineStopIfLike sLineProc, "xxxx"
                        
                        '~~ Prepare the code line for being parsed for any used public item
                        CodeLineForBeingParsed sLineProc, sLineToParse, sCompParsing, sProcParsing, sDelim, sIgnore ' prepare line for exploration
                        
                        '~~ Loop through all public items for being checked if used in the sLineToParse
                        For Each vPublic In dctPublicItems
                            sCompPublic = Split(vPublic, ".")(0)
                            sItemPublic = Split(vPublic, ".")(1)
                            '~~ Immediately skip the public item when it is not found in the sLineToParse
                            If InStr(sLineToParse, sItemPublic) = 0 Then GoTo nxp
                            '~~ Skip the public item when the to-be-parsed component is identical with the public item's component
                            If sCompParsing = sCompPublic Then GoTo nxp ' no need to explore the own module
                            
                            If CodeLineRefersToPublicItem(sLineToParse, vPublic, vbcm.Parent, sClass) Then
                                '~~ Move the found public item to the dctUsed dictionary
                                If Not dctUsed.Exists(vPublic) Then
                                    Set cllUsed = dctPublicItems(vPublic)
                                    cllUsed.Add sCompParsing & "." & sProcParsing   ' comp.proc where the public item was found
                                    cllUsed.Add lStartsAt                           ' code line number where the public item was found
                                    cllUsed.Add sLineProc                           ' code line where the public item was found
                                    AddAscByKey dctUsed, vPublic, cllUsed
                                End If
                                '~~ Remove the found public item
                                If dctPublicItems.Exists(vPublic) Then dctPublicItems.Remove (vPublic)
                                GoTo nxp
                            End If
                            
nxp:                    Next vPublic

                        For Each vPublic In dctPublicItems
                            Set cllUsed = dctPublicItems(vPublic)
                            If dctOnActions.Exists(vPublic) Then
                                If Not dctUsed.Exists(vPublic) Then
                                    cllUsed.Add sCompParsing & "." & sProcParsing
                                    cllUsed.Add lStartsAt
                                    cllUsed.Add sLineProc
                                    AddAscByKey dctUsed, vPublic, cllUsed
                                End If
                                If dctPublicItems.Exists(vPublic) Then dctPublicItems.Remove (vPublic)
                            End If
                        Next vPublic
                End Select
                Application.StatusBar = "Analysed Procedures: " & Format(lProcs, "000") & " (of " & Format(lProcsTotal, "000") & ") " & _
                                        "Lines: " & Format(lLinesAnalysed, "00000") & " analysed, " & Format(lLinesSkipped, "00000") & " skipped."
            Next i
        Next vProc
nxt: Next vComp
    
    For Each v In dctUsed
        If dctPublicItems.Exists(v) Then
            dctPublicItems.Remove v
        End If
    Next v
    
    Set dctCodeRefPublicItem = dct
    Set dct = Nothing
    Set cllUsed = Nothing
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PublicVariable(ByVal c_public As String, _
                           ByVal c_line As String, _
                           ByRef c_item As String, _
                           ByRef c_as As String)
' ------------------------------------------------------------------------------------
' Returns the name of the Public (c_public) as the public item (c_item).
' ------------------------------------------------------------------------------------
    Const PROC = "PublicVariable"
                       
    On Error GoTo eh
    PublicItem c_public, c_line, c_item
    If c_item = vbNullString Then Stop
    ItemAs c_line, c_as
    If c_as = vbNullString Then Stop
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub PushInstanceOnWithStack(ByVal r_comp_name As String, _
                                    ByVal r_proc_name As String, _
                                    ByVal r_line As String)
' ------------------------------------------------------------------------------
' Variants are: "With New *" ( * = a known Class Module's Name)
'               "With *"     ( * = a known global or local class instance)
' The item pushed on the WithStack will be a vbNullstring when the instance is
' not known as a Class-Module's instance.
' ------------------------------------------------------------------------------
    Const PROC = "PushInstanceOnWithStack"
    
    On Error GoTo eh
    Dim sInstance   As String
    Dim sClass      As String
    
    If r_line Like "With New *" Then
        sInstance = Trim(Split(r_line, "With New ")(1))
        If IsClassModule(sInstance) Then
            sClass = sInstance
        ElseIf IsClassInstance(r_comp_name, sInstance, sClass, r_proc_name) Then
            GoTo xt
        End If
    
    ElseIf r_line Like "With Me*" Then
        sClass = r_comp_name
        GoTo xt
    
    ElseIf r_line Like "With *" Then
        sInstance = Trim(Split(r_line, "With ")(1))
        If IsClassInstance(r_comp_name, sInstance, sClass, r_proc_name) Then
            GoTo xt
        End If
    End If
    If sClass = "wsConfig" Then Stop
xt: WithStackPush sClass
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub WithStackPop()
    WithStack.Remove WithStack.Count
End Sub

Private Sub WithStackPush(ByVal c_class_mod As String)
                              
    If WithStack Is Nothing Then Set WithStack = New Collection
    WithStack.Add c_class_mod
    
End Sub

Private Function WithStackTop(Optional ByRef c_class_mod As String) As String
' ------------------------------------------------------------------------------------
' Returns the class instance name (c_class_instance) and the class module nmae
' (c_class_mod) currently on top of the WithStack
' ------------------------------------------------------------------------------------
    If WithStack Is Nothing Then Set WithStack = New Collection
    If WithStack.Count > 0 Then
        c_class_mod = WithStack(WithStack.Count)
        WithStackTop = c_class_mod
    End If
End Function

Private Sub WriteToFile(ByVal s As String)
    If sFile = vbNullString Then sFile = FileTemp(tmp_extension:="txt")
    FileText(sFile) = s
End Sub

