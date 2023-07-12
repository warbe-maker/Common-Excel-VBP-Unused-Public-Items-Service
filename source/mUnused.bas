Attribute VB_Name = "mUnused"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mUnused:
' ========================
'
' Public services:
' ----------------
' - Unused  - Displays a Workbook(file) selection dialog when no serviced
'             Workbook argument is provided(u_wbk) is provided
'           - Displays a VBComponent selection dialog when no excluded
'             VBComponents argument is provided (a vbNullString declared none
'             excluded)
'           - Allows the specification of excluded Code-Lines
'           - Collects all Public items in the selected VBComponents and
'             displays those unused and used (the used ones only with the
'             compoent.procedure found and the code line)
'
' W. Rauschenberger, Berlin Apr 2023
'
' See also: https://github.com/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service
' ----------------------------------------------------------------------------

Public wbkServiced          As Workbook
Public Excluded             As String
Public vExcludedCodeLines   As Variant

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
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Sub Initialize()

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

Private Sub DisplayResult()
    Const PROC = "DisplayResult"
    
    On Error GoTo eh
    Dim cll                 As Collection
    Dim dctUnused           As Dictionary
    Dim lMaxCompProcKind    As Long
    Dim lMaxKindOfComp      As Long
    Dim lMaxKindOfItem      As Long
    Dim lMaxLenItems        As Long
    Dim lMaxLineNo          As Long
    Dim lMaxPublic          As Long
    Dim lMaxUsing           As Long
    Dim s                   As String
    Dim sComp               As String
    Dim sProc               As String
    Dim vPublic             As Variant
    Dim Log                 As clsLog
    
    lMaxKindOfComp = MaxKindOfComp
    lMaxKindOfItem = MaxKindOfItem
    lMaxLenItems = MaxLenItems(dctPublicItems)
    lMaxCompProcKind = lMaxKindOfComp + lMaxKindOfItem + 3
    
    With New clsLog
        .FileName = "PublicUnused.txt"
        .NewFile ' New result file with each execution
        .Title "The following " & dctPublicItems.Count & " Public declared items are  u n u s e d ! *)"
        .MaxItemLengths lMaxCompProcKind, lMaxLenItems
        .Headers "|Kind of Component.Item |Public item (component.item) |"
        .ColsDelimiter = " "
        .AlignmentItems "|L|L|"
        
        Set dctUnused = KeySort(dctPublicItems)
        For Each vPublic In dctUnused
            Set cll = dctPublicItems(vPublic)
            sComp = Split(vPublic, ".")(0)
            sProc = Split(vPublic, ".")(1)
            s = PublicItemCollKindOfCompItem(cll)
            .Entry s, vPublic
        Next vPublic
        .Entry " "
        .Entry "*) Public items are not analysed in their own component.  "
        .Entry "   I.e. an unused Public item may still be used within its own Component."
        .Entry "   In case the Public item should rather be turned into Private!"
        .Entry " "
    
        KeySort dctUsed
        For Each vPublic In dctUsed
            Set cll = dctUsed(vPublic)
            lMaxPublic = mBasic.Max(lMaxPublic, Len(vPublic))
            lMaxUsing = mBasic.Max(lMaxUsing, Len(cll(5)))
            lMaxLineNo = mBasic.Max(lMaxLineNo, Len(cll(6)))
        Next vPublic
    
        .Title "The following " & dctUsed.Count & " Public declared items had been found in at least one code line:"
        .MaxItemLengths lMaxPublic, lMaxUsing, lMaxLineNo, 60
        .Headers "|Public Item |Used in VBComponent.Procedure |Line| Code line "
        .Headers "|            |(by example)                  | No |           "
        .AlignmentItems "|L.|L.:|R| L|"
        .ColsDelimiter = " "
        
        For Each vPublic In dctUsed
            Set cll = dctUsed(vPublic)
            sComp = Split(vPublic, ".")(0)
            sProc = Split(vPublic, ".")(1)
            .Entry vPublic, cll(5), cll(6), cll(7)
        Next vPublic
        .Dsply
    End With

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function KeySort(ByRef s_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (s_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim Temp    As Variant
    Dim Txt     As String
    Dim i       As Long
    Dim j       As Long
    
    If s_dct.Count = 0 Then GoTo xt
    With s_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add key:=vKey, Item:=s_dct.Item(vKey)
    Next i
    
    Set s_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    
xt: Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mUnused" & "." & e_proc
End Function

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

Private Function MaxKindOfComp() As Long
    Dim en  As enKindOfComponent
    
    For en = enKindOfComponent.enA To enKindOfComponent.enZ
        MaxKindOfComp = mBasic.Max(MaxKindOfComp, Len(KindOfComponent(en)))
    Next en
    
End Function

Private Function MaxKindOfItem() As Long
    Dim en  As enKindOfItem
    
    For en = enKindOfItem.enA To enKindOfItem.enZ
        MaxKindOfItem = mBasic.Max(MaxKindOfItem, Len(KindOfItem(en)))
    Next en
    
End Function

Private Function MaxLenItems(ByVal dct As Dictionary) As Long
    Dim v As Variant
    For Each v In dct
        MaxLenItems = mBasic.Max(MaxLenItems, Len(v))
    Next v
End Function

Private Sub ProvisionOfExcludedCodelines(Optional ByVal p_excluded As String = vbNullString)
    Dim sSplit As String
    
    If p_excluded <> vbNullString Then
        If InStr(p_excluded, vbCrLf) <> 0 Then
            sSplit = vbCrLf
        ElseIf InStr(p_excluded, vbLf) <> 0 _
        Then
            sSplit = vbLf
        Else
            sSplit = vbCr
        End If
        vExcludedCodeLines = Split(p_excluded, sSplit)
    End If

End Sub

Private Sub ProvisionOfExcludedComponents(Optional ByVal p_excluded As String = "n o n e  s p e c i f i e d")
    Const PROC = "ProvisionOfExcludedComponents"
    
    On Error GoTo eh
    If p_excluded = "n o n e  s p e c i f i e d" Then
        fExcludeInclude.Show ' assembles in Excluded the ignored VBComponents
        If Terminated Then GoTo xt
        Set fExcludeInclude = Nothing
    Else
        Excluded = p_excluded
    End If

xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ProvisionOfTheServicedWorkbook(ByVal p_wbk As Workbook)
    Dim sWbk As String
    
    If p_wbk Is Nothing Then
        sWbk = WbkSelect
        If sWbk = vbNullString Then GoTo xt
        GetOpen sWbk, wbkServiced
    Else
        Set wbkServiced = p_wbk
    End If

xt: Exit Sub

End Sub

Public Sub Unused(Optional ByVal u_wbk As Workbook = Nothing, _
                  Optional ByVal u_excluded_components As String = "n o n e  s p e c i f i e d", _
                  Optional ByVal u_excluded_code_lines As String = vbNullString)
' ------------------------------------------------------------------------------
' - When no serviced Workbook (u_wbk) is provided, a file selection dialog is
'   displayed for the selection of a Workbook - which is opened when not already
'   open. When no Workbook is elected the procedure terminates,
' - When no excluded components are specified, i.e. not even indication by a
'   vbNullString that noe are excluded, a VBComponent selection dialog is
'   displayeded for a decision which ones to include or exclude,
' - All Public items in the selected VBComponents are collected and those not
'   used in any code line are displayed finally.
'
' W. Rauschenberger, Berlin Apr 2023
'
' See also: https://github.com/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service
' ------------------------------------------------------------------------------
    Const PROC  As String = "Unused"
    
    On Error GoTo eh
    mBasic.BoP ErrSrc(PROC)
    
    ProvisionOfTheServicedWorkbook u_wbk:   If wbkServiced Is Nothing Then GoTo xt
    ProvisionOfExcludedComponents u_excluded_components
    ProvisionOfExcludedCodelines u_excluded_code_lines
    
    Initialize
    mComps.Collect Excluded
    mProcs.Collect                   ' Collect all procedures in not exluded VBComponenKoItemts
    mClass.CollectInstncsScopeProject  ' Collect all class instance which are VB-Project global
    mClass.CollectInstncsScopeComp ' Collect all class instances which are ComponenKoItemt global
    mClass.CollectInstncsScopeProc  ' Collect all class instances in Procedures
    CollectOnActions
    mItems.CollectPublicItems
    mItems.CollectPublicUsage
    
    DisplayResult
    
xt: mBasic.EoP ErrSrc(PROC)
    mTrc.Dsply
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function WbkSelect() As String
    Dim fDialog As FileDialog

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

'Private Sub WriteToFile(ByVal s As String)
'    If sFile = vbNullString Then sFile = FileTemp(tmp_extension:="txt")
'    FileText(sFile) = s
'End Sub

