Attribute VB_Name = "mUnused"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mUnused
' =======================
' Public services:
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
Public sExcludedCodeLines   As String
Public vExcludedCodeLines   As Variant

Public Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
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
Public Const ERROR_SUCCESS = 32&
Public Const ERROR_NO_ASSOC = 31&
Public Const ERROR_OUT_OF_MEM = 0&
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_BAD_FORMAT = 11&

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

Private Sub DisplayResult()
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

    lMaxKindOfComp = MaxKindOfComp
    lMaxKindOfItem = MaxKindOfItem
    lMaxLenItems = MaxLenItems(dctPublicItems)
    
    lMaxCompProcKind = lMaxKindOfComp + lMaxKindOfItem + 3
    
    s = "The following " & dctPublicItems.Count & " Public declared items are  u n u s e d ! *)" & vbCrLf:                                              WriteToFile s
    s = Align("Kind of Component.Item", lMaxCompProcKind, AlignCentered) & " " & Align("Public item (component.item)", lMaxLenItems, AlignCentered):    WriteToFile s
    s = String(lMaxCompProcKind, "-") & " " & String(lMaxLenItems, "-"):                                                                                WriteToFile s
    
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
    s = "   I.e. an unused Public item may still be used within its own Component.":                            WriteToFile s
    WriteToFile "   In case the Public item should rather be turned into Private!"
    WriteToFile String(Len(s), "=")
    WriteToFile vbLf
    s = "The following " & dctUsed.Count & " Public declared items had been found in at least one code line:":  WriteToFile s
    WriteToFile String(Len(s), "-")
    
    For Each vPublic In dctUsed
        Set cll = dctUsed(vPublic)
        lMaxPublic = mPublic.Max(lMaxPublic, Len(vPublic))
        lMaxUsing = mPublic.Max(lMaxUsing, Len(cll(5)))
    Next vPublic
    
    WriteToFile Align("Public item", lMaxPublic, AlignLeft) & " " & Align("Used in (VBComponent.Procedure) by example", lMaxUsing + 2, AlignLeft) & "In code line"
    WriteToFile String(lMaxPublic, "-") & " " & String(lMaxUsing + 2, "-") & " " & String(80, "-")
    
    For Each vPublic In dctUsed
        Set cll = dctUsed(vPublic)
        sComp = Split(vPublic, ".")(0)
        sProc = Split(vPublic, ".")(1)
        WriteToFile Align(vPublic, lMaxPublic + 1, AlignLeft, , ".") & Align(cll(5), lMaxUsing + 1, AlignLeft, , ".") & ": " & cll(7)
    Next vPublic

    OpenUrlEtc sFile, WIN_NORMAL

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
    If err_source = vbNullString Then err_source = Err.Source
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
        MaxKindOfComp = mPublic.Max(MaxKindOfComp, Len(KindOfComponent(en)))
    Next en
    
End Function

Private Function MaxKindOfItem() As Long
    Dim en  As enKindOfItem
    
    For en = enKindOfItem.enA To enKindOfItem.enZ
        MaxKindOfItem = mPublic.Max(MaxKindOfItem, Len(KindOfItem(en)))
    Next en
    
End Function

Private Function MaxLenItems(ByVal dct As Dictionary) As Long
    Dim v As Variant
    For Each v In dct
        MaxLenItems = mPublic.Max(MaxLenItems, Len(v))
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

Private Sub ProvisionOfExcludedCodelines(Optional ByVal p_excluded As String = vbNullString)
    Dim sSplit As String
    
    If p_excluded <> vbNullString Then
        If InStr(p_excluded, vbCrLf) <> 0 Then
            sSplit = vbCrLf
        ElseIf InStr(p_excluded, vbLf) <> 0 Then
            sSplit = vbLf
        Else
            sSplit = vbCr
        End If
        vExcludedCodeLines = Split(p_excluded, sSplit)
    End If

End Sub

Private Sub ProvisionOfExcludedComponents(Optional ByVal p_excluded As String = "n o n e  s p e c i f i e d")
    
    If p_excluded = "n o n e  s p e c i f i e d" Then
        fExcludeInclude.Show ' assembles in Excluded the ignored VBComponents
        If Terminated Then GoTo xt
        Set fExcludeInclude = Nothing
    Else
        Excluded = p_excluded
    End If

xt: Exit Sub

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
    BoP ErrSrc(PROC)
    
    ProvisionOfTheServicedWorkbook u_wbk:   If wbkServiced Is Nothing Then GoTo xt
    
    ProvisionOfExcludedComponents u_excluded_components
    
    ProvisionOfExcludedCodelines u_excluded_code_lines
    
    mPublic.PublicItemsUsageCollect
    
    DisplayResult
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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

Private Sub WriteToFile(ByVal s As String)
    If sFile = vbNullString Then sFile = FileTemp(tmp_extension:="txt")
    FileText(sFile) = s
End Sub

