Attribute VB_Name = "mItems"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mItems:
' =======================
' Public services:
' - CollectOnActions
' - CollectPublicItems  Collects of not excluded VBComponents:
'                       - Any VBCpmponent which is a Class Module
'                       - Any item (Sub Function, Property) declared Public
'                       - Any public item's Kind
'                       - All public items with the same name
'                       - All Procedures start and end line
' - CollectPublicUsage
' - ItemAs
' - KindOfComponent
' - KindOfItem
' - KoPstring
' - Variable
'
' W. Rauschenberger, Berlin Apr 2023
' ----------------------------------------------------------------------------
Public Terminated As Boolean

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

Private lLenPublicItems         As Long
Private lLinesTotal             As Long
Private vbpServiced             As VBProject
Private lLinesExplored          As Long
Private lProcsTotal             As Long
Private sCompPublic             As String
Private sProcParsed             As String
Private sCompParsed             As String
Private sItemPublic             As String
Private dctCodeRefPublicItem    As Dictionary
Private dctOnActions            As Dictionary

Public dctKindOfItem            As Dictionary
Public dctPublicItemsUsed       As Dictionary   ' All Public items used
Public dctPublicItemsUnique     As Dictionary ' Collection of all those public items with a unique name
Public dctComps                 As Dictionary
Public dctProcLines             As Dictionary   ' All component's procedures with theit start and end line
Public dctProcs                 As Dictionary
Public dctPublicItems           As Dictionary   ' All Public ... and Friend ... - finally only those not used
Public dctUsed                  As Dictionary
Public sFile                    As String
                          
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

Public Sub CollectOnActions()
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
                    dct.Add sOnAction, vbNullString
                End If
            End If
        Next shp
    Next wsh
    
    Set dctOnActions = dct
    Set dct = Nothing
    
End Sub

Public Sub CollectPublicItem(ByVal c_comp As String, _
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
            dctPublicItems.Add sKey, cll
            Set cll = Nothing
        End If
    End If
      
    '~~ By the way collect all equally named items
    If Not dctPublicItemsUnique.Exists(c_item) Then
        Set cll = New Collection
        cll.Add c_comp
        dctPublicItemsUnique.Add c_item, cll
        Set cll = Nothing
    ElseIf dctPublicItemsUnique.Exists(c_item) Then
        Set cll = dctPublicItemsUnique(c_item)
        cll.Add c_comp
        dctPublicItemsUnique.Remove c_item
        dctPublicItemsUnique.Add c_item, cll
        Set cll = Nothing
    End If

End Sub

Public Sub CollectPublicItems()
' ------------------------------------------------------------------------------
' Collects in dctPublicItems (with comp-name.item-name as key) of not excluded
' VBComponents(dctComps):
' - In dctClassModules any VBCpmponent which is a Class Module
' - In dctPublicItems any item (Sub Function, Property) declared Public
' - In dctKindOfItem any public item's Kind
' - In dctPublicItemsUnique all public items with the same name
' - In dctProcLines all Procedures start and end line
' ------------------------------------------------------------------------------
    Const PROC  As String = "CollectPublicItems"
    
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
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Collect Public declared Sub, Function, Property
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
            sLine = mProcs.CollLine(cllProc)
            i = mProcs.CollLineFrom(cllProc)
            If mLine.DeclaresPublicItem(i, sLine, sItem, sAs, vbcm, enKoComp, enKoItem) Then
                CollectPublicItem sComp, sItem, i, sLine, enKoComp, enKoItem
            End If
        Next vProc
    Next vComp
        
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub CollectPublicUsage()
' ------------------------------------------------------------------------------------
' Loops through all collected Procedures (dctProcs) code lines and scans each
' for any of the collected Public items, removing those found from the collected
' public items. The finally remaing public items are listed as unused.
' ------------------------------------------------------------------------------------
    Const PROC  As String = "CollectPublicUsage"
    
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
    
    mBasic.BoP ErrSrc(PROC)
    
    Set dct = New Dictionary
    lItems = dctPublicItems.Count
    Set dctUsed = New Dictionary
    For Each vComp In dctProcs
        lProcsTotal = lProcsTotal + dctProcs(vComp).Count
    Next vComp
    
    lLinesExplored = 0
    For Each vComp In dctProcs
        sCompParsed = vComp
        Set dctCompProcs = dctProcs(vComp)
        For Each vProc In dctCompProcs
            lProcs = lProcs + 1
            sProcParsed = vProc
            Set cllProc = dctCompProcs(vProc)
            Set vbcm = mProcs.CollVBCM(cllProc)
            sLineProc = mProcs.CollLine(cllProc)
            lFrom = mProcs.CollLineFrom(cllProc)
            lTo = mProcs.CollLineTo(cllProc)
'            If sCompParsed = "mExport" And sProcParsed = "All" Then Stop
            i = lFrom - 1
            sLineProc = mLine.NextLine(vbcm, i, lStartsAt)
            While i <= lTo
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
                        PushInstanceOnWithStack sCompParsed, sProcParsed, sLineProc
                    Case sLineProc Like "End With":         mStack.Pop
                    Case sLineProc = "End Sub":             lLinesSkipped = lLinesSkipped + 1:  sIgnore = vbNullString
                    Case sLineProc = "End Function":        lLinesSkipped = lLinesSkipped + 1:  sIgnore = vbNullString
                    Case sLineProc = "Exit Sub":            lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc = "Exit Function":       lLinesSkipped = lLinesSkipped + 1
                    Case sLineProc = "End Property":        lLinesSkipped = lLinesSkipped + 1:  sIgnore = vbNullString
                    Case sLineProc = "End Select":          lLinesSkipped = lLinesSkipped + 1
                    Case LineExcluded(sLineProc):           lLinesSkipped = lLinesSkipped + 1
                    Case Else
                        '~~ Analyse the line for any Public Items used
                        lLinesAnalysed = lLinesAnalysed + 1
                        If sLineProc Like "*Function *" _
                        Or sLineProc Like "* Property Get *" _
                        Or sLineProc Like "* Property Set *" Then sIgnore = mLine.Ignore(sLineProc)
'                        If mLine.IsLike(sLineProc, "Set fMon = mMsg.MsgInstance(Title)") Then Stop
                        
                        '~~ Prepare the code line for being parsed for any used public item
                        mLine.ForBeingParsed sLineProc, sLineToParse, sCompParsed, sProcParsed, sDelim, sIgnore ' prepare line for exploration
                        
                        '~~ Loop through all public items for being checked if used in the sLineToParse
                        For Each vPublic In dctPublicItems
'                            If vPublic Like "*MsgInstance*" Then Stop
                            sCompPublic = Split(vPublic, ".")(0)
                            sItemPublic = Split(vPublic, ".")(1)
                            '~~ Immediately skip the public item when it is not found in the sLineToParse
                            If InStr(sLineToParse, sItemPublic) <> 0 Then
                                '~~ Skip the public item when the to-be-parsed component is identical with the public item's component
                                If sCompParsed <> sCompPublic Then ' no need to explore the own module
                                    If mLine.RefersToPublicItem(sLineToParse, vPublic, vbcm.Parent, sClass) Then
                                        '~~ Move the found public item to the dctUsed dictionary
                                        If Not dctUsed.Exists(vPublic) Then
                                            Set cllUsed = dctPublicItems(vPublic)
                                            cllUsed.Add sCompParsed & "." & sProcParsed ' comp.proc where the public item was found
                                            cllUsed.Add lStartsAt                       ' code line number where the public item was found
                                            cllUsed.Add sLineProc                       ' code line where the public item was found
                                            dctUsed.Add vPublic, cllUsed
                                            Application.StatusBar = Progress(lProcs, lProcsTotal, lLinesAnalysed, lLinesSkipped)
                                        End If
                                        '~~ Remove the found public item
                                        If dctPublicItems.Exists(vPublic) Then dctPublicItems.Remove (vPublic)
                                    End If
                                End If ' sCompParsed <> sCompPublic
                            End If ' public item found in line
                        Next vPublic

                        '~~ Collect all Public items in OnActions
                        For Each vPublic In dctPublicItems
                            sCompPublic = Split(vPublic, ".")(0)
                            sItemPublic = Split(vPublic, ".")(1)
'                            If vPublic Like "*_Click*" Then Stop
                            Set cllUsed = dctPublicItems(vPublic)
                            If dctOnActions.Exists(sItemPublic) Then
                                If Not dctUsed.Exists(vPublic) Then
                                    cllUsed.Add sCompParsed & "." & sProcParsed
                                    cllUsed.Add lStartsAt
                                    cllUsed.Add sLineProc
                                    dctUsed.Add vPublic, cllUsed
                                End If
                                If dctPublicItems.Exists(vPublic) Then dctPublicItems.Remove (vPublic)
                            End If
                        Next vPublic
                End Select
                Application.UseSystemSeparators = True
                sLineProc = mLine.NextLine(vbcm, i, lStartsAt)
            Wend
        Next vProc
nxt: Next vComp
    Application.StatusBar = Progress(lProcs, lProcsTotal, lLinesAnalysed, lLinesSkipped)
    Debug.Print Progress(lProcs, lProcsTotal, lLinesAnalysed, lLinesSkipped)
        
    Set dctCodeRefPublicItem = dct
    Set dct = Nothing
    Set cllUsed = Nothing
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function CompCollKind(ByVal cll As Collection) As enKindOfComponent:    CompCollKind = cll(2):  End Function

Public Function CompCollVBC(ByVal cll As Collection) As VBComponent:   Set CompCollVBC = cll(1):   End Function

Public Sub DeclaredAs(ByVal c_line As String, _
                      ByRef c_as As String)
' ------------------------------------------------------------------------------------
' Returns from the code line (c_line) the Public " As " (c_as).
' ------------------------------------------------------------------------------------
    Const PROC = "DeclaredAs"
                       
    On Error GoTo eh
    Select Case True
        Case c_line Like "*) As New *": c_as = Split(c_line, ") As New ")(1)
        Case c_line Like "*) As *":     c_as = Split(Split(c_line, ") As ")(1), " ")(0)
        Case c_line Like "* As New *":  c_as = Split(c_line, " As New ")(1)
        Case c_line Like "* As *":      c_as = Split(Split(c_line, " As ")(1), " ")(0)
    End Select
    If c_as = vbNullString Then Stop
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mItems" & "." & e_proc
End Function

Public Function IsUniqueItem(ByVal i_item As String) As Boolean
    If dctPublicItemsUnique.Exists(i_item) Then
        IsUniqueItem = dctPublicItemsUnique(i_item).Count = 1
    End If
    
End Function

Public Sub Item(ByVal c_public As String, _
                ByVal c_line As String, _
                ByRef c_item As String)
' ------------------------------------------------------------------------------------
' Returns the name of the Public (c_public) as the public item (c_item).
' ------------------------------------------------------------------------------------
    Const PROC = "Item"
                       
    On Error GoTo eh
    c_item = Trim(Split(c_line, c_public)(1))
    c_item = Trim(Split(c_item, " ")(0))
    c_item = Trim(Split(c_item, "(")(0))
    If c_item = vbNullString Then Stop
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ItemAs(ByVal c_public As String, _
                  ByVal c_line As String, _
                  ByRef c_item As String, _
                  ByRef c_as As String)
' ------------------------------------------------------------------------------------
' Returns the name of the Public (c_public) as the public item (c_item).
' ------------------------------------------------------------------------------------
    Const PROC = "ItemAs"
                       
    On Error GoTo eh
    Item c_public, c_line, c_item
    If c_item = vbNullString Then Stop
    
    DeclaredAs c_line, c_as
    If c_as = vbNullString Then Stop
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function KindOfComponent(ByVal en As enKindOfComponent) As String
    Select Case en
        Case enStandardModule:  KindOfComponent = "Standard-Module"
        Case enClassModule:     KindOfComponent = "Class-Module"
        Case enWorkbook:        KindOfComponent = "Workbook"
        Case enWorksheet:       KindOfComponent = "WorkSheet"
        Case enUserForm:        KindOfComponent = "UserForm"
    End Select
End Function

Public Function KindOfItem(ByVal en As enKindOfItem) As String

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

Public Function KoPstring(ByVal k_kop As vbext_ProcKind) As String
    Select Case k_kop
        Case vbext_pk_Get:  KoPstring = "Get"
        Case vbext_pk_Let:  KoPstring = "Let"
        Case vbext_pk_Proc: KoPstring = "Proc"
        Case vbext_pk_Set:  KoPstring = "Set"
    End Select
End Function

Private Function Progress(ByVal p_procs As Long, _
                          ByVal p_procs_total As Long, _
                          ByVal p_lines_analysed As Long, _
                          ByVal p_lines_skipped As Long) As String
    Progress = "Items (used/unused): " & Format(dctUsed.Count, "#,##0") & "/" & _
               Format(dctPublicItems.Count, "#,##0") & _
               " Analysed: " & Format(p_procs, "#,000") & _
               " (of " & Format(p_procs_total, "#,000") & _
               ") Procedures, " & Format(p_lines_analysed, "#0,000") & _
               " Lines analysed, " & Format(p_lines_skipped, "#0,000") & _
               " Lines skipped."
End Function

Private Function PublicItemCollCodeLine(ByVal cll As Collection) As String:                 PublicItemCollCodeLine = cll(4):                                                        End Function

Private Function PublicItemCollCodeLineNo(ByVal cll As Collection) As String:               PublicItemCollCodeLineNo = cll(3):                                                      End Function

Private Function PublicItemCollInCodeLine(ByVal cll As Collection) As String:               PublicItemCollInCodeLine = cll(9):                                                      End Function

Private Function PublicItemCollInCodeLineNo(ByVal cll As Collection) As String:             PublicItemCollInCodeLineNo = cll(8):                                                    End Function

Private Function PublicItemCollInKindOfComp(ByVal cll As Collection) As enKindOfComponent:  PublicItemCollInKindOfComp = cll(6):                                                    End Function

Private Function PublicItemCollInKindOfCompItem(ByVal cll As Collection) As String:         PublicItemCollInKindOfCompItem = KindOfComponent(cll(6)) & "." & KindOfItem(cll(7)):    End Function

Private Function PublicItemCollInKindOfItem(ByVal cll As Collection) As enKindOfItem:       PublicItemCollInKindOfItem = cll(7):                                                    End Function

Private Function PublicItemCollKindOfComp(ByVal cll As Collection) As enKindOfComponent:    PublicItemCollKindOfComp = cll(1):                                                      End Function

Public Function PublicItemCollKindOfCompItem(ByVal cll As Collection) As String:            PublicItemCollKindOfCompItem = KindOfComponent(cll(1)) & "." & KindOfItem(cll(2)):      End Function

Private Function PublicItemCollKindOfItem(ByVal cll As Collection) As enKindOfItem:         PublicItemCollKindOfItem = cll(2):                                                      End Function

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
        If mClass.IsClassModule(sInstance) Then
            sClass = sInstance
        ElseIf mClass.IsInstance(r_comp_name, sInstance, sClass, r_proc_name) Then
            GoTo xt
        End If
    
    ElseIf r_line Like "With Me*" Then
        sClass = r_comp_name
        GoTo xt
    
    ElseIf r_line Like "With *" Then
        sInstance = Trim(Split(r_line, "With ")(1))
        sInstance = Split(sInstance, "(")(0)
        If mClass.IsInstance(r_comp_name, sInstance, sClass, r_proc_name) Then
            GoTo xt
        End If
    End If
    If sClass = "wsConfig" Then Stop
xt: mStack.Push sClass
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Variable(ByVal c_public As String, _
                    ByVal c_line As String, _
                    ByRef c_item As String, _
                    ByRef c_as As String)
' ------------------------------------------------------------------------------------
' Returns the name of the code line (c_line) the public declared item (c_item) and its
' As declaration (c_as).
' ------------------------------------------------------------------------------------
    Const PROC = "Variable"
                       
    On Error GoTo eh
    Item c_public, c_line, c_item
    If c_item = vbNullString Then Stop
    DeclaredAs c_line, c_as
    If c_as = vbNullString Then Stop
    
xt: Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

