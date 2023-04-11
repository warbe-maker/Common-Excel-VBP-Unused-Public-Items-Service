Attribute VB_Name = "mItems"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mItems
' =======================
' Public services:
' - AddAscByKey Adds items to a Dictionary instantly oredered by key.
' - Collect     Collects (in Dictionary with comp-name.item-name as key):
'               1. All those VBComponen not explicitely excluded
'               1.1 All Class Modules
'               2. All items declared Public
'               3. All public item's Kind
'               4. All public items indicating unique True or False
'               5. All VBComponent's Procedures with their start and end line
' -
'
' W. Rauschenberger
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
Public dctUnused                As Dictionary
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

Public Sub CollectPublicItems()
' ------------------------------------------------------------------------------
' Collects (in Dictionary with comp-name.item-name as key):
' 1. All those VBComponenKoItemts not explicitely excluded (dctComponenKoItemts)
'    1.1 All Class Modules (dctClassModules)
' 2. All items declared Public (dctPublicItems)
' 3. All public item's Kind (dctKindOfItem)
' 4. All public items indicating unique True or False (dctPublicItemsUnique)
' 5. All VBComponenKoItemt's Procedures with their start and enKoItemd line (dctProcLines)
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
            sLine = mProc.CollLine(cllProc)
            i = mProc.CollLineFrom(cllProc)
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

Public Function CompCollKind(ByVal cll As Collection) As enKindOfComponent:    CompCollKind = cll(2):  End Function

Public Function CompCollVBC(ByVal cll As Collection) As VBComponent:   Set CompCollVBC = cll(1):   End Function

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mItems" & "." & e_proc
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

Public Function IsUniqueItem(ByVal i_item As String) As Boolean
    If dctPublicItemsUnique.Exists(i_item) Then
        IsUniqueItem = dctPublicItemsUnique(i_item).Count = 1
    Else
'        Debug.Print "The procedure named '" & i_item & "' is a procedure name in more than one VBComponent!"
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

Public Function KoPstring(ByVal k_kop As vbext_ProcKind) As String
    Select Case k_kop
        Case vbext_pk_Get:  KoPstring = "Get"
        Case vbext_pk_Let:  KoPstring = "Let"
        Case vbext_pk_Proc: KoPstring = "Proc"
        Case vbext_pk_Set:  KoPstring = "Set"
    End Select
End Function

Public Function Max(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the maximum value of all values provided (va).
' --------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

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
            Set vbcm = mProc.CollVBCM(cllProc)
            sLineProc = mProc.CollLine(cllProc)
            lFrom = mProc.CollLineFrom(cllProc)
            lTo = mProc.CollLineTo(cllProc)
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
                    Case sLineProc Like "End With":         mWithStack.Pop
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
'                        mLine.StopIfLike sLineProc, ".DueModificationWarning = True"
                        
                        '~~ Prepare the code line for being parsed for any used public item
                        mLine.ForBeingParsed sLineProc, sLineToParse, sCompParsed, sProcParsed, sDelim, sIgnore ' prepare line for exploration
                        
                        '~~ Loop through all public items for being checked if used in the sLineToParse
                        For Each vPublic In dctPublicItems
                            sCompPublic = Split(vPublic, ".")(0)
                            sItemPublic = Split(vPublic, ".")(1)
                            '~~ Immediately skip the public item when it is not found in the sLineToParse
                            If InStr(sLineToParse, sItemPublic) = 0 Then GoTo nxp
                            '~~ Skip the public item when the to-be-parsed component is identical with the public item's component
                            If sCompParsed = sCompPublic Then GoTo nxp ' no need to explore the own module
                            
                            If mLine.RefersToPublicItem(sLineToParse, vPublic, vbcm.Parent, sClass) Then
                                '~~ Move the found public item to the dctUsed dictionary
                                If Not dctUsed.Exists(vPublic) Then
                                    Set cllUsed = dctPublicItems(vPublic)
                                    cllUsed.Add sCompParsed & "." & sProcParsed   ' comp.proc where the public item was found
                                    cllUsed.Add lStartsAt                           ' code line number where the public item was found
                                    cllUsed.Add sLineProc                           ' code line where the public item was found
                                    dctUsed.Add vPublic, cllUsed
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
                Application.StatusBar = "Analysed: " & Format(lProcs, "#,000") & " (of " & Format(lProcsTotal, "#,000") & ") Procedures, " & _
                                        Format(lLinesAnalysed, "#0,000") & " Lines analysed, " & Format(lLinesSkipped, "#0,000") & " Lines skipped."
                sLineProc = mLine.NextLine(vbcm, i, lStartsAt)
            Wend
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
    
xt: mBasic.EoP ErrSrc(PROC)
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
        If mClass.IsInstance(r_comp_name, sInstance, sClass, r_proc_name) Then
            GoTo xt
        End If
    End If
    If sClass = "wsConfig" Then Stop
xt: mWithStack.Push sClass
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

