Attribute VB_Name = "mPublic"
Option Explicit
' ----------------------------------------------------------------------------
' Standard Module mUnusedPublic
' -----------------------------
' The 'Unused' service analyses the code of a Workbook for any unused
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
Private dctExcluded             As Dictionary
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
    mProc.Collect           ' Collect all procedures in not exluded VBComponenKoItemts
    mClass.CollectInstncsVBPGlobal   ' Collect all class instance which are VB-Project global
    mClass.CollectInstncsCompGlobal  ' Collect all class instances which are ComponenKoItemt global
    mClass.CollectInstncsProcLocal   ' Collect all class instances in Procedures
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
            sLine = mProc.CollLine(cllProc)
            i = mProc.CollLineFrom(cllProc)
            If mLine.DeclaresPublicItem(i, sLine, sItem, sAs, vbcm, enKoComp, enKoItem) Then
                ItemCollect sComp, sItem, i, sLine, enKoComp, enKoItem
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

Public Function CompCollKind(ByVal cll As Collection) As enKindOfComponent:    CompCollKind = cll(2):  End Function

Public Function CompCollVBC(ByVal cll As Collection) As VBComponent:   Set CompCollVBC = cll(1):   End Function

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
                    mClass.IsModule(sComp, vbc) = True
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

Private Function IsExcluded(ByVal i_comp_name As String) As Boolean
    IsExcluded = dctExcluded.Exists(i_comp_name)
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ItemCollect(ByVal c_comp As String, _
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

eh: Select Case ErrMsg(ErrSrc(PROC))
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

Public Sub PublicItemsUsageCollect()
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
                                    cllUsed.Add sCompParsed & "." & sProcParsed
                                    cllUsed.Add lStartsAt
                                    cllUsed.Add sLineProc
                                    AddAscByKey dctUsed, vPublic, cllUsed
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
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Variable(ByVal c_public As String, _
                    ByVal c_line As String, _
                    ByRef c_item As String, _
                    ByRef c_as As String)
' ------------------------------------------------------------------------------------
' Returns the name of the Public (c_public) as the public item (c_item).
' ------------------------------------------------------------------------------------
    Const PROC = "Variable"
                       
    On Error GoTo eh
    Item c_public, c_line, c_item
    If c_item = vbNullString Then Stop
    DeclaredAs c_line, c_as
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
        If mClass.IsModule(sInstance) Then
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub


