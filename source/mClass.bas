Attribute VB_Name = "mClass"
Option Explicit
' ------------------------------------------------------------------------------------
' Standard Module mClass: Collection and checks for Class Modules and Class Instances.
' =======================
' Public services:
' - CollectInstncsCompGlobal    Collection of Component global Class instances
' - CollectInstncsVBPGlobal     Collection of Project global Class instances
' - CollectInstncsProcLocal     Collection of Procedure local Class instances
' - IsInstance                  Check if a string is known as a Class instance
' - IsClassModule Get           Check if a Name is the Name of a Class Module
'                 Let           Register a Name is the Name of a Class Module
' ------------------------------------------------------------------------------------

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
    
    mBasic.BoP ErrSrc(PROC)
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
            While mLine.DeclaresGlobalClassInstance(sLine, sItem, sAs)
                lStopLoop = lStopLoop + 1
                If lStopLoop > 50 Then Stop
                If mClass.IsClassModule(sAs) Then
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

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    dct.Add wbkServiced.CodeName, wbkServiced.CodeName
    For Each wsh In wbkServiced.Worksheets
        dct.Add wsh.CodeName, wsh.CodeName
    Next wsh
    dctInstncsVBPrjctGlobal.Add vbNullString, dct
    Set dct = Nothing
    
    '~~ Collect any Public declared variables, constants, and class instances
    For Each v In dctComps
        Set dct = New Dictionary
        sComp = v
        Set cll = dctComps(v)
        Set vbc = CompCollVBC(cll)
'        If vbc.name = "mCompManClient" Then Stop
        enKoComp = CompCollKind(cll)
        Set vbcm = vbc.CodeModule
        i = 0
        sLine = mLine.NextLine(vbcm, i, lNextSubLine)
        While i <= vbcm.CountOfDeclarationLines And sLine <> vbNullString
            If sLine Like "Public *" Then
                If DeclaresPublicItem(i, sLine, sItem, sAs, vbcm, enKoComp, enKoItem) Then
                    If mClass.IsClassModule(sAs, vbc) Then
                        dct.Add sItem, sAs
                    Else
                        CollectPublicItem sComp, sItem, i, sLine, enKoComp, enKoItem
                    End If
                    
                End If
            End If
            sLine = mLine.NextLine(vbcm, i, lNextSubLine)
        Wend
        dctInstncsVBPrjctGlobal.Add sComp, dct
        Set dct = Nothing
    Next v
            
End Sub

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mClass" & "." & e_proc
End Function

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
    
    mBasic.BoP ErrSrc(PROC)
    Set dctInstncsProcLocal = New Dictionary
    For Each vComp In dctProcs
        Set dctInstComp = New Dictionary
        sComp = vComp
        Set dctCompProcs = dctProcs(sComp)
        For Each vProc In dctCompProcs
            sProc = vProc
            Set cllProc = dctCompProcs(vProc)
            Set vbcm = mProcs.CollVBCM(cllProc)
            lFrom = mProcs.CollLineFrom(cllProc)
            lTo = mProcs.CollLineTo(cllProc)
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
                        If mClass.IsClassModule(sAs) Then
                            dctInstProc.Add sItem, sAs
                        End If
                    Case sLine Like "* As *"
                        sItem = Split(Trim(Split(sLine, " As ")(0)), " ")(UBound(Split(Trim(Split(sLine, " As ")(0)), " ")))
                        sAs = Split(Trim(Split(sLine, " As ")(1)), " ")(0)
                        If mClass.IsClassModule(sAs) Then
                            If Not dctInstProc.Exists(sItem) Then
                                dctInstProc.Add sItem, sAs
                            End If
                        End If
                End Select
                sLine = mLine.NextLine(vbcm, i, lStartsAt)
            Wend
        
            If Not dctInstComp.Exists(sProc) And dctInstProc.Count <> 0 Then
                dctInstComp.Add sProc, dctInstProc
                Set dctInstProc = Nothing
            End If
        Next vProc
    
        If Not dctInstncsProcLocal.Exists(sComp) And dctInstComp.Count <> 0 Then
            dctInstncsProcLocal.Add sComp, dctInstComp
            Set dctInstComp = Nothing
        End If
    Next vComp
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function IsInstance(ByVal i_comp_name As String, _
                           ByVal i_instance_name As String, _
                           ByRef i_class_name As String, _
                  Optional ByVal i_proc_name As String = vbNullString) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE and the corresponding class' name (i_class_name) when the
' instance (i_instance_name) is a known Class instance .
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

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Property Get IsClassModule(Optional ByVal i_name As String, _
                             Optional ByVal i_vbc As VBComponent) As Boolean
    Set i_vbc = i_vbc
    If Not dctClassModules Is Nothing Then
        IsClassModule = dctClassModules.Exists(i_name)
    End If
End Property

Public Property Let IsClassModule(Optional ByVal i_name As String, _
                                  Optional ByVal i_vbc As VBComponent, _
                                           ByVal i_is As Boolean)
                                            
    If dctClassModules Is Nothing Then Set dctClassModules = New Dictionary
    If i_is Then
        If Not dctClassModules.Exists(i_name) Then
            dctClassModules.Add i_name, i_vbc
        End If
    End If
    
End Property



