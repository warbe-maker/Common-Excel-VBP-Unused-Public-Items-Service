Attribute VB_Name = "mComps"
Option Explicit
' ------------------------------------------------------------------------------------
' Standard Module mComps:
' =======================
' Public services:
' - Collect Provides a Dictionary with all VBComponents not excluded.
' ------------------------------------------------------------------------------------

Private dctExcluded As Dictionary

Private Function ErrSrc(ByVal e_proc As String) As String
    ErrSrc = "mComps" & "." & e_proc
End Function

Public Sub Collect(ByVal c_excluded As String)
' ------------------------------------------------------------------------------------
' Provides a Dictionary (dctComps) with all components not excluded
' XrefVBProject with the VBComponent's Name as the key and the VBComponent as item.
' ------------------------------------------------------------------------------------
    Const PROC  As String = "Collect"
    
    On Error GoTo eh
    Dim vbc     As VBComponent
    Dim enKind  As enKindOfComponent
    Dim sComp   As String
    Dim cll     As Collection
    
    mBasic.BoP ErrSrc(PROC)
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
                    dctComps.Add .name, cll
                    Set cll = Nothing
                End If
                If .Type = vbext_ct_ClassModule Then
                    mClass.IsClassModule(sComp, vbc) = True
                End If
            End If
        End With
    Next vbc
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

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


