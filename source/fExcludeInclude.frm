VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fExcludeInclude 
   Caption         =   "UserForm1"
   ClientHeight    =   5922
   ClientLeft      =   42
   ClientTop       =   392
   ClientWidth     =   7217
   OleObjectBlob   =   "fExcludeInclude.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "fExcludeInclude"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbListUnusedUsed_Click()
    Dim i       As Long
    Dim sDelim  As String
    
    With lbxExcludeInclude
        For i = 0 To .ListCount - 1
            If optInclude.Value = True Then
                If .Selected(i) = False Then
                    Excluded = Excluded & sDelim & .List(i)
                    sDelim = ","
                End If
            ElseIf optExclude.Value = True Then
                If .Selected(i) = True Then
                    Excluded = Excluded & sDelim & .List(i)
                    sDelim = ","
                End If
            End If
        Next i
    End With
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()
    Dim vbc As VBComponent
    Dim dct As Dictionary
    Dim v   As Variant
    
    For Each vbc In wbkServiced.VBProject.VBComponents
        dct.Add vbc.name, vbNullString
    Next vbc
    
    For Each v In dct
        With lbxExcludeInclude
            .AddItem v
        End With
    Next v
    Set dct = Nothing
    
End Sub

Private Sub UserForm_Terminate()
    mItems.Terminated = True
End Sub
