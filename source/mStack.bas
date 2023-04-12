Attribute VB_Name = "mStack"
Option Explicit

Private WithStack               As Collection

Public Sub Pop()
    WithStack.Remove WithStack.Count
End Sub

Public Sub Push(ByVal c_class_mod As String)
                              
    If WithStack Is Nothing Then Set WithStack = New Collection
    WithStack.Add c_class_mod
    
End Sub

Public Function Top(Optional ByRef c_class_mod As String) As String
' ------------------------------------------------------------------------------------
' Returns the class instance name (c_class_instance) and the class module nmae
' (c_class_mod) currently on top of the WithStack
' ------------------------------------------------------------------------------------
    If WithStack Is Nothing Then Set WithStack = New Collection
    If WithStack.Count > 0 Then
        c_class_mod = WithStack(WithStack.Count)
        Top = c_class_mod
    End If
    
End Function


