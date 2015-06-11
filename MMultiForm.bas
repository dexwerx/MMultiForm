Attribute VB_Name = "MMultiForm"
' Copyright © 2014 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' MMultiForm.bas
'
' Implements multiple duplicate forms in MS Access
'   - adapted from http://allenbrowne.com/ser-35.html
'   - Instatiate new form using 'OpenFormMulti New Form_frmEditRecord, "ID=" & ID'
'   - Call 'CloseFormMulti Me' from Form_Close() event, or you will get weird bugs...
'   - Use generic 'DoCmd.Close' to close the active multi form
'   - Do not use explicit 'DoCmd.Close acForm, Me.Name, acSaveNo'
'
Option Compare Database
Option Explicit

Private MultiForms As New Collection

Function OpenFormMulti(NewForm As Form, Optional Filter As String = "", Optional Cascade As Boolean = True)
    MultiForms.Add NewForm, CStr(NewForm.hWnd)

    With NewForm
        .AllowAdditions = Not (Len(Filter) > 0)
        .DataEntry = .AllowAdditions
        If Not .AllowAdditions Then
            .Filter = Filter
            .FilterOn = True
        End If
        .Caption = .Caption
        .Visible = True
        .SetFocus
    End With
    
    If Cascade And MultiForms.Count > 1 Then
        Dim LastForm As Form
        Set LastForm = MultiForms(MultiForms.Count - 1)
        If LastForm Is Nothing Then Else NewForm.Move LastForm.WindowLeft + 400, LastForm.WindowTop + 400
    End If
End Function

Function CloseFormMulti(CloseForm As Form) As Boolean
    Dim CurForm As Object
    
    If CloseForm Is Nothing Then Exit Function
    
    For Each CurForm In MultiForms
        If CurForm.hWnd = CloseForm.hWnd Then
            MultiForms.Remove CStr(CloseForm.hWnd)
            Set CloseForm = Nothing
            Set CurForm = Nothing
            Exit Function
        End If
    Next
End Function

Function CloseAllFormMulti(Optional CloseFormType As Object = Nothing)
    Dim CurForm As Object
    
    For Each CurForm In MultiForms
        If CloseFormType Is Nothing Then
            MultiForms.Remove CStr(CurForm.hWnd)
            Set CurForm = Nothing
        ElseIf TypeName(CurForm) = TypeName(CloseFormType) Then
            MultiForms.Remove CStr(CurForm.hWnd)
            Set CurForm = Nothing
        End If
    Next
    
    Set CloseFormType = Nothing
End Function
