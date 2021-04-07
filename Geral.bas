Attribute VB_Name = "Geral"
Option Explicit
Public Caminho  As String
Public oMembros As ClMembros


Sub Main()

If IsNull(Command$) Or Command$ = "" Then
    Caminho = App.Path & "\kairos.mdb"
Else
    Caminho = Trim$(Command$) & "\kairos.mdb"
End If
'Senha.Show vbModal
FRM_Membros.Show
End Sub
Public Sub FormataData(Txt As TextBox, KeyAscii As Integer, Optional MesAno = False)
Select Case KeyAscii
  Case Is = Asc("0"), Asc("1"), Asc("2"), Asc("3"), Asc("4"), Asc("5"), Asc("6"), Asc("7"), Asc("8"), Asc("9"), Asc("10"), 8
    If KeyAscii = 8 Then Exit Sub
  Case Else
  KeyAscii = 0
End Select

  If MesAno Then
If (Len(Txt.Text) = 2) Then
        Txt.Text = Txt.Text & "/"
        Txt.SelStart = Len(Txt.Text)
    End If
Else
    If (Len(Txt.Text) = 2) Or (Len(Txt.Text) = 5) Then
        Txt.Text = Txt.Text & "/"
        Txt.SelStart = Len(Txt.Text)
    End If
End If
End Sub
Public Sub PreenCheCombo(Combo As ComboBox, Texto As String)
Dim i As Byte
For i = 0 To Combo.ListCount - 1
        Combo.ListIndex = i
        If Combo.Text = Texto Then Exit For
        If i = Combo.ListCount - 1 Then
            Combo.ListIndex = 0
'            Combo.Text = ""
        End If
    Next
End Sub
Public Function LimpaCampo(NomeForm As Form)
Dim Controle   As Control
Dim Aux As String

Aux = ""
For Each Controle In NomeForm.Controls
    If TypeOf Controle Is TextBox Then Controle.Text = ""
    If TypeOf Controle Is ComboBox Then
        If Controle.ListIndex > -1 Then Controle.ListIndex = 0
    End If
    If TypeOf Controle Is MSFlexGrid Then Controle.Clear
Next Controle

End Function
