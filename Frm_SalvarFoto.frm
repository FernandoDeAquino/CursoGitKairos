VERSION 5.00
Begin VB.Form Frm_SalvarComo 
   Caption         =   "Kairos - Salvar Foto"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   ControlBox      =   0   'False
   Icon            =   "Frm_SalvarFoto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   0
      Pattern         =   "*.jpg"
      TabIndex        =   4
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Gravar"
      Height          =   240
      Left            =   600
      TabIndex        =   3
      Top             =   3330
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   240
      Left            =   1650
      TabIndex        =   2
      Top             =   3330
      Width           =   1005
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   48
      TabIndex        =   1
      Top             =   48
      Width           =   3330
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   48
      TabIndex        =   0
      Top             =   432
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2490
   End
End
Attribute VB_Name = "Frm_SalvarComo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Msg As String
Dim Style As String
Dim Title As String
Dim Response As Integer
Dim Arquivo As String
Dim Caminho As String
Dim AUX As String
Arquivo = File1.FileName
Caminho = File1.Path



If Len(Arquivo) <= 0 Then Exit Sub
 If Dir(App.Path & "\img", vbDirectory) = "" Then MkDir (App.Path & "\img")

    If Dir(App.Path & "\img\" & oMembros.Matricula & ".jpg", vbNormal) <> "" Then
        Msg = "Aten��o!" & Chr(13) & "O arquivo abaixo j� existe: " & Chr(13) & Msg & Chr(13) & "Deseja substitu�-lo ?" ' Define message.
        Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
        Title = "Sobrescrevendo arquivos na base de dados."   ' Define title.

        Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Response = 7 Then Exit Sub
       
        Kill (App.Path & "\img\" & oMembros.Matricula & ".jpg")
    End If
    
    Call FileCopy(Caminho & "\" & Arquivo, App.Path & "\img\" & oMembros.Matricula & ".jpg")
   
    Unload Me
End Sub

Private Sub Command2_Click()
    Flag = ""
    Unload Me
End Sub

Private Sub Dir1_Change()
Me.File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo TrErro
    Dir1.Path = Drive1.Drive
Exit Sub
TrErro:
MsgBox Err.Number & ":  " & Err.Description
End Sub

Private Sub File1_Click()
'Dim x As PictureBox
AUX = File1.Path & "\" & File1.FileName
Set Me.Image1 = LoadPicture(AUX)
'Picture1.AutoRedraw = True
'Me.Picture1.Image = Aux
'Me.Picture1.Refresh
End Sub

Private Sub Form_Load()
    Drive1.Drive = App.Path
    Dir1.Path = App.Path
    
    Call Dir1_Change
End Sub
