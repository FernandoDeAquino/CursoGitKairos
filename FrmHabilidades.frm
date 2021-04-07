VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_Habilidades 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Bindings        =   "FrmHabilidades.frx":0000
      Height          =   2055
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Gravar"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sair"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Excluir"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Novo"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "C_Habilidade"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   1
      Top             =   360
      Width           =   4425
   End
   Begin VB.TextBox TXT_ID 
      DataField       =   "ID_Habilidade"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Left            =   240
      MaxLength       =   10
      TabIndex        =   0
      Top             =   360
      Width           =   945
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Anderson\Kairos\Sistema\Secretaria\kairos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Habilidades"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Habilidade: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "FRM_Habilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Data1.Recordset.AddNew
   ' Data1.Refresh
    Me.Grid.TextMatrix(0, 1) = "Código"
    Me.Grid.TextMatrix(0, 2) = "Descrição"
End Sub

Private Sub Command2_Click()
    Data1.Recordset.Delete
    Data1.Refresh
End Sub

Private Sub Command4_Click()
    Data1.Recordset.Update
    Data1.Refresh
    Me.Grid.TextMatrix(0, 1) = "Código"
    Me.Grid.TextMatrix(0, 2) = "Descrição"
End Sub

Private Sub Form_Load()
  Data1.DatabaseName = App.Path & "\Kairos.MDB" 'path/nome do arquivo
  Data1.RecordSource = "Habilidades"         'fonte dos dados(tabela clientes)
  Data1.RecordsetType = 0                 'tipo de recordset(0-tabela)
  Data1.Refresh                           'implementa as alterações
'  Set Grid.DataSource = Data1.Recordset


'Me.Grid.Cols = 2
'Me.Grid.Rows = 1
Data1.Refresh
Me.Grid.TextMatrix(0, 1) = "Código"
Me.Grid.TextMatrix(0, 2) = "Descrição"
'Data1.Refresh
' Call EncheGrid

End Sub

Private Sub MSFlexGrid1_Click()

'Public Sub EncheGrid()
'Dim i As Integer
'Dim Rs As Recordset
'Set Rs = Data1.Recordset
'Rs.MoveFirst
'While Not Data1.Recordset.EOF
'i = i + 1
'Me.Grid.TextMatrix(1, 0) = Data1.Recordset.Fields(0)
'Me.Grid.TextMatrix(1, 1) = Data1.Recordset.Fields(1)
'Data1.Recordset.MoveNext
'
'Wend
End Sub

Private Sub Grid_Click()
'Dim Sql As String
'    Data1.Recordset.FindFirst ("ID_Habilidade=" & Me.TXT_ID.Text)
'    Data1.Recordset.Delete
'    Data1.Refresh

End Sub
