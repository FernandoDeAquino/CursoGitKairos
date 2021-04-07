VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_PesMembros 
   Caption         =   "Form2"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   4530
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Leva"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3915
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancela"
      Height          =   375
      Left            =   5910
      TabIndex        =   5
      Top             =   3900
      Width           =   1500
   End
   Begin VB.TextBox TXT_Matricula 
      Height          =   315
      Left            =   330
      TabIndex        =   2
      Top             =   270
      Width           =   1215
   End
   Begin VB.TextBox Txt_Nome 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   270
      Width           =   5745
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3075
      Left            =   270
      TabIndex        =   0
      Top             =   660
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5424
      _Version        =   393216
      Rows            =   1
      Cols            =   3
   End
   Begin VB.Label Label_Medida 
      AutoSize        =   -1  'True
      Caption         =   "Label_Medida"
      Height          =   315
      Left            =   840
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Matricula: "
      Height          =   225
      Left            =   330
      TabIndex        =   4
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Nome: "
      Height          =   225
      Left            =   1650
      TabIndex        =   3
      Top             =   60
      Width           =   675
   End
End
Attribute VB_Name = "FRM_PesMembros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rd As New ADODB.Recordset
' Mensagem para alteração no git
Private Sub Command1_Click()

Grid.Col = 0
oMembros.Matricula = Grid.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Dim Rd As New ADODB.Recordset
    Me.Grid.Row = 0
           Me.Grid.Col = 0
            Me.Grid = "Matriculas"
            Me.Grid.Col = 1
            Me.Grid = "Nomes"
            Me.Grid.Col = 2
            Me.Grid = "Status"


Set Rd = oMembros.pesquisa(Caminho)

   If (Not Rd.EOF) Then
        Call MontaGrid(Rd)
           
   End If

Grid.Row = IIf(Grid.Rows > 1, 1, 0)

End Sub
Private Sub MontaGrid(Rd As ADODB.Recordset)
While Not Rd.EOF
            Grid.Rows = Grid.Rows + 1
            Grid.Row = Grid.Rows - 1
            Me.Grid.Col = 0
            Me.Grid = IIf(Not IsNull(Rd.Fields("ID_Matr").Value), Rd.Fields("ID_Matr").Value, "")
            Me.Grid.Col = 1
            Me.Grid = IIf(Not IsNull(Rd.Fields("C_Nome").Value), Rd.Fields("C_Nome").Value, "")
            Me.Grid.Col = 2
            Me.Grid = IIf(Not IsNull(Rd.Fields("C_Situacao").Value), Rd.Fields("C_Situacao").Value, "")
            Rd.MoveNext
            
        Wend
        Call AjustaTamalho(Grid, 2240)
End Sub
Private Sub Grid_DblClick()
Call Command1_Click

End Sub
Public Sub AjustaTamalho(Grid As MSFlexGrid, MaxTamColuna As Integer)
Dim TotLinhas As Integer
Dim TotColunas As Integer
Dim i As Integer
Dim I2 As Integer
Dim TamWidth As Double
Dim TamWidthGrid As Double
TamWidthGrid = 0
TotLinhas = Grid.Rows - 1
TotColunas = Grid.Cols - 1
'Grid.Redraw = False
For i = 0 To TotColunas
    TamWidth = 0
    Grid.Col = i
        For I2 = 0 To TotLinhas
            Grid.Row = I2
            Label_Medida.Caption = Grid.Text
            If TamWidth < Label_Medida.Width Then
                TamWidth = Label_Medida.Width
            End If
        Next
        TamWidthGrid = TamWidthGrid + TamWidth
        Grid.ColWidth(i) = TamWidth + 200
Next


Grid.ColWidth(0) = 0
End Sub

Private Sub Txt_Nome_KeyUp(KeyCode As Integer, Shift As Integer)
Grid.Rows = 1
Set Rd = oMembros.pesquisa(Caminho, , UCase(Txt_Nome.Text))
If (Not Rd.EOF) Then
        Call MontaGrid(Rd)
           
   End If

Grid.Row = IIf(Grid.Rows > 1, 1, 0)
End Sub
