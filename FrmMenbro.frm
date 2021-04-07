VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FRM_Membros 
   Caption         =   "Cadastro de Membros"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FRAM_BUTTONS 
      Height          =   720
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9540
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   450
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   794
         ButtonWidth     =   714
         ButtonHeight    =   688
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   9
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Novo"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Primeiro Registro"
               Object.Tag             =   ""
               ImageIndex      =   6
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Registro Anterior"
               Object.Tag             =   ""
               ImageIndex      =   7
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Proximo Registro"
               Object.Tag             =   ""
               ImageIndex      =   8
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Ultimo Registro"
               Object.Tag             =   ""
               ImageIndex      =   9
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Alterar"
               Object.Tag             =   ""
               ImageIndex      =   2
            EndProperty
            BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Imprimir"
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Excluir"
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
            BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.ToolTipText     =   "Localicar"
               Object.Tag             =   ""
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Cmm_Sair 
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Cmm_Cancela 
         Caption         =   "Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grava"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   9315
      Begin VB.ComboBox Com_Funcao 
         Height          =   315
         ItemData        =   "FrmMenbro.frx":0000
         Left            =   6600
         List            =   "FrmMenbro.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Txt_Nome 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   6225
      End
      Begin VB.TextBox TXT_Matricula 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Função Eclesiástica: "
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
         Left            =   6600
         TabIndex        =   11
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Nome: "
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
         Left            =   210
         TabIndex        =   5
         Top             =   150
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Matricula: "
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
         TabIndex        =   1
         Top             =   150
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados Pessoais"
      TabPicture(0)   =   "FrmMenbro.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dados Eclesiástico"
      TabPicture(1)   =   "FrmMenbro.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   5145
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   8805
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "FrmMenbro.frx":003C
            Left            =   5280
            List            =   "FrmMenbro.frx":003E
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   4080
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   4800
            TabIndex        =   67
            Top             =   2130
            Width           =   2220
         End
         Begin VB.TextBox TXT_Natural 
            Height          =   315
            Left            =   2550
            TabIndex        =   65
            Top             =   4080
            Width           =   2220
         End
         Begin VB.TextBox TXT_EstadoCivil 
            Height          =   315
            Left            =   150
            TabIndex        =   63
            Top             =   4080
            Width           =   2220
         End
         Begin VB.TextBox TXT_Mae 
            Height          =   315
            Left            =   4440
            TabIndex        =   61
            Top             =   3480
            Width           =   4260
         End
         Begin VB.TextBox TXT_Pai 
            Height          =   315
            Left            =   150
            TabIndex        =   59
            Top             =   3480
            Width           =   4260
         End
         Begin VB.TextBox Txt_End 
            Height          =   315
            Left            =   150
            TabIndex        =   6
            Top             =   360
            Width           =   6555
         End
         Begin VB.TextBox TXT_Numero 
            Height          =   315
            Left            =   150
            TabIndex        =   8
            Top             =   960
            Width           =   1515
         End
         Begin VB.TextBox Txt_Complemento 
            Height          =   315
            Left            =   1830
            TabIndex        =   10
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox TXT_Bairro 
            Height          =   315
            Left            =   4320
            TabIndex        =   12
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox TXT_CEP 
            Height          =   315
            Left            =   150
            TabIndex        =   14
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox TXT_DTNascimento 
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   24
            Top             =   2880
            Width           =   1845
         End
         Begin VB.TextBox TXT_Municipio 
            Height          =   315
            Left            =   2790
            TabIndex        =   16
            Top             =   1530
            Width           =   2415
         End
         Begin VB.TextBox TXT_Estado 
            Height          =   315
            Left            =   5400
            TabIndex        =   18
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox TXT_Tel 
            Height          =   315
            Left            =   150
            TabIndex        =   20
            Top             =   2160
            Width           =   2220
         End
         Begin VB.TextBox TXT_Celular 
            Height          =   315
            Left            =   2490
            TabIndex        =   22
            Top             =   2160
            Width           =   2220
         End
         Begin VB.TextBox TXT_CPF 
            Height          =   315
            Left            =   2430
            TabIndex        =   26
            Top             =   2850
            Width           =   1395
         End
         Begin VB.TextBox TXT_Ident 
            Height          =   315
            Left            =   3960
            TabIndex        =   27
            Top             =   2850
            Width           =   1395
         End
         Begin VB.TextBox TXT_ORGAOEMISOR 
            Height          =   315
            Left            =   5520
            TabIndex        =   28
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox TXT_DTEmisao 
            Height          =   315
            Left            =   7020
            MaxLength       =   10
            TabIndex        =   29
            Top             =   2880
            Width           =   1395
         End
         Begin VB.TextBox TXT_Profissao 
            Height          =   315
            Left            =   150
            TabIndex        =   30
            Top             =   4680
            Width           =   4740
         End
         Begin VB.CommandButton CommFoto 
            Caption         =   "Inserir"
            Height          =   255
            Left            =   7320
            TabIndex        =   43
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label30 
            Caption         =   "Igreja: "
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
            Left            =   5280
            TabIndex        =   72
            Top             =   3840
            Width           =   2355
         End
         Begin VB.Label Label28 
            Caption         =   "Telefone Trabalho: "
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
            Left            =   4830
            TabIndex        =   68
            Top             =   1920
            Width           =   1965
         End
         Begin VB.Label Label27 
            Caption         =   "Natural:"
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
            Left            =   2520
            TabIndex        =   66
            Top             =   3840
            Width           =   1125
         End
         Begin VB.Label Label26 
            Caption         =   "Estado Civil:"
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
            Left            =   120
            TabIndex        =   64
            Top             =   3840
            Width           =   1125
         End
         Begin VB.Label Label25 
            Caption         =   "Mãe: "
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
            Left            =   4440
            TabIndex        =   62
            Top             =   3240
            Width           =   1125
         End
         Begin VB.Label Label24 
            Caption         =   "Pai:"
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
            Left            =   120
            TabIndex        =   60
            Top             =   3240
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "Endereço: "
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
            Left            =   150
            TabIndex        =   58
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Numero:"
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
            Left            =   120
            TabIndex        =   57
            Top             =   750
            Width           =   885
         End
         Begin VB.Label Label5 
            Caption         =   "Complemento:"
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
            Left            =   1800
            TabIndex        =   56
            Top             =   750
            Width           =   1305
         End
         Begin VB.Label Label6 
            Caption         =   "Bairro:"
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
            Left            =   4320
            TabIndex        =   55
            Top             =   750
            Width           =   1125
         End
         Begin VB.Label Label7 
            Caption         =   "CEP:"
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
            Left            =   180
            TabIndex        =   54
            Top             =   1350
            Width           =   1125
         End
         Begin VB.Label Label8 
            Caption         =   "Data de Nascimento:"
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
            Left            =   150
            TabIndex        =   53
            Top             =   2670
            Width           =   1995
         End
         Begin VB.Label Label10 
            Caption         =   "Municipio:"
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
            Left            =   2790
            TabIndex        =   52
            Top             =   1320
            Width           =   1125
         End
         Begin VB.Label Label11 
            Caption         =   "Estado:"
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
            Left            =   5400
            TabIndex        =   51
            Top             =   1350
            Width           =   1125
         End
         Begin VB.Label Label12 
            Caption         =   "Telefone: "
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
            Left            =   120
            TabIndex        =   50
            Top             =   1920
            Width           =   1125
         End
         Begin VB.Label Label13 
            Caption         =   "Celular: "
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
            Left            =   2520
            TabIndex        =   49
            Top             =   1950
            Width           =   1125
         End
         Begin VB.Label Label14 
            Caption         =   "CPF: "
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
            Left            =   2430
            TabIndex        =   48
            Top             =   2640
            Width           =   1125
         End
         Begin VB.Label Label15 
            Caption         =   "Identidade: "
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
            Left            =   3960
            TabIndex        =   47
            Top             =   2640
            Width           =   1125
         End
         Begin VB.Label Label16 
            Caption         =   "Orgão Emisor: "
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
            Left            =   5550
            TabIndex        =   46
            Top             =   2670
            Width           =   1395
         End
         Begin VB.Label Label17 
            Caption         =   "Data de Emisão: "
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
            Left            =   7050
            TabIndex        =   45
            Top             =   2670
            Width           =   1485
         End
         Begin VB.Label Label22 
            Caption         =   "Profissão: "
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
            Left            =   120
            TabIndex        =   44
            Top             =   4440
            Width           =   1125
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1770
            Left            =   6960
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5220
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   8715
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmMenbro.frx":0040
            Left            =   120
            List            =   "FrmMenbro.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox TXT_DTBatismo 
            Height          =   315
            Left            =   2370
            MaxLength       =   10
            TabIndex        =   36
            Top             =   450
            Width           =   1545
         End
         Begin VB.TextBox TXT_DTConversao 
            Height          =   315
            Left            =   120
            MaxLength       =   10
            TabIndex        =   33
            Top             =   450
            Width           =   1545
         End
         Begin VB.TextBox TXT_OBS 
            Height          =   1965
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   3000
            Width           =   4665
         End
         Begin VB.ComboBox CMB_Situacao 
            Height          =   315
            ItemData        =   "FrmMenbro.frx":0044
            Left            =   120
            List            =   "FrmMenbro.frx":0046
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   2040
            Width           =   2535
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Incluir"
            Height          =   255
            Left            =   5400
            TabIndex        =   25
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Excluir"
            Height          =   255
            Left            =   6960
            TabIndex        =   23
            Top             =   3720
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   3165
            Left            =   5280
            TabIndex        =   31
            Top             =   480
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   5583
            _Version        =   393216
            Rows            =   1
            Cols            =   3
         End
         Begin VB.Label Label29 
            Caption         =   "Ministerio:"
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
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   2355
         End
         Begin VB.Label Label_Medida 
            AutoSize        =   -1  'True
            Caption         =   "Label_Medida"
            Height          =   195
            Left            =   5760
            TabIndex        =   40
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Label Label9 
            Caption         =   "Data de Batismo:"
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
            Left            =   2400
            TabIndex        =   38
            Top             =   240
            Width           =   1845
         End
         Begin VB.Label Label21 
            Caption         =   "Data de Conversão:"
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
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "OBS: "
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
            Left            =   120
            TabIndex        =   35
            Top             =   2760
            Width           =   1125
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "Habilidades:"
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
            Left            =   5280
            TabIndex        =   34
            Top             =   240
            Width           =   3165
         End
         Begin VB.Label Label23 
            Caption         =   "Situação: "
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
            Left            =   120
            TabIndex        =   32
            Top             =   1800
            Width           =   2355
         End
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":0048
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":062A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":0C0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":11EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":17D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":1DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":22F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":2806
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenbro.frx":2D18
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRM_Membros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FlagStatus As String

Private Sub Cmm_Cancela_Click()
Dim Rd As New ADODB.Recordset
Set oMembros = New ClMembros
If FlagStatus = "N" Then
 Set Rd = oMembros.Primeiro(Caminho)
Else
Set Rd = oMembros.Localizar(Caminho, TXT_Matricula.Text)
End If
If Not Rd.EOF Then Call RecebeCampo(Rd)
Call EnibeFrame
Set oMembros = Nothing
Set Rd = Nothing
End Sub

Private Sub Cmm_Sair_Click()
Unload Me
End Sub


Public Sub TXT_Maiuscula()
Dim Controle   As Control
Dim Aux As String

Aux = ""
For Each Controle In Me.Controls
    If TypeOf Controle Is TextBox Then Controle.Text = UCase(Controle.Text)
Next Controle

End Sub

Private Sub Command1_Click()
    Call TXT_Maiuscula
    Set oMembros = New ClMembros
    Call oMembros.Grava(Caminho, Me.TXT_Matricula.Text, Me.Txt_Nome.Text, Me.Txt_End.Text, Me.Txt_Complemento.Text, Me.TXT_Numero.Text, Me.TXT_Bairro.Text, Me.TXT_Municipio.Text, Me.TXT_Estado.Text, Me.TXT_CEP.Text, Me.TXT_Celular, Me.TXT_Tel.Text, Me.TXT_DTNascimento, Me.TXT_DTBatismo, Me.TXT_CPF.Text, Me.TXT_Ident.Text, Me.TXT_ORGAOEMISOR.Text, Me.TXT_DTEmisao.Text, Me.TXT_OBS.Text, TXT_DTConversao.Text, TXT_Profissao.Text, Me.CMB_Situacao.Text, Me.Com_Funcao.Text)
    Set oMembros = Nothing
    Call EnibeFrame
End Sub

Private Sub Command2_Click()
Dim Rd As New ADODB.Recordset

Set oMembros = New ClMembros

FRM_PesMembros.Show vbModal
If oMembros.Matricula <> 0 Then
    Set Rd = oMembros.pesquisa(Caminho, oMembros.Matricula)

    If (Not Rd.EOF) Then
        'Preenche campos
        Call RecebeCampo(Rd)
    End If
End If
Set oMembros = Nothing

End Sub
Sub RecebeCampo(Rd As ADODB.Recordset)
Dim Aux As String

    Me.TXT_Matricula = IIf(Not IsNull(Rd.Fields("ID_Matr").Value), UCase(Rd.Fields("ID_Matr").Value), "")

Aux = App.Path & "\img\" & Me.TXT_Matricula & ".jpg"
    If Dir(Aux, vbNormal) <> "" Then
        Set Me.Image1 = LoadPicture(Aux)
    Else
        Set Me.Image1 = LoadPicture()
    End If

    Me.TXT_Estado.Text = IIf(Not IsNull(Rd.Fields("C_Estado").Value), UCase(Rd.Fields("C_Estado").Value), "")
    Me.Txt_Nome.Text = IIf(Not IsNull(Rd.Fields("c_Nome").Value), UCase(Rd.Fields("c_Nome").Value), "")
    Me.TXT_CPF.Text = IIf(Not IsNull(Rd.Fields("c_CPF").Value), UCase(Rd.Fields("c_CPF").Value), "")
    Me.Txt_End.Text = IIf(Not IsNull(Rd.Fields("c_end").Value), UCase(Rd.Fields("c_end").Value), "")
    Me.Txt_Complemento.Text = IIf(Not IsNull(Rd.Fields("c_comp").Value), UCase(Rd.Fields("c_comp").Value), "")
'    Che_Desativado.Value = IIf(IIf(Not IsNull(Rd.Fields("F_Desativado").Value), UCase(Rd.Fields("F_Desativado").Value), False) = True, 1, 0)
    Me.TXT_Numero.Text = IIf(Not IsNull(Rd.Fields("c_Numero").Value), UCase(Rd.Fields("c_Numero").Value), "")
    Me.TXT_Municipio.Text = IIf(Not IsNull(Rd.Fields("C_municipio").Value), UCase(Rd.Fields("C_Municipio").Value), "")
    Me.TXT_Bairro.Text = IIf(Not IsNull(Rd.Fields("c_bairro").Value), UCase(Rd.Fields("c_bairro").Value), "")
    Me.TXT_CEP.Text = IIf(Not IsNull(Rd.Fields("c_CEP").Value), UCase(Rd.Fields("c_CEP").Value), "")
    Me.TXT_DTNascimento.Text = IIf(Not IsNull(Rd.Fields("DT_Nascimento").Value), UCase(Rd.Fields("DT_Nascimento").Value), "")
    Me.TXT_DTEmisao.Text = IIf(Not IsNull(Rd.Fields("DT_Emissao").Value), UCase(Rd.Fields("DT_Emissao").Value), "")
    Me.TXT_DTConversao.Text = IIf(Not IsNull(Rd.Fields("DT_conversao").Value), UCase(Rd.Fields("DT_conversao").Value), "")
    Me.TXT_DTBatismo.Text = IIf(Not IsNull(Rd.Fields("DT_Batismo").Value), UCase(Rd.Fields("DT_Batismo").Value), "")
    Me.TXT_OBS.Text = IIf(Not IsNull(Rd.Fields("m_obs").Value), UCase(Rd.Fields("m_obs").Value), "")
   Me.TXT_Ident.Text = IIf(Not IsNull(Rd.Fields("C_INT").Value), UCase(Rd.Fields("C_INT").Value), "")
   Me.TXT_ORGAOEMISOR.Text = IIf(Not IsNull(Rd.Fields("c_orgao").Value), UCase(Rd.Fields("c_orgao").Value), "")
   'Com_Funcao.Text = IIf(Not IsNull(Rd.Fields("C_FuncaoEcle").Value), UCase(Rd.Fields("C_FuncaoEcle").Value), "")
   Me.TXT_Profissao = IIf(Not IsNull(Rd.Fields("C_profissao").Value), UCase(Rd.Fields("C_profissao").Value), "")
   Me.TXT_Celular = IIf(Not IsNull(Rd.Fields("C_celular").Value), UCase(Rd.Fields("C_celular").Value), "")
   Me.TXT_Tel = IIf(Not IsNull(Rd.Fields("C_Tel").Value), UCase(Rd.Fields("C_Tel").Value), "")
   Call PreenCheCombo(Com_Funcao, IIf(Not IsNull(Rd.Fields("C_FuncaoEcle").Value), UCase(Rd.Fields("C_FuncaoEcle").Value), ""))
   Call PreenCheCombo(Me.CMB_Situacao, IIf(Not IsNull(Rd.Fields("C_situacao").Value), UCase(Rd.Fields("C_situacao").Value), ""))
End Sub


Private Sub Command3_Click()
   ' Dim Rd As New ADODB.Recordset
   ' Set oMensalidade = New ClMensalidade
   ' Call oMensalidade.Grava(Caminho, TXT_Matricula.Text, Me.TXT_PagRef.Text, TXT_Valor.Text, Me.TXT_Pagamento.Text)
   ' Set oMensalidade = Nothing
   ' Call CaregaGrid
End Sub

Private Sub CommFoto_Click()
Set oMembros = New ClMembros
oMembros.Matricula = Me.TXT_Matricula
Frm_SalvarComo.Show vbModal

Dim Aux As String
Aux = App.Path & "\img\" & Me.TXT_Matricula & ".jpg"
    If Dir(Aux, vbNormal) <> "" Then
        Set Me.Image1 = LoadPicture(Aux)
    Else
        Set Me.Image1 = LoadPicture()
    End If


Set oMembros = Nothing
End Sub

Private Sub Form_Load()
Dim Rd As New ADODB.Recordset
 Call EnibeFrame
 
 Me.CMB_Situacao.Clear
 Me.CMB_Situacao.AddItem ""
 Me.CMB_Situacao.AddItem UCase("Ativo")
 Me.CMB_Situacao.AddItem UCase("Afastado")
 Me.CMB_Situacao.AddItem UCase("Excluido")
 
 Me.Com_Funcao.Clear
 Me.Com_Funcao.AddItem ""
 Me.Com_Funcao.AddItem "BISPO"
 Me.Com_Funcao.AddItem "CAND.COOPERADOR"
 Me.Com_Funcao.AddItem "COOPERADOR"
 Me.Com_Funcao.AddItem "DIÁCONO"
 Me.Com_Funcao.AddItem "OBREIRO"
 Me.Com_Funcao.AddItem "MEMBRO"
 Me.Com_Funcao.AddItem "PASTOR"
 
 Set oMembros = New ClMembros
 Set Rd = oMembros.Primeiro(Caminho)
 If Not Rd.BOF Then Call RecebeCampo(Rd)
 Set oMembros = Nothing
 
 'If Me.TXT_Matricula.Text <> "" Then Call CaregaGrid
  
' Grid.Row = 0
' Grid.Col = 0
' Grid.Text = "Pag. Referente"
' Grid.Col = 1
' Grid.Text = "Data de Pag."
' Grid.Col = 2
' Grid.Text = "Valor Pag."
'   Call AjustaTamalho(Grid, 2240)
 End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim Rd As New ADODB.Recordset

Select Case Button.Index

    Case 1 'Novo
       FlagStatus = "N"
       Call LimpaCampo(Me)
       Call Desenibeframe
       Set oMembros = New ClMembros
       
       Me.TXT_Matricula.Text = oMembros.ProxMatricula(Caminho)
        Set oMembros = Nothing
   Case 2 'Primeiro
        Set oMembros = New ClMembros
        Set Rd = oMembros.Primeiro(Caminho)
        If Not Rd.EOF Then
            Call RecebeCampo(Rd)
            'Call CaregaGrid
        End If
        Set oMembros = Nothing
    
    Case 3 ' Anterior
        Set oMembros = New ClMembros
        Set Rd = oMembros.MoveAnterior(Caminho, Me.TXT_Matricula)
        If Not Rd.EOF Then
            Call RecebeCampo(Rd)
            'Call CaregaGrid
        End If
        Set oMembros = Nothing
    Case 4 ' Proximo
        
        Set oMembros = New ClMembros
        Set Rd = oMembros.MoveSeguinte(Caminho, Me.TXT_Matricula)
        If Not Rd.EOF Then
            Call RecebeCampo(Rd)
            'Call CaregaGrid
        End If
        Set oMembros = Nothing
    Case 5 ' Ultimo
        
        Set oMembros = New ClMembros
        Set Rd = oMembros.Ultimo(Caminho)
        If Not Rd.EOF Then
            Call RecebeCampo(Rd)
            'Call CaregaGrid
        End If
        Set oMembros = Nothing
    Case 6 ' Alterar
        FlagStatus = "A"
        Call Desenibeframe
    Case 7 ' relatório
        FRM_Relatorio.Show vbModal
        
    Case 8 ' EXCLUIR
        Set oMembros = New ClMembros
        Dim mStatus As Boolean
        
        Call oMembros.Status(Caminho, Me.TXT_Matricula.Text, mStatus)
            
        Set oMembros = Nothing
    Case 9 'Pesquisa

        Set oMembros = New ClMembros

        FRM_PesMembros.Show vbModal
        If oMembros.Matricula <> 0 Then
            Set Rd = oMembros.pesquisa(Caminho, oMembros.Matricula)
            If (Not Rd.EOF) Then
                'Preenche campos
                Call RecebeCampo(Rd)
                  'Call CaregaGrid
            End If
        End If
        Set oMembros = Nothing
End Select
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



End Sub

Sub EnibeFrame()
Me.Frame1.Enabled = False
Me.Frame2.Enabled = False
Me.Frame3.Enabled = False

Me.Command1.Enabled = False

Me.Cmm_Cancela.Enabled = False

Me.Toolbar1.Enabled = True
Me.Cmm_Sair.Enabled = True

End Sub

Sub Desenibeframe()
Me.Frame1.Enabled = True
Me.Frame2.Enabled = True
Me.Frame3.Enabled = True
Me.Command1.Enabled = True
Me.Cmm_Cancela.Enabled = True

Me.Toolbar1.Enabled = False
Me.Cmm_Sair.Enabled = False

End Sub



Private Sub TXT_DTBatismo_KeyPress(KeyAscii As Integer)
Call FormataData(TXT_DTBatismo, KeyAscii)
End Sub

Private Sub TXT_DTConversao_KeyPress(KeyAscii As Integer)
    Call FormataData(TXT_DTConversao, KeyAscii)
End Sub

Private Sub TXT_DTEmisao_KeyPress(KeyAscii As Integer)
    Call FormataData(TXT_DTEmisao, KeyAscii)
End Sub

Private Sub TXT_DTNascimento_KeyPress(KeyAscii As Integer)
    Call FormataData(TXT_DTNascimento, KeyAscii)

End Sub
