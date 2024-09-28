VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmVentas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Venta"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12300
   Icon            =   "frmVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbM 
      Height          =   315
      Index           =   4
      ItemData        =   "frmVentas.frx":058A
      Left            =   8400
      List            =   "frmVentas.frx":0594
      TabIndex        =   68
      Text            =   "cmbMon"
      Top             =   5760
      Width           =   735
   End
   Begin VB.ComboBox cmbM 
      Height          =   315
      Index           =   3
      ItemData        =   "frmVentas.frx":05A0
      Left            =   8400
      List            =   "frmVentas.frx":05AA
      TabIndex        =   67
      Text            =   "cmbMon"
      Top             =   4680
      Width           =   735
   End
   Begin VB.ComboBox cmbM 
      Height          =   315
      Index           =   2
      ItemData        =   "frmVentas.frx":05B6
      Left            =   3600
      List            =   "frmVentas.frx":05C0
      TabIndex        =   66
      Text            =   "cmbMon"
      Top             =   6840
      Width           =   735
   End
   Begin VB.ComboBox cmbM 
      Height          =   315
      Index           =   1
      ItemData        =   "frmVentas.frx":05CC
      Left            =   3600
      List            =   "frmVentas.frx":05D6
      TabIndex        =   65
      Text            =   "cmbMon"
      Top             =   5760
      Width           =   735
   End
   Begin VB.ComboBox cmbM 
      Height          =   315
      Index           =   0
      ItemData        =   "frmVentas.frx":05E2
      Left            =   3600
      List            =   "frmVentas.frx":05EC
      TabIndex        =   64
      Text            =   "cmbMon"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtCM 
      Height          =   285
      Index           =   4
      Left            =   7800
      TabIndex        =   63
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox txtCM 
      Height          =   285
      Index           =   3
      Left            =   7800
      TabIndex        =   62
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtCM 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   61
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox txtCM 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   60
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox txtCM 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   59
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtSena 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   58
      Top             =   3240
      Width           =   975
   End
   Begin VB.CheckBox chkSena 
      Caption         =   "seña"
      Height          =   255
      Left            =   2280
      TabIndex        =   57
      Top             =   3240
      Width           =   735
   End
   Begin VB.ComboBox cmbFDP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      ItemData        =   "frmVentas.frx":05F8
      Left            =   7800
      List            =   "frmVentas.frx":060B
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtCuotas 
      Height          =   285
      Index           =   4
      Left            =   7800
      TabIndex        =   53
      Text            =   "txtCuotas"
      Top             =   5400
      Width           =   375
   End
   Begin VB.ComboBox cmbFDP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      ItemData        =   "frmVentas.frx":0638
      Left            =   7800
      List            =   "frmVentas.frx":064B
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtCuotas 
      Height          =   285
      Index           =   3
      Left            =   7800
      TabIndex        =   49
      Text            =   "txtCuotas"
      Top             =   4320
      Width           =   375
   End
   Begin VB.ComboBox cmbFDP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      ItemData        =   "frmVentas.frx":0678
      Left            =   3000
      List            =   "frmVentas.frx":068B
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtCuotas 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   45
      Text            =   "txtCuotas"
      Top             =   6480
      Width           =   375
   End
   Begin VB.ComboBox cmbFDP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      ItemData        =   "frmVentas.frx":06B8
      Left            =   3000
      List            =   "frmVentas.frx":06CB
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtCuotas 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   41
      Text            =   "txtCuotas"
      Top             =   5400
      Width           =   375
   End
   Begin VB.ComboBox cmbCFP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmVentas.frx":06F8
      Left            =   3000
      List            =   "frmVentas.frx":070B
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtModelo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   38
      Text            =   "txtModelo"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtMarca 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   37
      Text            =   "txtMarca"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtProd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   36
      Text            =   "txtProd"
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtCuotas 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   31
      Text            =   "txtCuotas"
      Top             =   4320
      Width           =   375
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      ItemData        =   "frmVentas.frx":071E
      Left            =   4080
      List            =   "frmVentas.frx":0728
      TabIndex        =   27
      Text            =   "cmbMon"
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox cmbFDP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      ItemData        =   "frmVentas.frx":0734
      Left            =   3000
      List            =   "frmVentas.frx":0747
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox cmbVendedores 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtCant 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   20
      Text            =   "txtCant"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtPDV 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   18
      Text            =   "txtPDV"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del cliente (Opcionales)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   5160
      TabIndex        =   9
      Top             =   480
      Width           =   6735
      Begin VB.TextBox txtDoc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   79
         Text            =   "txtDoc"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox cmbRI 
         Height          =   315
         ItemData        =   "frmVentas.frx":0774
         Left            =   5040
         List            =   "frmVentas.frx":077E
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtDom 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   75
         Text            =   "txtDom"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtRS 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   74
         Text            =   "txtRS"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtCuit 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   72
         Text            =   "txtCuit"
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cmbTC 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmVentas.frx":078A
         Left            =   1560
         List            =   "frmVentas.frx":0794
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cmbClientes 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton OptNU 
         Caption         =   "Antiguo Cliente"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptNU 
         Caption         =   "Nuevo cliente"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtTel 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   24
         Text            =   "txtTel"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtMail 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         Text            =   "txtMail"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtApe 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Text            =   "txtApe"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtNom 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Text            =   "txtNom"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Documento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   80
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Responsable Inscripto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   77
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Domicilio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   76
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Razon Social:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   73
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "CUIT:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   71
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   69
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Telefono:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "E-mail:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
   End
   Begin ChamaleonButton.ChameleonBtn cmdSave 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   8040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Guardar cambios"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVentas.frx":07BE
      PICN            =   "frmVentas.frx":07DA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdSalir 
      Height          =   495
      Left            =   10680
      TabIndex        =   1
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Salir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVentas.frx":0D74
      PICN            =   "frmVentas.frx":0D90
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Forma de pago:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   56
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label lblCuotas 
      Caption         =   "lblCuotas"
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   55
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Forma de pago:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   52
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label lblCuotas 
      Caption         =   "lblCuotas"
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   51
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Forma de pago:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   48
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label lblCuotas 
      Caption         =   "lblCuotas"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   47
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Forma de pago:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   44
      Top             =   5040
      Width           =   2535
   End
   Begin VB.Label lblCuotas 
      Caption         =   "lblCuotas"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   43
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad Formas de pago:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label lblProd 
      Caption         =   "lblProd"
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblCuotas 
      Caption         =   "lblCuotas"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   32
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   7560
      Width           =   4095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblModelo 
      Caption         =   "lblModelo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblMarca 
      Caption         =   "lblMarca"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblFDA 
      Caption         =   "lblFDA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblPID 
      Caption         =   "lblPID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Forma de pago:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Vendido por:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Precio de venta por unidad:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Modelo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha de alta:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Producto ID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cantidad As Integer
Dim Nuevo As Boolean






Private Sub chkSena_Click()
If txtSena.Visible = False Then
txtSena.Visible = True
Exit Sub
Else
txtSena.Visible = False
End If
End Sub

Private Sub cmbCFP_Click()
Dim l As Integer
For l = 0 To cmbCFP.ListCount - 1
Label7(l).Visible = False
cmbFDP(l).Visible = False
txtCuotas(l).Visible = False
lblCuotas(l).Visible = False
cmbM(l).Visible = False
txtCM(l).Visible = False
Next l
For l = 0 To cmbCFP.ListIndex
Label7(l).Caption = "Forma de pago " & l + 1
Label7(l).Visible = True
cmbFDP(l).Visible = True
cmbM(l).Visible = True
txtCM(l).Visible = True
If (cmbFDP(l).ListIndex = 0) Then
txtCuotas(l).Visible = False
lblCuotas(l).Visible = False
txtCuotas(l).Width = 375
txtCuotas(l).Text = ""
lblCuotas(l).Caption = ""
End If
If (cmbFDP(l).ListIndex = 1) Then
txtCuotas(l).Visible = True
lblCuotas(l).Visible = True
txtCuotas(l).Width = 375
End If
If (cmbFDP(l).ListIndex = 2) Then
txtCuotas(l).Visible = True
lblCuotas(l).Visible = True
txtCuotas(l).Width = 375
End If
If (cmbFDP(l).ListIndex = 3) Then
txtCuotas(l).Visible = True
lblCuotas(l).Visible = True
txtCuotas(l).Width = 375
End If
If (cmbFDP(l).ListIndex = 4) Then
txtCuotas(l).Visible = True
lblCuotas(l).Visible = False
txtCuotas(l).Width = 1150
End If
Next l
End Sub

Private Sub cmbFDP_Click(Index As Integer)
Dim p As Integer
For p = 0 To cmbCFP.ListIndex
If (cmbFDP(p).ListIndex = 0) Then
txtCuotas(p).Visible = False
lblCuotas(p).Visible = False
cmbM(p).Visible = True
txtCM(p).Visible = True
txtCuotas(p).Width = 375
txtCuotas(p).Text = ""
lblCuotas(p).Caption = ""
End If
If (cmbFDP(p).ListIndex = 1) Then
txtCuotas(p).Visible = True
lblCuotas(p).Visible = True
cmbM(p).Visible = True
txtCM(p).Visible = True
lblCuotas(p).Caption = "Cuotas"
txtCuotas(p).Text = "1"
txtCuotas(p).Width = 375
End If
If (cmbFDP(p).ListIndex = 2) Then
txtCuotas(p).Visible = True
lblCuotas(p).Visible = True
cmbM(p).Visible = True
txtCM(p).Visible = True
lblCuotas(p).Caption = "dias"
txtCuotas(p).Text = "30"
txtCuotas(p).Width = 375
End If
If (cmbFDP(p).ListIndex = 3) Then
txtCuotas(p).Visible = True
lblCuotas(p).Visible = True
cmbM(p).Visible = True
txtCM(p).Visible = True
lblCuotas(p).Caption = "Cuotas"
txtCuotas(p).Text = "1"
txtCuotas(p).Width = 375
End If
If (cmbFDP(p).ListIndex = 4) Then
txtCuotas(p).Visible = True
lblCuotas(p).Visible = False
txtCuotas(p).Text = ""
lblCuotas(p).Caption = ""
txtCuotas(p).Width = 1150
cmbM(p).Visible = True
txtCM(p).Visible = True
End If
Next p
End Sub

Private Sub cmbTC_Click()
If cmbTC.ListIndex = 1 Then
txtCuit.Visible = True
Label16.Visible = True
txtRS.Visible = True
Label17.Visible = True
txtDom.Visible = True
Label18.Visible = True
cmbRI.Visible = True
Label19.Visible = True
txtDoc.Visible = True
Label20.Visible = True
Else
txtCuit.Visible = False
Label16.Visible = False
txtRS.Visible = False
Label17.Visible = False
txtDom.Visible = False
Label18.Visible = False
cmbRI.Visible = False
Label19.Visible = False
txtDoc.Visible = False
Label20.Visible = False
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim FDPago As String
Dim cliente As Boolean
cliente = False
Dim precio As Single
If Not IsNumeric(Trim(txtPDV.Text)) Or Trim(txtPDV.Text) = "" Then
lblError.Caption = "El precio de venta tiene que ser un numero"
Exit Sub
End If
If Not IsNumeric(Trim(txtCant.Text)) Or Trim(txtCant.Text) = "" Then
lblError.Caption = "La cantidad tiene que ser un numero"
Exit Sub
End If

If (txtCant.Text > cantidad) Then
lblError.Caption = "La cantidad es mayor que la cantidad en stock"
Exit Sub
Else
cantidad = cantidad - txtCant.Text
End If
If Nuevo = True Then
If (cmbTC.ListIndex = 0) Then
If (Trim(txtNom.Text) <> "" Or Trim(txtApe.Text) <> "" Or Trim(txtMail.Text) <> "" Or txtTel.Text <> "") Then
cliente = True
End If
Else
If (Trim(txtNom.Text) <> "" Or Trim(txtApe.Text) <> "" Or Trim(txtMail.Text) <> "" Or txtTel.Text <> "" Or txtCuit.Text <> "" Or txtRS.Text <> "") Then
cliente = True
End If
End If
End If


If ES = False Then
If (Trim(txtProd.Text) = "" Or Trim(txtMarca.Text) = "" Or Trim(txtModelo.Text) = "") Then
lblError.Caption = "Tiene que completar los datos"
Exit Sub
End If
End If

If txtSena.Visible = True And Not IsNumeric(txtSena) Then
lblError.Caption = "La seña tiene que ser numerica"
Exit Sub
End If
If (txtSena.Visible = False) Then
txtSena.Text = " "
End If

lblError.Caption = ""

If cmbMon.ListIndex = 1 Then
precio = txtPDV.Text * pdolar
Else
precio = txtPDV.Text
End If

Dim h As Integer
For h = 0 To cmbCFP.ListIndex
If (h = 0) Then
FDPago = cmbFDP(h).Text & " " & txtCuotas(h).Text & " " & lblCuotas(h).Caption & " " & txtCM(h).Text & " " & cmbM(h).Text
Else
FDPago = FDPago & " / " & cmbFDP(h).Text & " " & txtCuotas(h).Text & " " & lblCuotas(h).Caption & " " & txtCM(h).Text & " " & cmbM(h).Text

End If
Next h


 If rs.State = adStateOpen Then rs.Close
 If (ES = True) Then
 If Nuevo = False Then
    cnn.Execute "INSERT INTO Ventas " & "(IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,Sena,FDP,Cliente) VALUES('" & _
                                 lblPID.Caption & "','" & _
                                lblProd.Caption & "','" & _
                                lblMarca.Caption & "','" & _
                               lblModelo.Caption & "','" & _
                                txtCant.Text & "','" & _
                                 cmbVendedores.Text & "','" & _
                               Format(precio, "#,##0.00") & "','" & _
                               Format(Date, "mm/dd/yyyy") & "','" & _
                                txtSena.Text & "','" & _
                                FDPago & "','" & _
                                  cmbClientes.Text & "')"
   End If
   
   
    If Nuevo = True Then
    cnn.Execute "INSERT INTO Ventas " & "(IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,Sena,FDP,Cliente) VALUES('" & _
                                 lblPID.Caption & "','" & _
                                lblProd.Caption & "','" & _
                                lblMarca.Caption & "','" & _
                               lblModelo.Caption & "','" & _
                                txtCant.Text & "','" & _
                                 cmbVendedores.Text & "','" & _
                               Format(precio, "#,##0.00") & "','" & _
                               Format(Date, "mm/dd/yyyy") & "','" & _
                               txtSena.Text & "','" & _
                                FDPago & "','" & _
                                  txtNom.Text & " " & txtApe.Text & "')"
   End If
   Else
   Dim ProductID As String
   ProductID = " "
    If Nuevo = False Then
    cnn.Execute "INSERT INTO Ventas " & "(IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,Sena,FDP,Cliente) VALUES('" & _
                                 ProductID & "','" & _
                                txtProd.Text & "','" & _
                                txtMarca.Text & "','" & _
                               txtModelo.Text & "','" & _
                                txtCant.Text & "','" & _
                                 cmbVendedores.Text & "','" & _
                               Format(precio, "#,##0.00") & "','" & _
                               Format(Date, "mm/dd/yyyy") & "','" & _
                                 txtSena.Text & "','" & _
                                 FDPago & "','" & _
                                  cmbClientes.Text & "')"
   End If
   
   
    If Nuevo = True Then
    cnn.Execute "INSERT INTO Ventas " & "(IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,Sena,FDP,Cliente) VALUES('" & _
                                 ProductID & "','" & _
                                txtProd.Text & "','" & _
                                txtMarca.Text & "','" & _
                               txtModelo.Text & "','" & _
                                txtCant.Text & "','" & _
                                 cmbVendedores.Text & "','" & _
                               Format(precio, "#,##0.00") & "','" & _
                               Format(Date, "mm/dd/yyyy") & "','" & _
                               txtSena.Text & "','" & _
                                FDPago & "','" & _
                                  txtNom.Text & " " & txtApe.Text & "')"
   End If
   
   End If
   
   
   
   
   
   
If Nuevo = True Then
If cliente = True Then

If (Trim(txtNom.Text) = "") Then
txtNom.Text = " "
End If
If (Trim(txtApe.Text) = "") Then
txtApe.Text = " "
End If
If (Trim(txtMail.Text) = "") Then
txtMail.Text = " "
End If
If (Trim(txtTel.Text) = "") Then
txtTel.Text = " "
End If
If (Trim(txtCuit.Text) = "") Then
txtCuit.Text = " "
End If
If (Trim(txtRS.Text) = "") Then
txtRS.Text = " "
End If
If (Trim(txtDom.Text) = "") Then
txtDom.Text = " "
End If
If (Trim(txtDoc.Text) = "") Then
txtDoc.Text = " "
End If
 If (cmbTC.ListIndex = 0) Then
 cnn.Execute "INSERT INTO Clientes " & "(Nombre,Apellido,Email,Telefono,FechaDeAlta) VALUES ('" & _
                        txtNom.Text & "','" & txtApe.Text & "','" & txtMail.Text & "','" & txtTel.Text & "','" & Format(Date, "mm/dd/yyyy") & "')"
Else

  cnn.Execute "INSERT INTO ClientesM " & "(CUIT,RS,Nombre,Apellido,Domicilio,Email,Telefono,RI,DNI,FechaDeAlta) VALUES ('" & _
                        txtCuit.Text & "','" & txtRS.Text & "','" & txtNom.Text & "','" & txtApe.Text & "','" & txtDom.Text & "','" & txtMail.Text & "','" & txtTel.Text & "','" & cmbRI.Text & "','" & txtDoc.Text & "','" & Format(Date, "mm/dd/yyyy") & "')"


End If
End If
End If



     If (ES = True) Then
    cnn.Execute "UPDATE Productos set Cantidad = " & cantidad & " where id = " & lblPID.Caption
 rs.Open "select * from Productos", cnn, adOpenStatic, adLockOptimistic
 Call CargarListView2(frmMain.LV2, rs)
If rs.State = adStateOpen Then rs.Close
 End If
  
  rs.Open "select * from Ventas", cnn, adOpenStatic, adLockOptimistic
 Call CargarListView3(frmMain.LV3, rs)
If rs.State = adStateOpen Then rs.Close
 Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next
cmbTC.Visible = True
cmbTC.ListIndex = 0
txtSena.Visible = False
If ES = True Then
lblProd.Visible = True
lblFDA.Visible = True
lblMarca.Visible = True
lblModelo.Visible = True
txtProd.Visible = False

txtMarca.Visible = False
txtModelo.Visible = False
Else
lblProd.Visible = False
lblFDA.Visible = False
lblMarca.Visible = False
lblModelo.Visible = False
txtProd.Text = ""
txtMarca.Text = ""
txtModelo.Text = ""
txtProd.Visible = True
txtMarca.Visible = True
txtModelo.Visible = True
End If

cmbCFP.ListIndex = 0
cmbMon.ListIndex = 0
Dim X As Integer
For X = 0 To 4
cmbFDP(X).ListIndex = 0
cmbM(X).ListIndex = 0
Next X

For X = 1 To 4
Label7(X).Visible = False
cmbFDP(X).Visible = False
txtCuotas(X).Visible = False
lblCuotas(X).Visible = False
cmbM(X).Visible = False
txtCM(X).Visible = False
Next X

cmbClientes.Visible = False
Dim i As Integer
Nuevo = True
i = 0
frmMain.Enabled = False
  If rs.State = adStateOpen Then rs.Close
   rs.Open "select * from Vendedores", cnn, adOpenStatic, adLockOptimistic
    While Not rs.EOF
    cmbVendedores.List(i) = rs!Nombre & " " & rs!Apellido
    i = i + 1
    rs.MoveNext
    Wend
    rs.Close
    txtCuit.Visible = False
Label16.Visible = False
txtRS.Visible = False
Label17.Visible = False
txtDom.Visible = False
Label18.Visible = False
cmbRI.Visible = False
Label19.Visible = False
txtDoc.Visible = False
Label20.Visible = False
txtCuit.Text = ""
txtRS.Text = ""
txtDom.Text = ""
cmbRI.ListIndex = 0
txtDoc.Text = ""
    txtNom.Text = ""
    txtApe.Text = ""
    txtMail.Text = ""
    txtTel.Text = ""
    cantidad = frmMain.LV2.SelectedItem.ListSubItems(5).Text
End Sub

Private Sub Form_Unload(Cancel As Integer)

frmMain.Enabled = True
End Sub

Private Sub OptNU_Click(Index As Integer)
On Error Resume Next
Dim i As Integer
i = 0
If Index = 0 Then
Nuevo = True
txtNom.Visible = True
txtApe.Visible = True
txtMail.Visible = True
txtTel.Visible = True
Label8.Visible = True
Label8.Caption = "Nombre:"
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
cmbClientes.Visible = False
cmbTC.Visible = True
Label15.Visible = True
If (cmbTC.ListIndex = 1) Then
txtCuit.Visible = True
Label16.Visible = True
txtRS.Visible = True
Label17.Visible = True
txtDom.Visible = True
Label18.Visible = True
cmbRI.Visible = True
Label19.Visible = True
txtDoc.Visible = True
Label20.Visible = True
End If

Else
Nuevo = False
txtNom.Visible = False
txtApe.Visible = False

txtCuit.Visible = False
Label16.Visible = False
txtRS.Visible = False
Label17.Visible = False
txtDom.Visible = False
Label18.Visible = False
cmbRI.Visible = False
Label19.Visible = False
txtDoc.Visible = False
Label20.Visible = False

txtMail.Visible = False
txtTel.Visible = False
Label8.Visible = True
Label8.Caption = "Clientes:"
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
cmbClientes.Visible = True
cmbTC.Visible = False
Label15.Visible = False


rs.Open "SELECT * from Clientes", cnn, adOpenStatic, adLockOptimistic
  While Not rs.EOF
    cmbClientes.List(i) = rs!Nombre & " " & rs!Apellido
    i = i + 1
    rs.MoveNext
    Wend
rs.Close
cmbClientes.ListIndex = 0
End If
End Sub
