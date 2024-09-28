VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmMain 
   Caption         =   "Controlador y administrador de Stock"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   18210
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10440
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Precio Ventas"
      Height          =   855
      Left            =   14400
      TabIndex        =   7
      Top             =   6000
      Width           =   3735
      Begin VB.Label lblPN 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblPF 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
   End
   Begin MSComctlLib.ListView LV2 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Marca"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modelo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Distribuidor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cantidad"
         Object.Width           =   1939
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Precio Unitario"
         Object.Width           =   2647
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Precio Unitario u$s"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Margen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Precio sugerido de venta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Fecha de alta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Deposito"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "IVA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Anotaciones"
         Object.Width           =   3449
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdAdd 
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "&Agregar Producto"
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
      FOCUSR          =   0   'False
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777152
      FCOLO           =   16777152
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":058A
      PICN            =   "frmMain.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdConf 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Configurar"
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
      FOCUSR          =   0   'False
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777152
      FCOLO           =   16777152
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":5E24
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView LV3 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID Producto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Producto"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Marca"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Modelo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cantidad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Vendido Por"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Precio de Venta"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Fecha de Baja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Seña"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Forma de Pago"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Cliente"
         Object.Width           =   2540
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdEditProd 
      Height          =   735
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   5880
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "&Editar Producto"
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
      FOCUSR          =   0   'False
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777152
      FCOLO           =   16777152
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":5E40
      PICN            =   "frmMain.frx":5E5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdVendedores 
      Height          =   735
      Index           =   0
      Left            =   13080
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "&Vendedores"
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
      FOCUSR          =   0   'False
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777152
      FCOLO           =   16777152
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":B68F
      PICN            =   "frmMain.frx":B6AB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn CmdElProd 
      Height          =   735
      Left            =   2520
      TabIndex        =   11
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "&Eliminar Producto"
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
      FCOL            =   16777088
      FCOLO           =   16777088
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":B9FD
      PICN            =   "frmMain.frx":BA19
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdClientes 
      Height          =   735
      Index           =   1
      Left            =   11760
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "&Clientes"
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
      FOCUSR          =   0   'False
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777152
      FCOLO           =   16777152
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":112BE
      PICN            =   "frmMain.frx":112DA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   375
      Index           =   5
      Left            =   11520
      TabIndex        =   13
      Top             =   6960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Imprimir Ventas"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":11680
      PICN            =   "frmMain.frx":1169C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones1 
      Height          =   375
      Index           =   0
      Left            =   14040
      TabIndex        =   14
      Top             =   6960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Imprimir Stock"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":11C36
      PICN            =   "frmMain.frx":11C52
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdBuscar 
      Height          =   375
      Index           =   4
      Left            =   10320
      TabIndex        =   15
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Buscar"
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
      FOCUSR          =   0   'False
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":121EC
      PICN            =   "frmMain.frx":12208
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelV 
      Height          =   375
      Index           =   0
      Left            =   9000
      TabIndex        =   16
      Top             =   6960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Exportar Ventas"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":127A2
      PICN            =   "frmMain.frx":127BE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExceS 
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   17
      Top             =   6960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Exportar Stock"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":17F94
      PICN            =   "frmMain.frx":17FB0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelIV 
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   18
      Top             =   6960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Importar Ventas"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":1D786
      PICN            =   "frmMain.frx":1D7A2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelIS 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   19
      Top             =   6960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Importar Stock"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":22F78
      PICN            =   "frmMain.frx":22F94
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdVenProd 
      Height          =   735
      Left            =   3720
      TabIndex        =   20
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Vender Producto"
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
      FCOL            =   16777088
      FCOLO           =   16777088
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":2876A
      PICN            =   "frmMain.frx":28786
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtnE 
      Height          =   375
      Left            =   8160
      TabIndex        =   21
      Top             =   5880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Egresos de caja"
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
      FCOL            =   16777088
      FCOLO           =   16777088
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":28D20
      PICN            =   "frmMain.frx":28D3C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn CBBusq 
      Height          =   375
      Left            =   8160
      TabIndex        =   22
      Top             =   6360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Resumen de ventas"
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
      FCOL            =   16777088
      FCOLO           =   16777088
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":292D6
      PICN            =   "frmMain.frx":292F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtnEV 
      Height          =   735
      Left            =   4920
      TabIndex        =   23
      Top             =   5880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "E&liminar Venta"
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
      FCOL            =   16777088
      FCOLO           =   16777088
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":2988C
      PICN            =   "frmMain.frx":298A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn btAD 
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Distribuidores"
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
      MICON           =   "frmMain.frx":2F14D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdDep 
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "D&epositos"
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
      MICON           =   "frmMain.frx":2F169
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Productos vendidos:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Stock:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SIV1 As Single
Dim SIV2 As Single
Dim SIV3 As Single
Dim SIV4 As Single
Dim SIV5 As Single

Dim PF As Single



Private Sub btAD_Click()

frmProveedores.Show
Me.Enabled = False

End Sub

Private Sub CBBusq_Click()
frmBusq.Show
End Sub



Private Sub ChameleonBtnE_Click()
frmEgresos.Show

End Sub

Private Sub ChameleonBtnEV_Click()
Call EliminarVentas
End Sub

Private Sub cmdAdd_Click(Index As Integer)
EditarProductos = 1
Me.Enabled = False
frmEditarProd.Show

End Sub

Private Sub cmdBuscar_Click(Index As Integer)
frmFilterPV.Show


End Sub

Private Sub cmdClientes_Click(Index As Integer)
frmClientes.Show
End Sub

Private Sub cmdConf_Click(Index As Integer)
frmConf.Show
Me.Enabled = False
End Sub


Private Sub cmdDep_Click()
frmDepositos.Show
Me.Enabled = False
End Sub

Private Sub cmdEditProd_Click(Index As Integer)
On Error Resume Next
EditarProductos = 2
Me.Enabled = False
ProdId = LV2.SelectedItem.Text
frmEditarProd.Show
End Sub

Private Sub CmdElProd_Click()
Call Eliminar
End Sub

Private Sub cmdExcelIS_Click(Index As Integer)
' On Error GoTo errorsub1
CommonDialog1.Filter = "Archivos de Excel|*.xls"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
'dimensiones
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
Dim lngUltimaColumna As Long
Dim sql As String
Dim X As Long
Dim Y As Long
'abrir programa Excel
Set xlApp = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en otra carpeta)
Set xlLibro = xlApp.Workbooks.Open _
(CommonDialog1.FileName, True, True, , "")
 
'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(CommonDialog1.FileName, True, True, , "")
Set xlHoja = xlApp.Worksheets(1)
 
'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range("A1:C10").Value

'2. Si no conoces el rango

lngUltimaColumna = 1

lngUltimaFila = _
xlHoja.Columns("A:A").Range("A65536").End(xlUp).Row
While xlHoja.Cells(1, lngUltimaColumna) <> ""
lngUltimaColumna = lngUltimaColumna + 1
Wend



'SQL = "INSERT INTO Ventas (IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,FDP,Anotaciones,Cliente) VALUES ('"



varMatriz = xlHoja.Range(xlHoja.Cells(1, 1), _
xlHoja.Cells(lngUltimaFila, lngUltimaColumna))
'utilizamos los datos...
For X = 1 To lngUltimaFila
For Y = 1 To lngUltimaColumna - 1
If (X <> 1) Then
If (Y <> 1) Then
If (Y <> lngUltimaColumna - 1) Then
If (Y <> 8) Then
If (Y <> 10) Then
sql = sql & varMatriz(X, Y) & "','"
End If
End If
Else
sql = sql & varMatriz(X, Y) & "')"
End If
End If
End If
 Next
 If (X <> 1) Then
 cnn.Execute sql
 End If
 sql = "INSERT INTO Productos (Producto,Marca,Modelo,distribuidor,Cantidad,Margen,PrecioU,FechaDeAlta,deposito,IVA,Anotaciones) VALUES ('"

 Next
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing





    rs.Open "select * from Productos", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView2(LV2, rs)
If rs.State = adStateOpen Then rs.Close


End If
Exit Sub
errorsub1:
MsgBox ("Error durante la importacion del archivo")
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing
End Sub

Private Sub cmdExcelIV_Click(Index As Integer)
 On Error GoTo ErrorSub
CommonDialog1.Filter = "Archivos de Excel|*.xls"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
'dimensiones
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
Dim lngUltimaColumna As Long
Dim sql As String
Dim X As Long
Dim Y As Long
'abrir programa Excel
Set xlApp = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en otra carpeta)
Set xlLibro = xlApp.Workbooks.Open _
(CommonDialog1.FileName, True, True, , "")
 
'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(CommonDialog1.FileName, True, True, , "")
Set xlHoja = xlApp.Worksheets(1)
 
'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range("A1:C10").Value

'2. Si no conoces el rango

lngUltimaColumna = 1

lngUltimaFila = _
xlHoja.Columns("A:A").Range("A65536").End(xlUp).Row
While xlHoja.Cells(1, lngUltimaColumna) <> ""
lngUltimaColumna = lngUltimaColumna + 1
Wend



'SQL = "INSERT INTO Ventas (IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,FDP,Anotaciones,Cliente) VALUES ('"



varMatriz = xlHoja.Range(xlHoja.Cells(1, 1), _
xlHoja.Cells(lngUltimaFila, lngUltimaColumna))
'utilizamos los datos...
For X = 1 To lngUltimaFila
For Y = 1 To lngUltimaColumna - 1
If (X <> 1) Then
If (Y <> 1) Then
If (Y <> lngUltimaColumna - 1) Then
sql = sql & varMatriz(X, Y) & "','"
Else
sql = sql & varMatriz(X, Y) & "')"

End If
End If
End If
 Next
 If (X <> 1) Then
 cnn.Execute sql
 End If
 sql = "INSERT INTO Ventas (IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,FDP,Anotaciones,Cliente) VALUES ('"

 Next
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing





    rs.Open "select * from Ventas", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView3(LV3, rs)
rs.Close


End If
Exit Sub
ErrorSub:
MsgBox ("Error durante la importacion del archivo")
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing

End Sub

Private Sub cmdExcelV_Click(Index As Integer)
'EXPORT to EXCEL
Dim xlApp As Excel.Application
Dim xlSh As Excel.Worksheet
Dim i As Long
Dim p As Long
Set xlApp = New Excel.Application

xlApp.Visible = True
xlApp.Workbooks.Add
Set xlSh = xlApp.Workbooks(1).Worksheets(1)

  
 For i = 1 To LV3.ListItems.Count
 For p = 1 To LV3.ListItems(i).ListSubItems.Count
 If (i = 1) Then
 xlSh.Cells(i, p).Font.Bold = True
   xlSh.Cells(i, p).Value = LV3.ColumnHeaders(p).Text
   If (p = LV3.ListItems(i).ListSubItems.Count) Then
    xlSh.Cells(i, p + 1).Font.Bold = True
    xlSh.Cells(i, p + 1).Value = LV3.ColumnHeaders(p + 1).Text
    End If
   ' xlSh.Cells(i, p).Value = LV3.ListItems.Item(p)
End If
    If (p = 1) Then
    xlSh.Cells(i + 1, p).Value = LV3.ListItems(i).Text
    End If
    xlSh.Cells(i + 1, p + 1).Value = LV3.ListItems(i).ListSubItems(p).Text
       
          Next p
        Next i
 

Set xlSh = Nothing
Set xlApp = Nothing

End Sub

Private Sub cmdExceS_Click(Index As Integer)
'EXPORT to EXCEL
Dim xlApp As Excel.Application
Dim xlSh As Excel.Worksheet
Dim i As Long
Dim p As Long
Set xlApp = New Excel.Application

xlApp.Visible = True
xlApp.Workbooks.Add
Set xlSh = xlApp.Workbooks(1).Worksheets(1)

  
 For i = 1 To LV2.ListItems.Count
 For p = 1 To LV2.ListItems(i).ListSubItems.Count
 If (i = 1) Then
 xlSh.Cells(i, p).Font.Bold = True
   xlSh.Cells(i, p).Value = LV2.ColumnHeaders(p).Text
   If (p = LV2.ListItems(i).ListSubItems.Count) Then
    xlSh.Cells(i, p + 1).Font.Bold = True
    xlSh.Cells(i, p + 1).Value = LV2.ColumnHeaders(p + 1).Text
    End If
   ' xlSh.Cells(i, p).Value = LV3.ListItems.Item(p)
End If
    If (p = 1) Then
    xlSh.Cells(i + 1, p).Value = LV2.ListItems(i).Text
    End If
    xlSh.Cells(i + 1, p + 1).Value = LV2.ListItems(i).ListSubItems(p).Text
       
          Next p
        Next i
 

Set xlSh = Nothing
Set xlApp = Nothing
End Sub

Private Sub cmdOpciones_Click(Index As Integer)
  If rs.State = adStateOpen Then rs.Close
   rs.Open "select * from Ventas", cnn, adOpenStatic, adLockOptimistic
    Set DataReport2.DataSource = rs
    DataReport2.Show 1
       If rs.State = adStateOpen Then rs.Close
End Sub

Private Sub cmdOpciones1_Click(Index As Integer)
  If rs.State = adStateOpen Then rs.Close
   rs.Open "select * from Productos", cnn, adOpenStatic, adLockOptimistic
    Set DataReport3.DataSource = rs
    DataReport3.Show 1
       If rs.State = adStateOpen Then rs.Close
End Sub

Private Sub cmdVendedores_Click(Index As Integer)
FrmPrincipal.Show
End Sub

Private Sub cmdVenProd_Click()

On Error Resume Next
ES = False
       
    With frmVentas
        ' obtiene el elemento seleccionado
        .lblPID = ""
        .lblFDA = ""
         .lblMarca = ""
        .lblModelo = ""
        .lblProd = ""
        .txtPDV = ""
        .txtCant = ""
        .Show vbModal
    End With
End Sub

Private Sub Form_Load()
On Error Resume Next
Call IniciarIVA
Dim s As Integer, p As Integer
If (Prin(5) = 0) Then
cmdAdd(0).Visible = False
cmdExcelIS(1).Visible = False
End If
If (Prin(6) = 0) Then
cmdEditProd(1).Visible = False
End If
If (Prin(7) = 0) Then
CmdElProd.Visible = False
End If
If (Prin(8) = 0) Then
ChameleonBtnEV.Visible = False
End If
If (Prin(9) = 0) Then
cmdVenProd.Visible = False
cmdExcelIV(0).Visible = False
End If
If (Prin(10) = 0) Then
ChameleonBtnE.Visible = False
End If
If (Prin(11) = 0) Then
cmdDep.Visible = False
End If
If (Prin(12) = 0) Then
cmdClientes(1).Visible = False
End If
If (Prin(13) = 0) Then
cmdVendedores(0).Visible = False
End If
If (Prin(14) = 0) Then
btAD.Visible = False
End If
If (Prin(15) = 0) Then
cmdConf(1).Visible = False
End If






    ' carga el Recorset con todos los datos
     rs.Open "select * from Configuracion", cnn, adOpenStatic, adLockOptimistic
    pdolar = rs!pdolar
    SIV1 = rs!IVA1
    SIV2 = rs!IVA2
    If (rs!IVA3 <> " ") Then
    SIV3 = rs!IVA3
    End If
     If (rs!IVA4 <> " ") Then
    SIV4 = rs!IVA4
    End If
     If (rs!IVA5 <> " ") Then
    SIV5 = rs!IVA5
    End If
    rs.Close
    rs.Open "select * from Productos", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView2(LV2, rs)
    rs.Close
    
   
    
    
    
    
    
     rs.Open "select * from Ventas", cnn, adOpenStatic, adLockOptimistic
 Call CargarListView3(frmMain.LV3, rs)
 rs.Close
    


    
    
    Dim GN As Single
Dim GIV As Single
If (LV2.SelectedItem.ListSubItems(12).Text = 0) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = 0
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(6).Text + GN + GIV), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 1) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV1 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(1) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(1) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(6).Text + GN + GIV), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 2) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV2 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(2) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(2) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(6).Text + GN + GIV), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 3) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV3 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(3) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(3) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(6).Text + GN + GIV), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 4) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV4 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(4) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(4) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(6).Text + GN + GIV), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 5) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV5 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(5) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(5) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(6).Text + GN + GIV), "#,##0.00")

End If
    
    
    
    
    
    
    
    
      If rs.State = adStateOpen Then rs.Close
 
    
    
End Sub



Private Sub LV2_Click()
On Error Resume Next
Dim GN As Single
Dim GIV As Single

If (LV2.SelectedItem.ListSubItems(12).Text = 0) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = 0
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(0) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(0) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 1) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV1 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(1) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(1) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 2) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV2 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(2) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(2) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 3) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV3 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(3) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(3) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 4) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV4 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(4) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(4) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00")

End If
If (LV2.SelectedItem.ListSubItems(12).Text = 5) Then

GN = (LV2.SelectedItem.ListSubItems(6).Text * LV2.SelectedItem.ListSubItems(8).Text) / 100
GIV = (SIV5 * (LV2.SelectedItem.ListSubItems(6).Text + GN)) / 100
lblPF = "Precio Final: " & Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text) / pdolar, "#,##0.00") & " u$s"
lblPN = "Precio Neto: " & Format((LV2.SelectedItem.ListSubItems(9).Text) / (1 + (IVA(5) / 100)), "#,##0.00") & " $ / " & Format((LV2.SelectedItem.ListSubItems(9).Text / (1 + (IVA(5) / 100)) / pdolar), "#,##0.00") & " u$s"
PF = Format((LV2.SelectedItem.ListSubItems(9).Text), "#,##0.00")


End If
End Sub


Private Sub Eliminar()

    

    If (LV2.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV2.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LV2.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "Producto " & .ListSubItems(1).Text & vbNewLine & _
                 "Cantidad: " & .ListSubItems(4).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Productos where Id = " & .Text & ""
            ' refresca el recordset
            rs.Open "select * from Productos", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView2(LV2, rs)
        End If
    End With
    If rs.State = adStateOpen Then rs.Close
End Sub

Private Sub LV2_DblClick()
On Error Resume Next
If (Prin(9) = 0) Then
Exit Sub
End If

ES = True

    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV2.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LV2.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
       
    With frmVentas
        ' obtiene el elemento seleccionado
        .lblPID = LV2.SelectedItem.Text
        .lblFDA = LV2.SelectedItem.ListSubItems(8).Text
         .lblMarca = LV2.SelectedItem.ListSubItems(2).Text
        .lblModelo = LV2.SelectedItem.ListSubItems(3).Text
        .lblProd = LV2.SelectedItem.ListSubItems(1).Text
        .txtPDV = PF
        .txtCant = ""
        .Show vbModal
    End With
End Sub

Private Sub EliminarVentas()

    

    If (LV3.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV3.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LV3.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "Producto " & .ListSubItems(2).Text & vbNewLine & _
                 "Cantidad: " & .ListSubItems(5).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Ventas where Id = " & .Text & ""
            ' refresca el recordset
            rs.Open "select * from Ventas", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView3(LV3, rs)
        End If
    End With
    If rs.State = adStateOpen Then rs.Close
End Sub

