VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   11550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11100
   Icon            =   "frmClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11550
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5535
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Apellido"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "E-Mail"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Telefono"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Fecha de alta"
         Object.Width           =   2540
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Nuevo"
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
      MICON           =   "frmClientes.frx":058A
      PICN            =   "frmClientes.frx":05A6
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
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmClientes.frx":095A
      PICN            =   "frmClientes.frx":0976
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
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Eliminar"
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
      MICON           =   "frmClientes.frx":0D2A
      PICN            =   "frmClientes.frx":0D46
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
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
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
      MICON           =   "frmClientes.frx":10F8
      PICN            =   "frmClientes.frx":1114
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmClientes.frx":16AE
      PICN            =   "frmClientes.frx":16CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   11040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      FOCUSR          =   0   'False
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmClientes.frx":1C64
      PICN            =   "frmClientes.frx":1C80
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelcl 
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Exportar Clientes"
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
      MICON           =   "frmClientes.frx":221A
      PICN            =   "frmClientes.frx":2236
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelclI 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Importar Clientes"
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
      MICON           =   "frmClientes.frx":7A0C
      PICN            =   "frmClientes.frx":7A28
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView LVCM 
      Height          =   5535
      Left            =   1440
      TabIndex        =   9
      Top             =   6000
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CUIT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Apellido"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Domicilio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "E-Mail"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Telefono"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Condicion IVA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "DNI"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Fecha de alta"
         Object.Width           =   2540
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelcmlI 
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   9960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Importar Clientes Mayoristas"
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
      MICON           =   "frmClientes.frx":D1FE
      PICN            =   "frmClientes.frx":D21A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelcml 
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   8880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Exportar Clientes Mayoristas"
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
      MICON           =   "frmClientes.frx":129F0
      PICN            =   "frmClientes.frx":12A0C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpcionesM 
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Nuevo"
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
      MICON           =   "frmClientes.frx":181E2
      PICN            =   "frmClientes.frx":181FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpcionesM 
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmClientes.frx":185B2
      PICN            =   "frmClientes.frx":185CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpcionesM 
      Height          =   615
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Eliminar"
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
      MICON           =   "frmClientes.frx":18982
      PICN            =   "frmClientes.frx":1899E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpcionesM 
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   8160
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
      MICON           =   "frmClientes.frx":18D50
      PICN            =   "frmClientes.frx":18D6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdExcelcl_Click(Index As Integer)
'EXPORT to EXCEL
Dim xlApp As Excel.Application
Dim xlSh As Excel.Worksheet
Dim i As Long
Dim p As Long
Set xlApp = New Excel.Application

xlApp.Visible = True
xlApp.Workbooks.Add
Set xlSh = xlApp.Workbooks(1).Worksheets(1)

  
 For i = 1 To LV.ListItems.Count
 For p = 1 To LV.ListItems(i).ListSubItems.Count
 If (i = 1) Then
 xlSh.Cells(i, p).Font.Bold = True
   xlSh.Cells(i, p).Value = LV.ColumnHeaders(p).Text
   If (p = LV.ListItems(i).ListSubItems.Count) Then
    xlSh.Cells(i, p + 1).Font.Bold = True
    xlSh.Cells(i, p + 1).Value = LV.ColumnHeaders(p + 1).Text
    End If
   ' xlSh.Cells(i, p).Value = LV3.ListItems.Item(p)
End If
    If (p = 1) Then
    xlSh.Cells(i + 1, p).Value = LV.ListItems(i).Text
    End If
    xlSh.Cells(i + 1, p + 1).Value = LV.ListItems(i).ListSubItems(p).Text
       
          Next p
        Next i
 

Set xlSh = Nothing
Set xlApp = Nothing

End Sub

Private Sub cmdExcelclI_Click(Index As Integer)
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
 sql = "INSERT INTO ClientesM (CUIT,RS,Nombre,Apellido,Domicilio,Email,Telefono,RI,DNI,FechaDeAlta) VALUES ('"

 Next
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing





    rs.Open "select * from ClientesM", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView8(LVCM, rs)
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

Private Sub cmdExcelcml_Click(Index As Integer)
'EXPORT to EXCEL
Dim xlApp As Excel.Application
Dim xlSh As Excel.Worksheet
Dim i As Long
Dim p As Long
Set xlApp = New Excel.Application

xlApp.Visible = True
xlApp.Workbooks.Add
Set xlSh = xlApp.Workbooks(1).Worksheets(1)

  
 For i = 1 To LVCM.ListItems.Count
 For p = 1 To LVCM.ListItems(i).ListSubItems.Count
 If (i = 1) Then
 xlSh.Cells(i, p).Font.Bold = True
   xlSh.Cells(i, p).Value = LVCM.ColumnHeaders(p).Text
   If (p = LVCM.ListItems(i).ListSubItems.Count) Then
    xlSh.Cells(i, p + 1).Font.Bold = True
    xlSh.Cells(i, p + 1).Value = LVCM.ColumnHeaders(p + 1).Text
    End If
   ' xlSh.Cells(i, p).Value = LV3.ListItems.Item(p)
End If
    If (p = 1) Then
    xlSh.Cells(i + 1, p).Value = LVCM.ListItems(i).Text
    End If
    xlSh.Cells(i + 1, p + 1).Value = LVCM.ListItems(i).ListSubItems(p).Text
       
          Next p
        Next i
 

Set xlSh = Nothing
Set xlApp = Nothing


End Sub

Private Sub cmdExcelcmlI_Click(Index As Integer)
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
 sql = "INSERT INTO ClientesM (CUIT,RS,Nombre,Apellido,Domicilio,Email,Telefono,RI,DNI,FechaDeAlta) VALUES ('"

 Next
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing





    rs.Open "select * from Clientes", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView4(LV, rs)
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

' Botones de opción
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOpciones_Click(Index As Integer)
    Select Case Index
        Case 0: Call Agregar
        Case 1: Call Editar
        Case 2: Call Eliminar
        Case 3: Unload Me
        Case 4: frmFilterC.Show , Me
        Case 5: Call mnuImprimir_Click
    End Select
End Sub


'Abre el formulario para Editar el registro seleccionado en el ListView
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editar()

    Dim i As Integer
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LV.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
    With frmEditC
        ' obtiene el elemento seleccionado
        .lblID = LV.SelectedItem.Text
        
            .Text1(1).Text = LV.SelectedItem.ListSubItems(1).Text
            .Text1(2).Text = LV.SelectedItem.ListSubItems(2).Text
             .Text1(0).Text = LV.SelectedItem.ListSubItems(3).Text
            .Text1(3).Text = LV.SelectedItem.ListSubItems(4).Text
       
         .lblFecha = LV.SelectedItem.ListSubItems(5).Text
        .IdRegistro = LV.SelectedItem.Text
        .ACCION = EDITAR_REGISTRO1
        
        .Show vbModal
    End With

End Sub

' Elimina el registro actual seleccionado
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub Eliminar()

    

    If (LV.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LV.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "Nombre " & .ListSubItems(1).Text & vbNewLine & _
                 "Apellido: " & .ListSubItems(2).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Clientes where Id = " & .Text & ""
            ' refresca el recordset
            rs.Open "select * from Clientes", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView4(LV, rs)
        End If
    End With
    If rs.State = adStateOpen Then rs.Close
End Sub


' Elimina el registro actual seleccionado
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub EliminarM()

    

    If (LVCM.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LVCM.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LVCM.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "CUIT: " & .ListSubItems(1).Text & vbNewLine & _
                 "Razon Social: " & .ListSubItems(2).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from ClientesM where Id = " & .Text & ""
            ' refresca el recordset
            rs.Open "select * from ClientesM", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView8(LVCM, rs)
        End If
    End With
    If rs.State = adStateOpen Then rs.Close
End Sub




Sub Agregar()
    
    ' Acción
    frmEditC.ACCION = AGREGAR_REGISTRO1
    
    frmEditC.lblFecha = Format(Date, "mm/dd/yyyy")
    ' Abre el Form
    frmEditC.Show 1
End Sub
Sub AgregarM()
    
    ' Acción
    frmEditCM.ACCION = AGREGAR_REGISTRO3
    frmEditCM.cmbRI.ListIndex = 0
    frmEditCM.lblFecha = Format(Date, "mm/dd/yyyy")
    ' Abre el Form
    frmEditCM.Show 1
End Sub
'Abre el formulario para Editar el registro seleccionado en el ListView
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditarM()

    Dim i As Integer
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LVCM.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LVCM.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
    If LVCM.SelectedItem.ListSubItems(8).Text = "Si" Then
frmEditCM.cmbRI.ListIndex = 1
Else
frmEditCM.cmbRI.ListIndex = 0
End If
    
    
    With frmEditCM
        ' obtiene el elemento seleccionado
        .lblID = LVCM.SelectedItem.Text
        
            .Text1(1).Text = LVCM.SelectedItem.ListSubItems(1).Text
            .Text1(2).Text = LVCM.SelectedItem.ListSubItems(2).Text
             .Text1(3).Text = LVCM.SelectedItem.ListSubItems(3).Text
            .Text1(0).Text = LVCM.SelectedItem.ListSubItems(4).Text
           .Text1(7).Text = LVCM.SelectedItem.ListSubItems(5).Text
          .Text1(6).Text = LVCM.SelectedItem.ListSubItems(6).Text
          .Text1(5).Text = LVCM.SelectedItem.ListSubItems(7).Text
          .Text1(4).Text = LVCM.SelectedItem.ListSubItems(9).Text
          
         .lblFecha = LVCM.SelectedItem.ListSubItems(10).Text
        .IdRegistro = LVCM.SelectedItem.Text
        .ACCION = EDITAR_REGISTRO3
        
        .Show vbModal
    End With


End Sub


Sub Salir()
    If rs.State = adStateOpen Then rs.Close
    Unload Me
    End
End Sub


Private Sub cmdOpcionesM_Click(Index As Integer)
   Select Case Index
        Case 6: Call AgregarM
        Case 7: Call EditarM
        Case 8: Call EliminarM
        Case 9: frmFilterCM.Show , Me
  '      Case 5: Call mnuImprimir_Click
    End Select
End Sub

Private Sub Form_Load()
If Cli(0) = 0 Then
cmdOpciones(0).Visible = False
cmdOpcionesM(6).Visible = False
cmdExcelclI(0).Visible = False
cmdExcelcmlI(0).Visible = False
End If
If Cli(1) = 0 Then
cmdOpciones(1).Visible = False
cmdOpcionesM(7).Visible = False
End If
If Cli(2) = 0 Then
cmdOpciones(2).Visible = False
cmdOpcionesM(8).Visible = False
End If
frmMain.Enabled = False
    ' llena el ListView
   
      rs.Open "select * from Clientes", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView4(LV, rs)
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT * from ClientesM", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView8(LVCM, rs)
    If rs.State = adStateOpen Then rs.Close

End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
    If rs.State = adStateOpen Then rs.Close
End Sub

Private Sub LV_DblClick()
If Cli(1) = 0 Then
Exit Sub
End If
    Call Editar
End Sub



Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Item As ListItem
    
    Set Item = LV.HitTest(X, Y)
    
    If Not Item Is Nothing And Button = vbRightButton Then
       Item.Selected = True
     End If
End Sub

' menues
'''''''''''''''''''''''''''''

Private Sub mnuAgregar_Click()
    Call Agregar
End Sub

Private Sub mnuEditarRegistro_Click()
    Call Editar
End Sub

Private Sub mnuEliminarReg_Click()
    Call Eliminar
End Sub

Private Sub mnuImprimir_Click()
    If rs.State = adStateOpen Then rs.Close
   rs.Open "select * from Clientes", cnn, adOpenStatic, adLockOptimistic
    Set DataReport4.DataSource = rs
    DataReport4.Show 1
       If rs.State = adStateOpen Then rs.Close
End Sub

' salir

''''''''''''''''''''''''
Private Sub mnuSalir_Click()
   Unload Me
End Sub

Private Sub LVCM_DblClick()
If Cli(1) = 0 Then
Exit Sub
End If
Call EditarM
End Sub
