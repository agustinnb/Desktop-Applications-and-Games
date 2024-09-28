VERSION 5.00
Begin VB.Form frmAddVis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar vista"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8940
   LinkTopic       =   "Agregar vista"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkStock 
      Caption         =   "Imprimir Stock"
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   20
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Exportar Stock"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   19
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Importar Stock"
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   18
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Anotaciones"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   17
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Fecha de alta"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   16
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "IVA"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Precio unitario"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   14
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Cantidad"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   13
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Distribuidor"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Modelo"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Marca"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Producto"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Eliminar producto"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Editar producto"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Agregar producto"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox chkConf 
      Caption         =   "Configurar vistas"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CheckBox chkConf 
      Caption         =   "Precio dolar"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CheckBox chkConf 
      Caption         =   "IVA"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox chkConf 
      Caption         =   "Nombre del local"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CheckBox chkConf 
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmAddVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
