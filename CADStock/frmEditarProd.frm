VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmEditarProd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4995
   Icon            =   "frmEditarProd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtProds 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1680
      TabIndex        =   39
      Top             =   7560
      Width           =   855
   End
   Begin VB.ComboBox cmbDep 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditarProd.frx":058A
      Left            =   1680
      List            =   "frmEditarProd.frx":058C
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   7200
      Width           =   2415
   End
   Begin VB.ComboBox cmbCod 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditarProd.frx":058E
      Left            =   1680
      List            =   "frmEditarProd.frx":0590
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox cmbProd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditarProd.frx":0592
      Left            =   1680
      List            =   "frmEditarProd.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ComboBox cmbD 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditarProd.frx":0596
      Left            =   1680
      List            =   "frmEditarProd.frx":0598
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox CmbMon 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditarProd.frx":059A
      Left            =   2880
      List            =   "frmEditarProd.frx":05A4
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtProds 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox txtFecha 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   20
      Top             =   240
      Width           =   2415
   End
   Begin VB.ComboBox CmbIVA 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEditarProd.frx":05B0
      Left            =   1680
      List            =   "frmEditarProd.frx":05B2
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox txtProds 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   16
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txtProds 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   15
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtProds 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   3
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox txtProds 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtProds 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtProds 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin ChamaleonButton.ChameleonBtn cmdCancelar 
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   9120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmEditarProd.frx":05B4
      PICN            =   "frmEditarProd.frx":05D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdGuardar 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   9120
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
      MICON           =   "frmEditarProd.frx":0B6A
      PICN            =   "frmEditarProd.frx":0B86
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   40
      Top             =   7560
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   37
      Top             =   7200
      Width           =   630
   End
   Begin VB.Label lblModelo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   36
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblMarca 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   35
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   34
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distribuidor"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   31
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      Caption         =   "lblError"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   8160
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   27
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblPFU 
      Caption         =   "0"
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
      Left            =   3360
      TabIndex        =   26
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label lblGPU 
      Caption         =   "0"
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
      Left            =   3360
      TabIndex        =   25
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Precio Final por unidad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   24
      Top             =   8640
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Ganancia por unidad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anotaciones"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   21
      Top             =   6240
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   18
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Margen"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   17
      Top             =   5280
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Height          =   9495
      Left            =   120
      Top             =   240
      Width           =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   240
      X2              =   4080
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de alta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1680
      TabIndex        =   14
      Top             =   240
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de alta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   360
      X2              =   4440
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Unitario"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   3240
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id de registro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   1665
   End
End
Attribute VB_Name = "frmEditarProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public IV1, IV2, IV3, IV4, IV5 As Single
Dim FC As Boolean
Dim FP As Boolean
Dim depo As String



Private Sub cmbCod_Click()
If FC = False Then
FC = True
Exit Sub
Else
If rs.State = adStateOpen Then rs.Close
cmbProd.ListIndex = cmbCod.ListIndex
If cmbCod.Text <> "Otro" Then
rs.Open "SELECT * FROM ProdXDist WHERE Codigo = '" & cmbCod.Text & "'", cnn, adOpenStatic, adLockOptimistic
lblMarca.Caption = rs!marca
lblModelo.Caption = rs!modelo
txtProds(1).Text = rs!producto
txtProds(2).Text = rs!marca
txtProds(3).Text = rs!modelo
txtProds(0).Text = rs!preciou
cmbMon.ListIndex = rs!Moneda - 1
 txtProds(5).Text = rs!Margen
txtProds(1).Visible = False
txtProds(2).Visible = False
txtProds(3).Visible = False
rs.Close
Else
txtProds(0).Text = ""
lblMarca.Caption = ""
lblModelo.Caption = ""
txtProds(1).Text = ""
txtProds(2).Text = ""
txtProds(3).Text = ""
 txtProds(5).Text = ""
txtProds(1).Visible = True
txtProds(2).Visible = True
txtProds(3).Visible = True
End If
End If
End Sub

Private Sub cmbD_Click()
Dim getid As Integer
rs.Open "SELECT Id FROM Distribuidores WHERE Nombre = '" & cmbD.Text & "'", cnn, adOpenStatic, adLockOptimistic
getid = rs!id
rs.Close
rs.Open "SELECT * FROM ProdXDist WHERE idProv = '" & getid & "'", cnn, adOpenStatic, adLockOptimistic
cmbCod.Clear
cmbProd.Clear
While Not rs.EOF
cmbCod.AddItem rs!codigo
cmbProd.AddItem rs!producto
rs.MoveNext
Wend
rs.Close
cmbCod.AddItem "Otro"
cmbProd.AddItem "Otro"
cmbCod.ListIndex = 0
cmbProd.ListIndex = 0
End Sub

Private Sub CmbIVA_Click()
If txtProds(0).Text <> "" And txtProds(5).Text <> "" And cmbIVA.Text <> "" Then
lblGPU.Caption = (txtProds(0) * txtProds(5).Text) / 100
Dim ivtemp As Single
If (cmbIVA.ListIndex = 0) Then
ivtemp = 0
ElseIf (cmbIVA.ListIndex = 1) Then
ivtemp = IV1
ElseIf (cmbIVA.ListIndex = 2) Then
ivtemp = IV2
ElseIf (cmbIVA.ListIndex = 3) Then
ivtemp = IV3
ElseIf (cmbIVA.ListIndex = 4) Then
ivtemp = IV4
ElseIf (cmbIVA.ListIndex = 5) Then
ivtemp = IV5
End If

lblPFU.Caption = txtProds(0) + ((txtProds(0) * txtProds(5).Text) / 100) + (((txtProds(0) + ((txtProds(0) * txtProds(5).Text) / 100)) * ivtemp) / 100)


End If

End Sub

Private Sub cmbProd_Click()
If FP = False Then
FP = True
Exit Sub
Else
cmbCod.ListIndex = cmbProd.ListIndex
If cmbProd.Text <> "Otro" Then
rs.Open "SELECT * FROM ProdXDist WHERE Codigo = '" & cmbCod.Text & "'", cnn, adOpenStatic, adLockOptimistic
lblMarca.Caption = rs!marca
lblModelo.Caption = rs!modelo
txtProds(1).Text = rs!producto
txtProds(2).Text = rs!marca
txtProds(3).Text = rs!modelo
txtProds(0).Text = rs!preciou
cmbMon.ListIndex = rs!Moneda - 1
 txtProds(5).Text = rs!Margen
txtProds(1).Visible = False
txtProds(2).Visible = False
txtProds(3).Visible = False
rs.Close
Else
lblMarca.Caption = ""
lblModelo.Caption = ""
txtProds(1).Text = ""
txtProds(2).Text = ""
txtProds(3).Text = ""
 txtProds(5).Text = ""
txtProds(1).Visible = True
txtProds(2).Visible = True
txtProds(3).Visible = True
End If
End If


End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdGuardar_Click()
Dim a As Boolean
Dim precio As Single
a = True


If (Trim(txtProds(1).Text = "")) Then
a = False
lblError.Caption = "Tiene que poner un nombre al producto"
Exit Sub
End If

If (txtProds(4).Text = "") Or Not IsNumeric(Trim(txtProds(4).Text)) Then
a = False
lblError.Caption = "La cantidad tiene que ser numerica"
Exit Sub
End If




If (txtProds(0).Text = "") Or Not IsNumeric(Trim(txtProds(0).Text)) Then
a = False
lblError.Caption = "El Precio Unitario tiene que ser numerico"
Exit Sub
End If
If Not IsNumeric(Trim(txtProds(7).Text)) Then
If (txtProds(7).Text <> "") Then
a = False
lblError.Caption = "La cantidad a pasar a otro deposito tiene que ser numerica"
Exit Sub
Else
txtProds(7).Text = 0
End If
End If

If (val(txtProds(4).Text) < val(txtProds(7).Text)) Then
a = False
lblError.Caption = "La cantidad tiene que ser menor a la cantidad a pasar al deposito"
Exit Sub
End If

If (txtProds(5).Text = "") Or Not IsNumeric(Trim(txtProds(5).Text)) Then
a = False
lblError.Caption = "El Margen tiene que ser numerico"
Exit Sub
End If

If cmbMon.ListIndex = 1 Then
precio = Format(txtProds(0) * pdolar, "#,##0.00")
Else
precio = Format(txtProds(0), "#,##0.00")
End If

If a = True Then
lblError.Caption = ""

If (EditarProductos = 1) Then
Dim canti As Integer
Dim canti2 As Integer
If (cmbProd.Text <> "Otro") Then

If (txtProds(7).Text <> "") And (val(txtProds(7).Text) < val(txtProds(4).Text)) Then

canti = val(txtProds(4).Text)
canti2 = val(txtProds(7).Text)
cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,Deposito,Distribuidor,IVA,Anotaciones,Margen) VALUES('" & _
                                 txtProds(1) & "','" & _
                                txtProds(2) & "','" & _
                                txtProds(3) & "','" & _
                               txtProds(7) & "','" & _
                               precio & "','" & _
                               txtFecha(6) & "','" & _
                               cmbDep.Text & "','" & _
                               cmbD.Text & "','" & _
                                (cmbIVA.ListIndex) & "','" & _
                                  txtProds(6) & "','" & _
                                 txtProds(5) & "')"
cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,Deposito,Distribuidor,IVA,Anotaciones,Margen) VALUES('" & _
                                 txtProds(1) & "','" & _
                                txtProds(2) & "','" & _
                                txtProds(3) & "','" & _
                               (canti - canti2) & "','" & _
                               precio & "','" & _
                               txtFecha(6) & "','" & _
                               "No definido" & "','" & _
                               cmbD.Text & "','" & _
                                (cmbIVA.ListIndex) & "','" & _
                                  txtProds(6) & "','" & _
                                 txtProds(5) & "')"

Else
cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,Deposito,Distribuidor,IVA,Anotaciones,Margen) VALUES('" & _
                                 txtProds(1) & "','" & _
                                txtProds(2) & "','" & _
                                txtProds(3) & "','" & _
                               txtProds(4) & "','" & _
                               precio & "','" & _
                               txtFecha(6) & "','" & _
                               cmbDep.Text & "','" & _
                               cmbD.Text & "','" & _
                                (cmbIVA.ListIndex) & "','" & _
                                  txtProds(6) & "','" & _
                                 txtProds(5) & "')"

End If

Else
If (txtProds(7).Text <> "") And (val(txtProds(7).Text) < val(txtProds(4).Text)) Then


canti = val(txtProds(4).Text)
canti2 = val(txtProds(7).Text)

   cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,Deposito,IVA,Anotaciones,Margen) VALUES('" & _
                                 txtProds(1) & "','" & _
                                txtProds(2) & "','" & _
                                txtProds(3) & "','" & _
                               txtProds(7) & "','" & _
                               precio & "','" & _
                               txtFecha(6) & "','" & _
                                  cmbDep.Text & "','" & _
                                (cmbIVA.ListIndex) & "','" & _
                                  txtProds(6) & "','" & _
                                 txtProds(5) & "')"
       cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,Deposito,IVA,Anotaciones,Margen) VALUES('" & _
                                 txtProds(1) & "','" & _
                                txtProds(2) & "','" & _
                                txtProds(3) & "','" & _
                               (canti - canti2) & "','" & _
                               precio & "','" & _
                               txtFecha(6) & "','" & _
                                  "No definido" & "','" & _
                                (cmbIVA.ListIndex) & "','" & _
                                  txtProds(6) & "','" & _
                                 txtProds(5) & "')"

Else
   cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,Deposito,IVA,Anotaciones,Margen) VALUES('" & _
                                 txtProds(1) & "','" & _
                                txtProds(2) & "','" & _
                                txtProds(3) & "','" & _
                               txtProds(4) & "','" & _
                               precio & "','" & _
                               txtFecha(6) & "','" & _
                                  cmbDep.Text & "','" & _
                                (cmbIVA.ListIndex) & "','" & _
                                  txtProds(6) & "','" & _
                                 txtProds(5) & "')"

End If


End If
If (cmbDep.Text <> "No definido") Then
If (txtProds(7).Text <> "") And (val(txtProds(7).Text) < val(txtProds(4).Text)) Then
cnn.Execute "INSERT INTO MovDep (Desde, Producto,Cantidad,Hasta,Fecha) VALUES (' ','" & _
            txtProds(1) & " " & txtProds(2) & " " & txtProds(3) & "','" & _
            txtProds(7) & "','" & _
            cmbDep.Text & "','" & _
            txtFecha(6) & "')"
            Else
            cnn.Execute "INSERT INTO MovDep (Desde, Producto,Cantidad,Hasta,Fecha) VALUES (' ','" & _
            txtProds(1) & " " & txtProds(2) & " " & txtProds(3) & "','" & _
            txtProds(4) & "','" & _
            cmbDep.Text & "','" & _
            txtFecha(6) & "')"
            
End If

            


End If

    rs.Open "select * from Productos", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView2(frmMain.LV2, rs)
     



Unload Me


End If

If (EditarProductos = 2) Then
Dim depoact As String
Dim sql As String
txtFecha(6).Text = lblFecha.Caption
If (cmbCod.Text <> "Otro") Then

If (txtProds(7).Text <> "") And (val(txtProds(7).Text) < val(txtProds(4).Text)) Then
rs.Open "SELECT * FROM Productos WHERE id=" & ProdId
If Len(rs!deposito) <> 0 Then
depoact = rs!deposito
Else
depoact = "No definido"
End If
rs.Close
If depoact <> cmbDep.Text Then

cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,Deposito,Distribuidor,IVA,Anotaciones,Margen) VALUES('" & _
                                 txtProds(1) & "','" & _
                                txtProds(2) & "','" & _
                                txtProds(3) & "','" & _
                               val(txtProds(7)) & "','" & _
                               precio & "','" & _
                               txtFecha(6) & "','" & _
                               cmbDep.Text & "','" & _
                               cmbD.Text & "','" & _
                                (cmbIVA.ListIndex) & "','" & _
                                  txtProds(6) & "','" & _
                                 txtProds(5) & "')"

  sql = "UPDATE Productos set Producto = '" & txtProds(1) & _
   "', Marca = '" & txtProds(2) & _
   "', Modelo = '" & txtProds(3) & _
   "', Cantidad = '" & (val(txtProds(4)) - val(txtProds(7))) & _
    "', Distribuidor = '" & cmbD.Text & _
   "', PrecioU = '" & precio & _
   "', FechaDeAlta = '" & lblFecha.Caption & _
   "', IVA = '" & (cmbIVA.ListIndex) & _
   "', Anotaciones = '" & txtProds(6) & _
  "', Margen = '" & txtProds(5) & "'"
   sql = sql & ", deposito='" & depoact & "'"
 sql = sql & " where id = " & ProdId
  cnn.Execute sql
Else
  sql = "UPDATE Productos set Producto = '" & txtProds(1) & _
   "', Marca = '" & txtProds(2) & _
   "', Modelo = '" & txtProds(3) & _
   "', Cantidad = '" & txtProds(4) & _
    "', Distribuidor = '" & cmbD.Text & _
   "', PrecioU = '" & precio & _
   "', FechaDeAlta = '" & lblFecha.Caption & _
   "', IVA = '" & (cmbIVA.ListIndex) & _
   "', Anotaciones = '" & txtProds(6) & _
   "', Margen = '" & txtProds(5) & "'"
   If cmbDep.Text <> "No definido" Then
   sql = sql & ", deposito='" & cmbDep.Text & "'"
   End If
   sql = sql & " where id = " & ProdId
   cnn.Execute sql
End If
Else
  sql = "UPDATE Productos set Producto = '" & txtProds(1) & _
   "', Marca = '" & txtProds(2) & _
   "', Modelo = '" & txtProds(3) & _
   "', Cantidad = '" & txtProds(4) & _
    "', Distribuidor = '" & cmbD.Text & _
   "', PrecioU = '" & precio & _
   "', FechaDeAlta = '" & lblFecha.Caption & _
   "', IVA = '" & (cmbIVA.ListIndex) & _
   "', Anotaciones = '" & txtProds(6) & _
   "', Margen = '" & txtProds(5) & "'"
   If cmbDep.Text <> "No definido" Then
   sql = sql & ", deposito='" & cmbDep.Text & "'"
   End If
   sql = sql & " where id = " & ProdId
   cnn.Execute sql


End If



Else
  
If (txtProds(7).Text <> "") And (val(txtProds(7).Text) < val(txtProds(4).Text)) Then
rs.Open "SELECT * FROM Productos WHERE id=" & ProdId
depoact = rs!deposito
rs.Close
If depoact <> cmbDep.Text Then

   cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,Deposito,IVA,Anotaciones,Margen) VALUES('" & _
                                 txtProds(1) & "','" & _
                                txtProds(2) & "','" & _
                                txtProds(3) & "','" & _
                               txtProds(7) & "','" & _
                               precio & "','" & _
                               txtFecha(6) & "','" & _
                                  cmbDep.Text & "','" & _
                                (cmbIVA.ListIndex) & "','" & _
                                  txtProds(6) & "','" & _
                                 txtProds(5) & "')"

   sql = "UPDATE Productos set Producto = '" & txtProds(1) & _
   "', Marca = '" & txtProds(2) & _
   "', Modelo = '" & txtProds(3) & _
   "', Cantidad = '" & (val(txtProds(4)) - val(txtProds(7))) & _
   "', PrecioU = '" & precio & _
   "', Distribuidor = '" & _
   "', FechaDeAlta = '" & lblFecha.Caption & _
   "', IVA = '" & (cmbIVA.ListIndex) & _
   "', Deposito = '" & cmbDep.Text & _
   "', Anotaciones = '" & txtProds(6) & _
   "', Margen = '" & txtProds(5) & "'"
   If cmbDep.Text <> "No definido" Then
   sql = sql & ", deposito='" & cmbDep.Text & "'"
   End If
   sql = sql & " where id = " & ProdId
  cnn.Execute sql
Else
   sql = "UPDATE Productos set Producto = '" & txtProds(1) & _
   "', Marca = '" & txtProds(2) & _
   "', Modelo = '" & txtProds(3) & _
   "', Cantidad = '" & txtProds(4) & _
   "', PrecioU = '" & precio & _
   "', Distribuidor = '" & _
   "', FechaDeAlta = '" & lblFecha.Caption & _
   "', IVA = '" & (cmbIVA.ListIndex) & _
   "', Deposito = '" & cmbDep.Text & _
   "', Anotaciones = '" & txtProds(6) & _
   "', Margen = '" & txtProds(5) & "'"
   If cmbDep.Text <> "No definido" Then
   sql = sql & ", deposito='" & cmbDep.Text & "'"
   End If
   sql = sql & " where id = " & ProdId
  cnn.Execute sql
End If
Else
   sql = "UPDATE Productos set Producto = '" & txtProds(1) & _
   "', Marca = '" & txtProds(2) & _
   "', Modelo = '" & txtProds(3) & _
   "', Cantidad = '" & txtProds(4) & _
   "', PrecioU = '" & precio & _
   "', Distribuidor = '" & _
   "', FechaDeAlta = '" & lblFecha.Caption & _
   "', IVA = '" & (cmbIVA.ListIndex) & _
   "', Deposito = '" & cmbDep.Text & _
   "', Anotaciones = '" & txtProds(6) & _
   "', Margen = '" & txtProds(5) & "'"
   If cmbDep.Text <> "No definido" Then
   sql = sql & ", deposito='" & cmbDep.Text & "'"
   End If
   sql = sql & " where id = " & ProdId
  cnn.Execute sql


End If
  
  
End If
If (cmbDep.Text <> "No definido") Then
If (depo <> cmbDep.Text) Then
If (depo <> "") Then
If (txtProds(7).Text <> "") And (val(txtProds(7).Text) < val(txtProds(4).Text)) Then
cnn.Execute "INSERT INTO MovDep (Desde,Producto,Cantidad,Hasta,Fecha) VALUES ('" & _
            depo & "','" & _
            txtProds(1) & " " & txtProds(2) & " " & txtProds(3) & "','" & _
            txtProds(7) & "','" & _
            cmbDep.Text & "','" & _
            Format(Date, "dd/mm/yyyy") & "')"
cnn.Execute "INSERT INTO MovDep (Desde,Producto,Cantidad,Hasta,Fecha) VALUES ('" & _
            depo & "','" & _
            txtProds(1) & " " & txtProds(2) & " " & txtProds(3) & "','" & _
            (val(txtProds(4)) - val(txtProds(7))) & "','" & _
            depoact & "','" & _
            Format(Date, "dd/mm/yyyy") & "')"

Else
cnn.Execute "INSERT INTO MovDep (Desde,Producto,Cantidad,Hasta,Fecha) VALUES ('" & _
            depo & "','" & _
            txtProds(1) & " " & txtProds(2) & " " & txtProds(3) & "','" & _
            txtProds(4) & "','" & _
            cmbDep.Text & "','" & _
            Format(Date, "dd/mm/yyyy") & "')"

End If
Else
depo = "No definido"
cnn.Execute "INSERT INTO MovDep (Desde,Producto,Cantidad,Hasta,Fecha) VALUES ('" & _
            depo & "','" & _
            txtProds(1) & " " & txtProds(2) & " " & txtProds(3) & "','" & _
            txtProds(4) & "','" & _
            cmbDep.Text & "','" & _
            Format(Date, "dd/mm/yyyy") & "')"
End If
Else

If (txtProds(7).Text <> "") And (val(txtProds(7).Text) < val(txtProds(4).Text)) Then
cnn.Execute "INSERT INTO MovDep (Desde,Producto,Cantidad,Hasta,Fecha) VALUES ('" & _
            " ','" & _
            txtProds(1) & " " & txtProds(2) & " " & txtProds(3) & "','" & _
            txtProds(7) & "','" & _
            cmbDep.Text & "','" & _
            Format(Date, "dd/mm/yyyy") & "')"

Else
cnn.Execute "INSERT INTO MovDep (Desde,Producto,Cantidad,Hasta,Fecha) VALUES ('" & _
            " ','" & _
            txtProds(1) & " " & txtProds(2) & " " & txtProds(3) & "','" & _
            txtProds(4) & "','" & _
            cmbDep.Text & "','" & _
            Format(Date, "dd/mm/yyyy") & "')"

End If

        
End If


End If


'cnn.Execute "UPDATE Personas set Nombre = '" & Text1(1) & _
 '                                        "', Apellido = '" & Text1(2) & _
  '                                       "', Telefono = '" & Text1(3) & _
   '                                      "', Direccion = '" & Text1(4) & _
    '                                     "', Sexo = '" & CmbSexo.ListIndex & _
     '                                    "' where Id = " & IdRegistro & ""



    rs.Open "select * from Productos", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView2(frmMain.LV2, rs)
    


Unload Me


End If


End If
End Sub

Private Sub Form_Load()
depo = ""
FC = False
FP = False
cmbMon.ListIndex = 0
  Dim first As Boolean
  Dim firstnom As String
  first = False
  If rs.State = adStateOpen Then rs.Close
rs.Open "select * from Distribuidores", cnn, adOpenStatic, adLockOptimistic
       While Not rs.EOF
     cmbD.AddItem rs!Nombre
     If (first = False) Then
     firstnom = rs!id
     first = True
     End If
     rs.MoveNext
     Wend
     rs.Close
     first = False
     rs.Open "select * from ProdXDist WHERE idProv = '" & firstnom & "'", cnn, adOpenStatic, adLockOptimistic
      While Not rs.EOF
        cmbCod.AddItem rs!codigo
        cmbProd.AddItem rs!producto
        If (first = False) Then
        txtProds(0).Text = rs!preciou
        cmbMon.ListIndex = rs!Moneda - 1
        txtProds(5).Text = rs!Margen
        lblMarca.Caption = rs!marca
lblModelo.Caption = rs!modelo
txtProds(1).Text = rs!producto
txtProds(2).Text = rs!marca
txtProds(3).Text = rs!modelo
txtProds(0).Text = rs!preciou
cmbMon.ListIndex = rs!Moneda - 1
 txtProds(5).Text = rs!Margen
         first = True
         End If
         rs.MoveNext
     Wend
     
     rs.Close
      cmbCod.AddItem "Otro"
     cmbProd.AddItem "Otro"
  
 If cmbD.ListCount > 0 Then
cmbD.ListIndex = 0
End If
first = False
If (EditarProductos = 1) Then

   
    cmbCod.ListIndex = 0
    cmbProd.ListIndex = 0




If (cmbProd.Text <> "Otro") Then
txtProds(1).Visible = False
txtProds(2).Visible = False
txtProds(3).Visible = False
rs.Open "select * from ProdXDist WHERE Codigo = '" & cmbCod.Text & "'", cnn, adOpenStatic, adLockOptimistic
lblMarca.Caption = rs!marca
lblModelo.Caption = rs!modelo
rs.Close
End If



  If rs.State = adStateOpen Then rs.Close
lblError.Caption = ""
frmEditarProd.Caption = "Agregar Producto"
lblFecha.Visible = False
txtFecha(6).Visible = True
Label1(0).Visible = False
txtFecha(6).Text = Format(Date, "dd/mm/yyyy")
cmdGuardar.Caption = "Guardar producto"
 
  cmbDep.AddItem "No definido"
       
         If rs.State = adStateOpen Then rs.Close
rs.Open "select * from Depositos", cnn, adOpenStatic, adLockOptimistic
       While Not rs.EOF
     cmbDep.AddItem rs!Nombre
     rs.MoveNext
     Wend
     rs.Close
 cmbDep.ListIndex = 0
 
    ' carga el Recorset con todos los datos
    rs.Open "select * from Configuracion", cnn, adOpenStatic, adLockOptimistic
     On Error GoTo ErrorSub
    cmbIVA.AddItem ("No")
    cmbIVA.AddItem (rs!IVA1 & " %")
    cmbIVA.AddItem (rs!IVA2 & " %")
    IV1 = rs!IVA1
    IV2 = rs!IVA2
    If (rs!IVA3 <> " ") Then
    IV3 = rs!IVA3
    cmbIVA.AddItem (rs!IVA3 & " %")
    End If
     If (rs!IVA4 <> " ") Then
    IV4 = rs!IVA4
    cmbIVA.AddItem (rs!IVA4 & " %")
    End If
     If (rs!IVA5 <> " ") Then
    IV5 = rs!IVA5
    cmbIVA.AddItem (rs!IVA5 & " %")
    End If
  
      rs.Close
       cmbIVA.ListIndex = 0
      
       
      Exit Sub
    
    
    
    
ErrorSub:
    rs.Close
    MsgBox ("Primero tiene que configurar el sistema")
    
      If rs.State = adStateOpen Then rs.Close
   

    If Err.Number = 94 Then Resume Next
Unload Me

End If

If (EditarProductos = 2) Then
  If rs.State = adStateOpen Then rs.Close
      ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (frmMain.LV2.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Unload Me
       Exit Sub
    End If
    If (frmMain.LV2.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Unload Me
       Exit Sub
    End If
    

lblError.Caption = ""
frmEditarProd.Caption = "Editar Producto"
lblFecha.Visible = True
txtFecha(6).Visible = False
Label1(0).Visible = True
cmdGuardar.Caption = "Guardar producto"




    ' carga el Recorset con todos los datos
        
        If rs.State = adStateOpen Then rs.Close
    rs.Open "select * from Configuracion", cnn, adOpenStatic, adLockOptimistic
   '  On Error GoTo ErrorSub2
      cmbIVA.AddItem ("No")
    cmbIVA.AddItem (rs!IVA1 & " %")
    cmbIVA.AddItem (rs!IVA2 & " %")
    IV1 = rs!IVA1
    IV2 = rs!IVA2
    If (rs!IVA3 <> " ") Then
    IV3 = rs!IVA3
    cmbIVA.AddItem (rs!IVA3 & " %")
    End If
     If (rs!IVA4 <> " ") Then
    IV4 = rs!IVA4
    cmbIVA.AddItem (rs!IVA4 & " %")
    End If
     If (rs!IVA5 <> " ") Then
    IV5 = rs!IVA5
    cmbIVA.AddItem (rs!IVA5 & " %")
    End If
      rs.Close
      
   
      
       rs.Open "select * from Productos where Id = " & ProdId, cnn, adOpenStatic, adLockOptimistic
  Dim a As Integer
  a = 0
  Dim producto As String, marca As String, modelo As String, cantidad As String, fechadealta As String, anotaciones As String, preciou As String, Margen As String, IVA As String, id As Integer
  producto = rs!producto
If Not IsNull(rs!deposito) Then
If Trim(rs!deposito) <> "" Then
  depo = rs!deposito
  End If
End If

  marca = rs!marca
  modelo = rs!modelo
  cantidad = rs!cantidad
  fechadealta = rs!fechadealta
   preciou = rs!preciou
  Margen = rs!Margen
  id = rs!id
  IVA = rs!IVA
  If (rs!anotaciones <> "") Then
  anotaciones = rs!anotaciones
  End If
  
  If rs.State = adStateOpen Then rs.Close
 
  
   txtProds(1) = producto
      txtProds(2) = marca
      txtProds(3) = modelo
      txtProds(4) = cantidad
      txtProds(0) = preciou
      lblFecha.Caption = fechadealta
       txtProds(6) = anotaciones
      
      txtProds(5) = Margen
      lblID.Caption = id
      cmbIVA.ListIndex = IVA
      
  While a <> cmbProd.ListCount - 1
  
  If (cmbProd.List(a) = producto) Then
  cmbCod.ListIndex = a
  cmbProd.ListIndex = a
  End If
  a = a + 1
  Wend
  
  If (cmbProd.Text = "") Then
  cmbProd.ListIndex = cmbProd.ListCount - 1
  cmbCod.ListIndex = cmbCod.ListCount - 1
  End If
 If (cmbProd.Text <> "Otro") Then
 txtProds(1).Visible = False
 txtProds(2).Visible = False
 txtProds(3).Visible = False
 lblMarca.Caption = marca
 lblModelo.Caption = modelo
 End If
     
     cmbDep.AddItem "No definido"
       Dim numer As Integer
       Dim gnom As String
       Dim gnum As Integer
       numer = 0
         If rs.State = adStateOpen Then rs.Close
         
rs.Open "select * from Depositos", cnn, adOpenStatic, adLockOptimistic
       While Not rs.EOF
       gnom = rs!Nombre
      numer = numer + 1
       If (gnom <> "") Then
       If (gnom = depo) Then
       gnum = numer
       End If
       End If
     cmbDep.AddItem rs!Nombre
     
     rs.MoveNext
     Wend
     rs.Close
    
 cmbDep.ListIndex = gnum
     
 
      
     '   cnn.Execute "INSERT INTO Productos " & "(Producto,Marca,Modelo,Cantidad,PrecioU,FechaDeAlta,IVA,Anotaciones,Margen) VALUES('" & _
      '                           txtProds(1) & "','" & _
       '                         txtProds(2) & "','" & _
        '                        txtProds(3) & "','" & _
         '                      txtProds(4) & "','" & _
          '                     txtProds(0) & "','" & _
           '                    txtFecha(6) & "','" & _
            '                    (CmbIVA.ListIndex + 1) & "','" & _
             '                     txtProds(6) & "','" & _
              '                   txtProds(5) & "')"
    
      Exit Sub
    
    
    
    
ErrorSub2:
    If rs.State = adStateOpen Then rs.Close
    MsgBox ("Primero tiene que configurar el sistema")
    
  


    If Err.Number = 94 Then Resume Next

Unload Me


















End If






End Sub


Private Sub Form_Unload(Cancel As Integer)
  If rs.State = adStateOpen Then rs.Close
 


    frmMain.Enabled = True
    
End Sub

Private Sub txtProds_Change(Index As Integer)

If txtProds(0).Text <> "" And txtProds(5).Text <> "" And cmbIVA.Text <> "" Then
lblGPU.Caption = (txtProds(0) * txtProds(5).Text) / 100

Dim ivtemp As Single
If (cmbIVA.ListIndex = 0) Then
ivtemp = 0
ElseIf (cmbIVA.ListIndex = 1) Then
ivtemp = IV1
ElseIf (cmbIVA.ListIndex = 2) Then
ivtemp = IV2
ElseIf (cmbIVA.ListIndex = 3) Then
ivtemp = IV3
ElseIf (cmbIVA.ListIndex = 4) Then
ivtemp = IV4
ElseIf (cmbIVA.ListIndex = 5) Then
ivtemp = IV5
End If

lblPFU.Caption = txtProds(0) + ((txtProds(0) * txtProds(5).Text) / 100) + (((txtProds(0) + ((txtProds(0) * txtProds(5).Text) / 100)) * ivtemp) / 100)


End If
End Sub
