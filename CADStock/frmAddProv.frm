VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmAddProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar distribuidor"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3135
   Icon            =   "frmAddProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPDolar 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtTel 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtCon 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtAddProv 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   615
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Agregar"
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
      MICON           =   "frmAddProv.frx":058A
      PICN            =   "frmAddProv.frx":05A6
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
      Left            =   480
      TabIndex        =   13
      Top             =   2640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Agregar Productos"
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
      MICON           =   "frmAddProv.frx":095A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      Caption         =   "Precio Dolar:"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblID 
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "ID:"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Direccion:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Telefono:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Contacto:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmAddProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public modificar As Boolean
Private Sub cmdOpciones_Click(Index As Integer)
If (Trim(txtAddProv.Text) <> "") Then
Dim pdolar As String
If (modificar = False) Then
 rs.Open "select * from Distribuidores", cnn, adOpenStatic, adLockOptimistic
Dim val As String
While Not (rs.EOF)
val = rs!Nombre
If (val = txtAddProv.Text) Then
MsgBox "Ese distribuidor ya existe!"
Exit Sub
End If
rs.MoveNext
Wend
rs.Close
End If
If (txtCon.Text = "") Then
txtCon.Text = " "
End If
If (txtTel.Text = "") Then
txtTel.Text = " "
End If
If (txtDir.Text = "") Then
txtDir.Text = " "
End If
If (Trim(txtPDolar.Text) = "" Or Not IsNumeric(txtPDolar.Text)) Then
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM configuracion", cnn, adOpenStatic, adLockOptimistic
pdolar = rs!pdolar
rs.Close
Else
pdolar = txtPDolar.Text

End If

If (modificar = False) Then
cnn.Execute "INSERT INTO Distribuidores(Nombre,Contacto,Telefono,Direccion,PDolar) VALUES ('" & txtAddProv.Text & "','" & txtCon.Text & "','" & txtTel.Text & "','" & txtDir.Text & "','" & pdolar & "')"
rs.Open "SELECT * FROM Distribuidores", cnn, adOpenStatic, adLockOptimistic
Call CargarListView6(frmProveedores.LVD, rs)
rs.Close
Else
cnn.Execute "UPDATE Distribuidores SET Nombre = '" & txtAddProv.Text & "', Contacto ='" & txtCon.Text & "', Telefono = '" & txtTel.Text & "', Direccion = '" & txtDir.Text & "', PDolar = '" & pdolar & "' WHERE Id = " & lblID.Caption
rs.Open "SELECT * FROM Distribuidores", cnn, adOpenStatic, adLockOptimistic
Call CargarListView6(frmProveedores.LVD, rs)
rs.Close
End If


If (Index = 1) Then
frmAddProdDist.pdolar = pdolar
If (modificar = False) Then
frmAddProdDist.modificar = False
rs.Open "SELECT id FROM Distribuidores where Nombre = '" & txtAddProv.Text & "'", cnn, adOpenStatic, adLockOptimistic
frmAddProdDist.id = rs!id
rs.Close

Else
frmAddProdDist.modificar = True
frmAddProdDist.id = lblID.Caption
End If
With frmAddProdDist
.lblProv.Caption = txtAddProv.Text
.Show vbModal
End With
End If

Unload Me
Else
MsgBox "Tiene que ingresar el nombre de un distribuidor"
End If


End Sub

