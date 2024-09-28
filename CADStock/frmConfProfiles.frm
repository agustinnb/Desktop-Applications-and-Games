VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmConfProfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Perfil"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   Icon            =   "frmConfProfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin ChamaleonButton.ChameleonBtn cmdDelPer 
      Height          =   255
      Left            =   2640
      TabIndex        =   43
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Borrar perfil"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConfProfiles.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver precio unitario"
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   42
      Top             =   1800
      Width           =   2175
   End
   Begin ChamaleonButton.ChameleonBtn cmdGP 
      Height          =   495
      Left            =   360
      TabIndex        =   41
      Top             =   7200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Guardar Perfil"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConfProfiles.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbProf 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Left            =   4320
      TabIndex        =   39
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      TabIndex        =   37
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox chkVen 
      Caption         =   "Puede eliminar vendedores"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   35
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CheckBox chkVen 
      Caption         =   "Puede modificar vendedores"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   34
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CheckBox chkVen 
      Caption         =   "Puede agregar vendedores"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   33
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CheckBox chkCli 
      Caption         =   "Puede eliminar clientes"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   31
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CheckBox chkCli 
      Caption         =   "Puede modificar clientes"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   30
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CheckBox chkCli 
      Caption         =   "Puede agregar clientes"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   29
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CheckBox chkDep 
      Caption         =   "Puede eliminar depositos"
      Height          =   315
      Index           =   2
      Left            =   3360
      TabIndex        =   27
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CheckBox chkDep 
      Caption         =   "Puede modificar depositos"
      Height          =   315
      Index           =   1
      Left            =   3360
      TabIndex        =   26
      Top             =   3960
      Width           =   2535
   End
   Begin VB.CheckBox chkDep 
      Caption         =   "Puede agregar depositos"
      Height          =   315
      Index           =   0
      Left            =   3360
      TabIndex        =   25
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CheckBox chkDist 
      Caption         =   "Puede eliminar pedidos"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   23
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CheckBox chkDist 
      Caption         =   "Puede modificar pedidos"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   22
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CheckBox chkDist 
      Caption         =   "Puede hacer pedidos"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   21
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CheckBox chkDist 
      Caption         =   "Puede eliminar distribuidores"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   20
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CheckBox chkDist 
      Caption         =   "Puede modificar distribuidores"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   19
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CheckBox chkDist 
      Caption         =   "Puede agregar distribuidores"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   18
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver configuración"
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   16
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver distribuidores"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   15
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver vendedores"
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   14
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver clientes"
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   13
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver depositos"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   12
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver egresos de caja"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   11
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede vender productos"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   10
      Top             =   4680
      Width           =   2535
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede eliminar ventas"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede eliminar productos"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede editar productos"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede agregar productos"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver seña"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver precio de venta"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver forma de pago"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver cantidad"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CheckBox chkPrin 
      Caption         =   "Puede ver margen"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Password:"
      Height          =   255
      Left            =   3360
      TabIndex        =   38
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Vendedores"
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
      TabIndex        =   32
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Clientes"
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
      TabIndex        =   28
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Depositos"
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
      TabIndex        =   24
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Distribuidores"
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
      TabIndex        =   17
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Panel principal"
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
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmConfProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub chkPrin_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If chkPrin(6).Value = 1 Then
If (Index = 16 Or Index = 0) Then
If chkPrin(Index).Value = 0 Then
MsgBox "Recuerde que el usuario podra ver el margen y el precio unitario si tiene habilitado el area de editar producto", vbInformation
End If
End If
End If
End Sub

Private Sub cmbProf_Click()
If cmbProf.Text <> "Nuevo..." Then
rs.Open "SELECT * FROM Perfiles WHERE usuarios = '" & cmbProf.Text & "'", cnn, adOpenStatic, adLockOptimistic
txtUser.Text = rs!Usuarios
txtPass.Text = rs!Passwords
Dim X As Integer

chkPrin(0).Value = rs!Prin0
chkPrin(1).Value = rs!Prin1
chkPrin(2).Value = rs!Prin2
chkPrin(3).Value = rs!Prin3
chkPrin(4).Value = rs!Prin4
chkPrin(5).Value = rs!Prin5
chkPrin(6).Value = rs!Prin6
chkPrin(7).Value = rs!Prin7
chkPrin(8).Value = rs!Prin8
chkPrin(9).Value = rs!Prin9
chkPrin(10).Value = rs!Prin10
chkPrin(11).Value = rs!Prin11
chkPrin(12).Value = rs!Prin12
chkPrin(13).Value = rs!Prin13
chkPrin(14).Value = rs!Prin14
chkPrin(15).Value = rs!Prin15
chkPrin(16).Value = rs!Prin16


chkDist(0).Value = rs!Dist0
chkDist(1).Value = rs!Dist1
chkDist(2).Value = rs!Dist2
chkDist(3).Value = rs!Dist3
chkDist(4).Value = rs!Dist4
chkDist(5).Value = rs!Dist5

chkDep(0).Value = rs!Dep0
chkDep(1).Value = rs!Dep1
chkDep(2).Value = rs!Dep2

chkCli(0).Value = rs!Cli0
chkCli(1).Value = rs!Cli1
chkCli(2).Value = rs!Cli2

chkVen(0).Value = rs!Ven0
chkVen(1).Value = rs!Ven1
chkVen(2).Value = rs!Ven2
rs.Close

End If
End Sub

Private Sub cmdDelPer_Click()
If cmbProf.Text <> "Nuevo..." Then
If cmbProf.ListCount = 2 Then
MsgBox "No se puede borrar este perfil", vbOKOnly
Exit Sub
End If

cnn.Execute "DELETE FROM Perfiles WHERE usuarios = '" & cmbProf.Text & "'"

cmbProf.Clear
cmbProf.AddItem "Nuevo..."
rs.Open "SELECT * FROM Perfiles", cnn, adOpenStatic, adLockOptimistic
While Not rs.EOF
cmbProf.AddItem rs!Usuarios
rs.MoveNext
Wend
rs.Close
cmbProf.ListIndex = 0


Else
MsgBox "Debe seleccionar un perfil para borrar", vbOKOnly
End If
End Sub

Private Sub cmdGP_Click()


Dim query As String
Dim esta As Boolean
Dim X As Integer
If cmbProf.Text = "Nuevo..." Then

esta = False

If Trim(txtUser.Text) <> "" And Trim(txtPass.Text) <> "" Then
rs.Open "SELECT * FROM Perfiles", cnn, adOpenStatic, adLockOptimistic
While Not rs.EOF
If (LCase$(rs!Usuarios) = LCase$(txtUser.Text)) Then
esta = True
End If
rs.MoveNext
Wend
rs.Close
If esta = True Then
MsgBox "El usuario ya se encuentra registrado", vbExclamation
Exit Sub
End If
query = "INSERT INTO Perfiles " & "(Usuarios,Passwords"

For X = 0 To 16
query = query & ",Prin" & X
Next X
For X = 0 To 5
query = query & ",Dist" & X
Next X
For X = 0 To 2
query = query & ",Dep" & X
Next X
For X = 0 To 2
query = query & ",Cli" & X
Next X
For X = 0 To 2
query = query & ",Ven" & X
Next X
query = query & ") VALUES('" & txtUser.Text & "','" & txtPass.Text & "'"
For X = 0 To 16
query = query & ",'" & chkPrin(X).Value & "'"
Next X
For X = 0 To 5
query = query & ",'" & chkDist(X).Value & "'"
Next X
For X = 0 To 2
query = query & ",'" & chkDep(X).Value & "'"
Next X
For X = 0 To 2
query = query & ",'" & chkCli(X).Value & "'"
Next X
For X = 0 To 2
query = query & ",'" & chkVen(X).Value & "'"
Next X
query = query & ")"
Call submitquery(query)
MsgBox "El usuario se añadio con exito", vbOKOnly
Unload Me


Else
MsgBox "Tiene que escribir un usuario y una contraseña", vbOKOnly
Exit Sub
End If
Else
If cmbProf.ListCount = 2 Then
If chkPrin(15).Value = 0 Then
MsgBox "No se puede guardar este perfil, debe haber por lo menos un usuario administrador", vbOKOnly
Exit Sub
End If
End If

If Trim(txtUser.Text) <> "" And Trim(txtPass.Text) <> "" Then
If LCase$(txtUser.Text) <> LCase$(cmbProf.Text) Then
esta = False
rs.Open "SELECT * FROM Perfiles", cnn, adOpenStatic, adLockOptimistic
While Not rs.EOF
If (LCase$(rs!Usuarios) = LCase$(txtUser.Text)) Then
esta = True
End If
rs.MoveNext
Wend
rs.Close
If esta = True Then
MsgBox "El usuario ya se encuentra registrado", vbExclamation
Exit Sub
End If
End If

query = "UPDATE Perfiles SET Usuarios = '" & txtUser.Text & "', Passwords = '" & txtPass.Text & "'"
For X = 0 To 16
query = query & ", Prin" & X & " = '" & chkPrin(X).Value & "'"
Next X
For X = 0 To 5
query = query & ", Dist" & X & " = '" & chkDist(X).Value & "'"
Next X
For X = 0 To 2
query = query & ", Dep" & X & " = '" & chkDep(X).Value & "'"
Next X
For X = 0 To 2
query = query & ", Cli" & X & " = '" & chkCli(X).Value & "'"
Next X
For X = 0 To 2
query = query & ", Ven" & X & " = '" & chkVen(X).Value & "'"
Next X

query = query & " WHERE Usuarios = '" & txtUser.Text & "'"

cnn.Execute query
MsgBox "El usuario se actualizo con exito", vbOKOnly
Unload Me
Else
MsgBox "Tiene que escribir un usuario y una contraseña", vbOKOnly
Exit Sub
End If














End If


End Sub

Private Sub Form_Load()
cmbProf.AddItem "Nuevo..."
rs.Open "SELECT * FROM Perfiles", cnn, adOpenStatic, adLockOptimistic
While Not rs.EOF
cmbProf.AddItem rs!Usuarios
rs.MoveNext
Wend
rs.Close
cmbProf.ListIndex = 0
Dim X As Integer
For X = 0 To 14
chkPrin(X).Value = 1
Next X
For X = 0 To 4
chkDist(X).Value = 1
Next X
For X = 0 To 1
chkDep(X).Value = 1
Next X
For X = 0 To 1
chkCli(X).Value = 1
Next X
For X = 0 To 1
chkVen(X).Value = 1
Next X

chkPrin(15).Value = 0
chkPrin(16).Value = 1
End Sub
