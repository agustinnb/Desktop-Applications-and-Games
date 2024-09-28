VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CADStock Login"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3375
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3375
   StartUpPosition =   1  'CenterOwner
   Begin ChamaleonButton.ChameleonBtn cmdLogin 
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Entrar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "frmLogin.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdLogin_Click()
 Dim gLog As Boolean

gLog = doLogin(txtUser.Text, txtPass.Text)
If gLog = True Then
Unload Me
frmMain.Show
End If
End Sub

Private Sub Form_Load()
 Dim gLog As Boolean
 gLog = False
  firsttime = False
    ' Abre la conexión
    Call IniciarConexion
   Dim disco As String
   rs.Open "SELECT Disco FROM configuracion", cnn, adOpenStatic, adLockOptimistic
If rs.RecordCount > 0 Then
   disco = rs!disco
   If (disco <> Get_Numero_Serie(Mid(App.Path, 1, 3))) Then
   MsgBox "La copia del producto a otras computadoras no esta permitida."
   Unload Me
   rs.Close
   Else
   rs.Close
   End If
   
   End If
   
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM Perfiles", cnn, adOpenStatic, adLockOptimistic
If rs.RecordCount = 0 Then
OtraVez:

firsttime = True
Dim usuario As String
usuario = InputBox("Es la primera vez que ejecuta el programa. " & vbNewLine & "Por favor, Introduzca el nombre del usuario administrador", "Primera vez")
Dim password As String
password = InputBox("Por favor, Introduzca la contraseña del usuario administrador", "Primera vez")
If (Trim(usuario) = "") Or (Trim(password) = "") Then
response = MsgBox("El usuario o la contraseña son invalidos." & "¿Desea intentarlo de nuevo?", vbYesNo)
If response = vbYes Then
GoTo OtraVez
Else
Unload Me
Exit Sub
End If
End If
Dim quer As String
quer = "INSERT INTO Perfiles (Usuarios,Passwords"
Dim X As Integer
For X = 0 To 16
quer = quer & ",Prin" & X
Next X
For X = 0 To 5
quer = quer & ",Dist" & X
Next X
For X = 0 To 2
quer = quer & ",Dep" & X
Next X
For X = 0 To 2
quer = quer & ",Cli" & X
Next X
For X = 0 To 2
quer = quer & ",Ven" & X
Next X
quer = quer & ") VALUES ('" & usuario & "','" & password & "'"
For X = 0 To 16
quer = quer & ",'1'"
Next X
For X = 0 To 5
quer = quer & ",'1'"
Next X
For X = 0 To 2
quer = quer & ",'1'"
Next X
For X = 0 To 2
quer = quer & ",'1'"
Next X
For X = 0 To 2
quer = quer & ",'1'"
Next X
quer = quer & ")"
cnn.Execute quer
MsgBox "El usuario fue ingresado exitosamente." & vbNewLine & "Por favor, acuerdese de setear la configuración antes de empezar a usar el programa.", vbOKOnly



gLog = doLogin(usuario, password)
If gLog = True Then
frmConf.Show
Unload Me
End If

End If
If gLog = False Then
rs.Close
End If
End Sub
