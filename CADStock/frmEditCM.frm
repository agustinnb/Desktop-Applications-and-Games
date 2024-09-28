VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmEditCM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   Icon            =   "frmEditCM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbRI 
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
      ItemData        =   "frmEditCM.frx":058A
      Left            =   1680
      List            =   "frmEditCM.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   17
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   16
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   15
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   14
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
   End
   Begin ChamaleonButton.ChameleonBtn cmdCancelar 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   6120
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
      MICON           =   "frmEditCM.frx":05A0
      PICN            =   "frmEditCM.frx":05BC
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
      Top             =   6120
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
      MICON           =   "frmEditCM.frx":0B56
      PICN            =   "frmEditCM.frx":0B72
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
      Caption         =   "Responsable inscripto"
      ForeColor       =   &H00808080&
      Height          =   435
      Index           =   10
      Left            =   240
      TabIndex        =   22
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   21
      Top             =   5400
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   20
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   3960
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   4440
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   555
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
      Left            =   1560
      TabIndex        =   12
      Top             =   600
      Width           =   1665
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
      TabIndex        =   11
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUIT"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   0
      X1              =   120
      X2              =   3960
      Y1              =   5880
      Y2              =   5880
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
      TabIndex        =   7
      Top             =   240
      Width           =   1245
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
      TabIndex        =   6
      Top             =   240
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   240
      X2              =   4080
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmEditCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum EEACCION3
    AGREGAR_REGISTRO3 = 0
    EDITAR_REGISTRO3 = 1
End Enum

Public IdRegistro
Public ACCION As EEACCION3



Private Sub cmdGuardar_Click()

On Error GoTo ErrorSub
    
    
    ' Valida el Nombre que no este vacio
    ''''''''''''''''''''''''''''''''
    If Trim(Text1(1)) = "" Then
        MsgBox "El CUIT de registro no puede estar vacio", vbCritical, "Datos incompletos"
        Text1(1).SetFocus
        Exit Sub
    
    ' Valida el Apellido
    ''''''''''''''''''''''''''''''''
    ElseIf Trim(Text1(2)) = "" Then
        MsgBox "La Razon social no puede estar vacia", vbCritical, "Datos incompletos"
        Text1(2).SetFocus
        Exit Sub
End If



    'Agrega el registro
    '''''''''''''''''''''''''''''''
    
    Select Case ACCION
    Case EDITAR_REGISTRO3
        cnn.Execute "UPDATE ClientesM set CUIT = '" & Text1(1) & _
                                         "', RS = '" & Text1(2) & _
                                        "', Nombre = '" & Text1(3) & _
                                        "', Apellido = '" & Text1(0) & _
                                        "', Domicilio = '" & Text1(7) & _
                                        "', Email = '" & Text1(6) & _
                                         "', Telefono = '" & Text1(5) & _
                                         "', RI = '" & cmbRI.Text & _
                                         "', DNI = '" & Text1(4) & _
                                         "' where Id = " & IdRegistro & ""
    Case AGREGAR_REGISTRO3
        
        cnn.Execute "INSERT INTO ClientesM " & "(CUIT,RS,Nombre,Apellido,Domicilio,Email,Telefono,RI,DNI,FechaDeAlta) VALUES('" & _
                                 Text1(1) & "','" & _
                                 Text1(2) & "','" & _
                                 Text1(3) & "','" & _
                                 Text1(0) & "','" & _
                                 Text1(7) & "','" & _
                                 Text1(6) & "','" & _
                                 Text1(5) & "','" & _
                                 cmbRI.Text & "','" & _
                                 Text1(4) & "','" & _
                                 Format(Date, "dd/mm/yyyy") & "')"

    End Select
    
  

    DoEvents
    Unload Me
    Set FrmEdit = Nothing
Exit Sub
ErrorSub:
MsgBox Err.Description
If rs.State = adStateOpen Then rs.Close
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rs.Open "select * from ClientesM", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView8(frmClientes.LVCM, rs)
If rs.State = adStateOpen Then rs.Close
End Sub


