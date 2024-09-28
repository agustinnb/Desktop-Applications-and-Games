VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmConf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3885
   Icon            =   "frmConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIVA5 
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtIVA3 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtIVA4 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtPD 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtIVA2 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtIVA 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtNL 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1815
   End
   Begin ChamaleonButton.ChameleonBtn cmdSave 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3480
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
      MICON           =   "frmConf.frx":058A
      PICN            =   "frmConf.frx":05A6
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
      Left            =   2400
      TabIndex        =   9
      Top             =   3480
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
      MICON           =   "frmConf.frx":0B40
      PICN            =   "frmConf.frx":0B5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdConfVis 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Configurar vistas"
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
      MICON           =   "frmConf.frx":10F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "IVA 5:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "IVA 3:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "IVA 4:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblPD 
      Alignment       =   1  'Right Justify
      Caption         =   "Precio Dolar:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblIVA2 
      Alignment       =   1  'Right Justify
      Caption         =   "IVA 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblIVA 
      Alignment       =   1  'Right Justify
      Caption         =   "IVA:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblNL 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre del local:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfVis_Click()
frmConfProfiles.Show
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
   
     
     If (txtNL.Text = "") Then
     MsgBox "El Nombre del Local no puede estar en blanco"

     Exit Sub
     End If
      If Not IsNumeric(txtIVA.Text) Then
     MsgBox "El IVA tiene que ser un numero y es obligatorio"

     Exit Sub
     End If
       If Not IsNumeric(txtIVA2.Text) Then
     MsgBox "El IVA 2 tiene que ser un numero y es obligatorio"

     Exit Sub
     End If
   
       If Not IsNumeric(txtPD.Text) Then
     MsgBox "El Precio del dolar tiene que ser un numero"

     Exit Sub
     End If
     
     If (Trim(txtIVA3.Text) = "") Then
     txtIVA3.Text = " "
     Else
      txtIVA3.Text = Trim(txtIVA3.Text)
      If Not IsNumeric(txtIVA3.Text) Then
        MsgBox "El IVA 3 tiene que ser un numero"
     Exit Sub
      End If
     End If
       If (Trim(txtIVA4.Text) = "") Then
     txtIVA4.Text = " "
     
     Else
      txtIVA4.Text = Trim(txtIVA4.Text)
      If Not IsNumeric(txtIVA4.Text) Then
        MsgBox "El IVA 4 tiene que ser un numero"
       Exit Sub
      End If
     End If
       If (Trim(txtIVA5.Text) = "") Then
     txtIVA5.Text = " "
     Else
     txtIVA5.Text = Trim(txtIVA5.Text)
     If Not IsNumeric(txtIVA5.Text) Then
        MsgBox "El IVA 5 tiene que ser un numero"
        Exit Sub
      End If
     End If
      rs.Open "select * from Configuracion", cnn, adOpenStatic, adLockOptimistic
   
     If (rs.RecordCount > 0) Then
  cnn.Execute "UPDATE Configuracion set NombreCasa = '" & txtNL.Text & _
                                         "', IVA1 = '" & txtIVA.Text & _
                                         "', IVA2 = '" & txtIVA2.Text & _
                                         "', IVA3 = '" & txtIVA3.Text & _
                                         "', IVA4 = '" & txtIVA4.Text & _
                                         "', IVA5 = '" & txtIVA5.Text & _
                                         "', PDolar = '" & txtPD.Text & _
                                         "' where Id = 1"
    pdolar = txtPD.Text
 rs.Close
   rs.Open "select * from Productos", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView2(frmMain.LV2, rs)
If (rs.State = adStateOpen) Then rs.Close
   MsgBox "La configuracion fue actualizada con exito"
Unload Me
     Exit Sub
Else

cnn.Execute "INSERT INTO Configuracion " & "(id,NombreCasa,IVA1,IVA2,IVA3,IVA4,IVA5,PDolar,Disco) VALUES('1','" & _
                                 txtNL.Text & "','" & _
                                 txtIVA.Text & "','" & _
                                 txtIVA2.Text & "','" & _
                                 txtIVA3.Text & "','" & _
                                 txtIVA4.Text & "','" & _
                                 txtIVA5.Text & "','" & _
                                 txtPD.Text & "','" & Get_Numero_Serie(Mid(App.Path, 1, 3)) & "')"
   If (rs.State = adStateOpen) Then rs.Close
   MsgBox "La configuracion fue ingresada con exito"
   If firsttime = True Then
frmMain.Show
Unload Me
End If
End If


End Sub

Private Sub Form_Load()
    ' Abre la conexión

    ' carga el Recorset con todos los datos
    rs.Open "select * from Configuracion", cnn, adOpenStatic, adLockOptimistic
  '  On Error GoTo ErrorSub
 If (rs.RecordCount > 0) Then
    txtNL.Text = rs!NombreCasa
    txtIVA.Text = rs!IVA1
    txtIVA2.Text = rs!IVA2
    txtIVA3.Text = rs!IVA3
    txtIVA4.Text = rs!IVA4
    txtIVA5.Text = rs!IVA5
  txtPD.Text = rs!pdolar
  Else
     txtNL.Text = ""
          txtIVA.Text = ""
    txtIVA2.Text = ""
    txtIVA3.Text = ""
    txtIVA4.Text = ""
    txtIVA5.Text = ""
    txtPD.Text = ""
  End If
  
      rs.Close
    
      Exit Sub
    
ErrorSub:
        txtNL.Text = ""
          txtIVA.Text = ""
    txtIVA2.Text = ""
    txtIVA3.Text = ""
    txtIVA4.Text = ""
    txtIVA5.Text = ""
    txtPD.Text = ""
If (rs.State = adStateOpen) Then rs.Close
    
    If Err.Number = 94 Then Resume Next

End Sub


Private Sub Form_Unload(Cancel As Integer)

    
  If rs.State = adStateOpen Then rs.Close
  
 

    frmMain.Enabled = True
    
End Sub
