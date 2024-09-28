VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmEgresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmEgresos"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   Icon            =   "frmEgresos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin ChamaleonButton.ChameleonBtn cmdBorrar 
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Borrar todos"
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
      MICON           =   "frmEgresos.frx":058A
      PICN            =   "frmEgresos.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar Egreso"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtMot 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1200
      TabIndex        =   5
      Top             =   2760
      Width           =   5055
   End
   Begin VB.ComboBox CmbMon 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEgresos.frx":0B40
      Left            =   2160
      List            =   "frmEgresos.frx":0B4A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtMon 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin MSComctlLib.ListView LVE 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3625
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Monto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Motivo"
         Object.Width           =   5292
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdCancelar 
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
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
      MICON           =   "frmEgresos.frx":0B56
      PICN            =   "frmEgresos.frx":0B72
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdBorrarU 
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Borrar"
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
      MICON           =   "frmEgresos.frx":110C
      PICN            =   "frmEgresos.frx":1128
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblError 
      Alignment       =   2  'Center
      Caption         =   "lblError"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   6135
   End
   Begin VB.Label lblMot 
      Alignment       =   1  'Right Justify
      Caption         =   "Motivo:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblMon 
      Alignment       =   1  'Right Justify
      Caption         =   "Monto:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
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
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
 On Error Resume Next

Dim precio As Single
If Trim(txtMon.Text) = "" Or Not IsNumeric(txtMon.Text) Then
lblError.Caption = "Tiene que escribir un monto numerico"
Exit Sub
End If
If Trim(txtMot.Text) = "" Then
txtMot = " "
Exit Sub
End If
If (CmbMon.ListIndex = 1) Then
precio = Format(txtMon.Text * PDolar, "#,##0.00")
Else
precio = Format(txtMon.Text, "#,##0.00")
End If
lblError.Caption = ""
cnn.Execute "INSERT into Egresos (Monto,Motivo) VALUES ('" & precio & "','" & txtMot.Text & "')"
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM Egresos", cnn, adOpenStatic, adLockOptimistic
Call CargarListView5(LVE, rs)
rs.Close


End Sub

Private Sub cmdBorrar_Click()
Call EliminarT
End Sub

Private Sub cmdBorrarU_Click()
Call Eliminar
End Sub

Private Sub cmdCancelar_Click()
Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next
frmMain.Enabled = False
lblError.Caption = ""
Me.Caption = "Egresos de caja"
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM Egresos", cnn, adOpenStatic, adLockOptimistic
Call CargarListView5(LVE, rs)
rs.Close
CmbMon.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
frmMain.SetFocus
If rs.State = adStateOpen Then rs.Close

End Sub



Private Sub Eliminar()
 On Error Resume Next

    

    If (LVE.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LVE.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LVE.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "Monto " & .ListSubItems(1).Text & vbNewLine & _
                 "Motivo: " & .ListSubItems(2).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Egresos where Id = " & .Text & ""
            ' refresca el recordset
            rs.Open "select * from Egresos", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView5(LVE, rs)
        End If
    End With
    If rs.State = adStateOpen Then rs.Close
End Sub

Private Sub EliminarT()

    
 On Error Resume Next
    
        ' pregunta
        If MsgBox("Se van a eliminar todos los registros" & vbNewLine & "¿Esta seguro que desea hacerlo?", _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Egresos"
            ' refresca el recordset
            rs.Open "select * from Egresos", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView5(LVE, rs)
        End If
    
    If rs.State = adStateOpen Then rs.Close
End Sub

