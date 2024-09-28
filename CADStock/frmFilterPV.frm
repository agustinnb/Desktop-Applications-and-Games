VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmFilterPV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtrar y ordenar"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6390
   Icon            =   "frmFilterPV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmFilterPV.frx":058A
      Left            =   1320
      List            =   "frmFilterPV.frx":05AF
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Font Size"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmFilterPV.frx":062E
      Left            =   3960
      List            =   "frmFilterPV.frx":0653
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Font Size"
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cerrar"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFilterPV.frx":06D2
      PICN            =   "frmFilterPV.frx":06EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn CmdOrdenar 
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmFilterPV.frx":0C88
      PICN            =   "frmFilterPV.frx":0CA4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   -1  'True
   End
   Begin ChamaleonButton.ChameleonBtn CmdOrdenar 
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   5
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmFilterPV.frx":123E
      PICN            =   "frmFilterPV.frx":125A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Height          =   1455
      Left            =   120
      Top             =   120
      Width           =   6135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   3
      X1              =   360
      X2              =   6240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ordenar por"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   1
      X1              =   3600
      X2              =   3600
      Y1              =   1440
      Y2              =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Por el campo"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmFilterPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChameleonBtn1_Click()

    Unload Me
End Sub

' Ordena en forma Ascendente y descendente el LV
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CmdOrdenar_Click(Index As Integer)
    CmdOrdenar(0).Value = False
    CmdOrdenar(1).Value = False
    CmdOrdenar(Index).Value = True
    ChameleonBtn1.Enabled = False
    Call Filtrar
    ChameleonBtn1.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Me
    End If
End Sub

Private Sub Form_Load()

    With frmMain
        Me.Move (.Left + .LV3.Left), _
                (.LV3.Height + .LV3.Top + .Top + 500)
    End With
    frmMain.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
If rs.State = adStateOpen Then rs.Close
End Sub

Private Sub txtSearch_Change()
ChameleonBtn1.Enabled = False
    Call Filtrar
    ChameleonBtn1.Enabled = True
End Sub

Private Sub Combo1_Click()
ChameleonBtn1.Enabled = False
    Call Filtrar
    ChameleonBtn1.Enabled = True
    End Sub

Private Sub Combo2_Click()
ChameleonBtn1.Enabled = False
    Call Filtrar
    ChameleonBtn1.Enabled = True
End Sub

Public Sub Filtrar()

Dim Campo, OrderByCampo, Orden As String
Dim sql As String

    If Combo1.ListIndex = -1 Then
        Combo1.ListIndex = 0
    End If
    If Combo2.ListIndex = -1 Then
        Combo2.ListIndex = 0
    End If
    If Combo1.ListIndex = 0 Then
        Campo = "Id"
    ElseIf Combo1.ListIndex = 1 Then
        Campo = "Producto"
    ElseIf Combo1.ListIndex = 2 Then
        Campo = "Marca"
        ElseIf Combo1.ListIndex = 3 Then
        Campo = "Modelo"
    ElseIf Combo1.ListIndex = 4 Then
        Campo = "Cantidad"
            ElseIf Combo1.ListIndex = 5 Then
        Campo = "VendidoPor"
           ElseIf Combo1.ListIndex = 6 Then
        Campo = "PrecioU"
           ElseIf Combo1.ListIndex = 7 Then
        Campo = "PrecioVenta"
           ElseIf Combo1.ListIndex = 8 Then
        Campo = "FechaDeAlta"
           ElseIf Combo1.ListIndex = 9 Then
        Campo = "FechaDeBaja"
          ElseIf Combo1.ListIndex = 10 Then
        Campo = "Deposito"
    End If
    
    Select Case Combo2.ListIndex
        Case 0: OrderByCampo = "Id"
        Case 1: OrderByCampo = "Producto"
        Case 2: OrderByCampo = "Marca"
        Case 3: OrderByCampo = "Modelo"
        Case 4: OrderByCampo = "Cantidad"
        Case 5: OrderByCampo = "VendidoPor"
        Case 6: OrderByCampo = "PrecioU"
        Case 7: OrderByCampo = "PrecioVenta"
        Case 8: OrderByCampo = "FechaDeAlta"
        Case 9: OrderByCampo = "FechaDeBaja"
          Case 10: OrderByCampo = "Deposito"
    End Select

    If CmdOrdenar(0).Value Then Orden = "asc"
    If CmdOrdenar(1).Value Then Orden = "desc"

    ' si el recorset está abierto lo cierra
    If rs.State = adStateOpen Then
        rs.Close
    End If
    
    If Campo = "FechaDeBaja" Then
    Campo = "FechaDeAlta"
    End If
    If Campo = "PrecioVenta" Then
    Campo = "PrecioU"
    End If
    If OrderByCampo = "FechaDeBaja" Then
    OrderByCampo = "FechaDeAlta"
    End If
    If OrderByCampo = "PrecioVenta" Then
    OrderByCampo = "PrecioU"
    End If
    
    If (Campo <> "VendidoPor") And (OrderByCampo <> "VendidoPor") Then
    sql = "SELECT * FROM Productos Where " & _
                         Campo & " like '" & txtSearch & _
                        "%' order by " & OrderByCampo & " " & Orden
    
    rs.Open sql, cnn, adOpenStatic, adLockOptimistic
    
    Call CargarListView2(frmMain.LV2, rs)
 End If
If Campo = "Id" Then
Campo = "IDProd"
End If
If Campo = "FechaDeAlta" Then
    Campo = "FechaDeBaja"
    End If
    If Campo = "PrecioU" Then
    Campo = "PrecioVenta"
    End If
     If OrderByCampo = "FechaDeAlta" Then
    OrderByCampo = "FechaDeBaja"
    End If
    If OrderByCampo = "PrecioU" Then
    OrderByCampo = "PrecioVenta"
    End If
    If OrderByCampo <> "Deposito" And Campo <> "Deposito" Then
    sql = "SELECT * FROM Ventas Where " & _
                         Campo & " like '" & txtSearch & _
                        "%' order by " & OrderByCampo & " " & Orden
    If rs.State = adStateOpen Then rs.Close
    rs.Open sql, cnn, adOpenStatic, adLockOptimistic
    Else
    sql = "SELECT * FROM Ventas"
     If rs.State = adStateOpen Then rs.Close
    rs.Open sql, cnn, adOpenStatic, adLockOptimistic
    End If
    
   Call CargarListView3(frmMain.LV3, rs)
  
 



End Sub


