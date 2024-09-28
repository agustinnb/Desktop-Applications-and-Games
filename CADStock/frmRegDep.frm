VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegDep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de depositos"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13035
   Icon            =   "frmRegDep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   13035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   4260
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
         Text            =   "Descripcion"
         Object.Width           =   18416
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRegDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ObjItem1 As ListItem

Private Sub Form_Load()

rs.Open "SELECT * FROM MovDep", cnn, adOpenStatic, adLockOptimistic
Dim desde As String, producto As String, cantidad As String, hasta As String, fecha As String
Dim texto As String
While Not rs.EOF
If Not IsNull(rs!desde) Then
desde = rs!desde
Else
desde = ""
End If
producto = rs!producto
cantidad = rs!cantidad
hasta = rs!hasta
fecha = rs!fecha
If Trim(desde) = "" Then
texto = "Se ingresaron " & cantidad & " unidades del producto " & producto & " al deposito " & hasta
Else
texto = "Se pasaron " & cantidad & " unidades del producto " & producto & " desde el deposito " & desde & " hasta el deposito " & hasta
End If
    Set ObjItem1 = LV.ListItems.Add(, , rs!id)
                
           ObjItem1.SubItems(1) = texto
            ObjItem1.SubItems(2) = rs!fecha
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmDepositos.Enabled = True
End Sub

Private Sub Text1_Change()
Dim sql As String
sql = "SELECT * FROM MovDep WHERE id LIKE '%" & Text1.Text & "%' OR " & _
        "desde LIKE '%" & Text1.Text & "%' OR " & _
        "producto LIKE '%" & Text1.Text & "%' OR " & _
        "cantidad LIKE '%" & Text1.Text & "%' OR " & _
        "hasta LIKE '%" & Text1.Text & "%' OR " & _
        "fecha LIKE '%" & Text1.Text & "%'"
 LV.ListItems.Clear
rs.Open sql, cnn, adOpenStatic, adLockOptimistic

Dim desde As String, producto As String, cantidad As String, hasta As String, fecha As String
Dim texto As String
While Not rs.EOF
If Not IsNull(rs!desde) Then
desde = rs!desde
Else
desde = ""
End If
producto = rs!producto
cantidad = rs!cantidad
hasta = rs!hasta
fecha = rs!fecha
If Trim(desde) = "" Then
texto = "Se ingresaron " & cantidad & " unidades del producto " & producto & " al deposito " & hasta
Else
texto = "Se pasaron " & cantidad & " unidades del producto " & producto & " desde el deposito " & desde & " hasta el deposito " & hasta
End If
    Set ObjItem1 = LV.ListItems.Add(, , rs!id)
                
           ObjItem1.SubItems(1) = texto
            ObjItem1.SubItems(2) = rs!fecha
rs.MoveNext
Wend



rs.Close
End Sub
