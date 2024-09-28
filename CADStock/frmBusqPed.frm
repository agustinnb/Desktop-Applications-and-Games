VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBusqPed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar en pedido"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMod 
      Height          =   285
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtMarc 
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtProd 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtCod 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.ListView LVBusqPed 
      Height          =   2655
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4683
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Producto"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Marca"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Modelo"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Modelo:"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Marca:"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Producto:"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmBusqPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public idprovBusq As Integer
Private Sub Form_Load()
If (rs.State = adStateOpen) Then rs.Close
rs.Open "SELECT * FROM ProdXDist WHERE idProv = '" & idprovBusq & "'", cnn, adOpenStatic, adLockOptimistic
Call CargarListView9(LVBusqPed, rs)
If (rs.State = adStateOpen) Then rs.Close

End Sub

Private Sub LVBusqPed_DblClick()
Call Editar
    
End Sub
Private Sub Editar()

    Dim i As Integer
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LVBusqPed.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LVBusqPed.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
    frmAddPed.cmbCod.Text = LVBusqPed.SelectedItem.ListSubItems(1).Text
    Unload Me
  '  frmAddPed.cmbProd.Text = LVBusqPed.SelectedItem.ListSubItems(2).Text
   ' frmAddPed.lblMarc.Caption = LVBusqPed.SelectedItem.ListSubItems(3).Text
   ' frmAddPed.lblMod.Caption = LVBusqPed.SelectedItem.ListSubItems(4).Text
    
    
    End Sub

Private Sub txtCod_Change()
Call buscar
End Sub

Private Sub txtMarc_Change()
Call buscar
End Sub

Private Sub txtMod_Change()
Call buscar
End Sub

Private Sub txtProd_Change()
Call buscar
End Sub

Private Sub buscar()
Dim query As String
query = "SELECT * FROM ProdXDist WHERE idProv = '" & idprovBusq & "'"
If Not txtCod.Text = "" Then
query = query & " AND Codigo LIKE '%" & txtCod.Text & "%'"
End If
If Not txtProd.Text = "" Then
query = query & " AND Producto LIKE '%" & txtProd.Text & "%'"
End If
If Not txtMarc.Text = "" Then
query = query & " AND Marca LIKE '%" & txtMarc.Text & "%'"
End If
If Not txtMod.Text = "" Then
query = query & " AND Modelo LIKE '%" & txtMod.Text & "%'"
End If
If (rs.State = adStateOpen) Then rs.Close
rs.Open query, cnn, adOpenStatic, adLockOptimistic
Call CargarListView9(LVBusqPed, rs)
rs.Close
End Sub
