VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmAddPed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hacer Pedido"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   Icon            =   "frmAddPed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuscPed 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3240
      TabIndex        =   26
      Top             =   1080
      Width           =   1095
   End
   Begin ChamaleonButton.ChameleonBtn cmdAddPed 
      Height          =   735
      Left            =   3000
      TabIndex        =   25
      Top             =   9000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Agregar Pedido"
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
      MICON           =   "frmAddPed.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      ItemData        =   "frmAddPed.frx":05A6
      Left            =   4680
      List            =   "frmAddPed.frx":05B0
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtMargen 
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdEl 
      Caption         =   "-"
      Height          =   255
      Left            =   7680
      TabIndex        =   20
      Top             =   5280
      Width           =   255
   End
   Begin VB.ComboBox cmbIVA 
      Height          =   315
      ItemData        =   "frmAddPed.frx":05BC
      Left            =   1320
      List            =   "frmAddPed.frx":05BE
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3960
      Width           =   975
   End
   Begin VB.ComboBox cmbEst 
      Height          =   315
      ItemData        =   "frmAddPed.frx":05C0
      Left            =   1680
      List            =   "frmAddPed.frx":05CD
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddTo 
      Caption         =   "Agregar a pedido"
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   4680
      Width           =   2055
   End
   Begin VB.ComboBox cmbProd 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox cmbCod 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtCant 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin MSComctlLib.ListView LVP 
      Height          =   3015
      Left            =   480
      TabIndex        =   16
      Top             =   5280
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Marca"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modelo"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Precio Unitario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cantidad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "IVA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Margen"
         Object.Width           =   2540
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelSE 
      Height          =   735
      Index           =   0
      Left            =   6600
      TabIndex        =   24
      Top             =   9000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Exportar Pedido"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAddPed.frx":05EF
      PICN            =   "frmAddPed.frx":060B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Margen:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblPrecioU 
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblMod 
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblMarc 
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "IVA:"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Estado:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Precio Unitario:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Modelo:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Marca:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Producto:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblProv 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Proveedor:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmAddPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As Integer
Public ProvN As String
Public modificara As Boolean
Dim idprov As Integer
Public pdolar
Dim hubocambio As Boolean

Private Sub cmbCod_Click()
On Error Resume Next
If modificara = False Then
rs.Open "SELECT * FROM ProdXDist WHERE IdProv = '" & id & "' AND Codigo = '" & cmbCod.Text & "'", cnn, adOpenStatic, adLockOptimistic
cmbProd.ListIndex = cmbCod.ListIndex
lblMarc.Caption = rs!marca
lblMod.Caption = rs!modelo
lblPrecioU.Caption = rs!preciou
cmbIVA.ListIndex = rs!IVA
cmbIVA.Refresh

cmbMon.ListIndex = rs!Moneda - 1

cmbMon.Refresh

txtMargen.Text = rs!Margen
rs.Close
Else
rs.Open "SELECT * FROM ProdXDist WHERE IdProv = '" & idprov & "' AND Codigo = '" & cmbCod.Text & "'", cnn, adOpenStatic, adLockOptimistic
cmbProd.ListIndex = cmbCod.ListIndex
lblMarc.Caption = rs!marca
lblMod.Caption = rs!modelo
lblPrecioU.Caption = rs!preciou

cmbMon.ListIndex = rs!Moneda - 1
cmbMon.Refresh
cmbIVA.ListIndex = rs!IVA
cmbIVA.Refresh
txtMargen.Text = rs!Margen
rs.Close
End If
End Sub



Private Sub cmbProd_Click()
On Error Resume Next

If modificara = False Then
rs.Open "SELECT * FROM ProdXDist WHERE IdProv = '" & id & "' AND Producto = '" & cmbProd.Text & "'", cnn, adOpenStatic, adLockOptimistic
cmbCod.ListIndex = cmbProd.ListIndex
lblMarc.Caption = rs!marca
lblMod.Caption = rs!modelo
lblPrecioU.Caption = rs!preciou
cmbMon.ListIndex = rs!Moneda - 1
cmbMon.Refresh
cmbIVA.ListIndex = rs!IVA
cmbIVA.Refresh
txtMargen.Text = rs!Margen
rs.Close

Else

rs.Open "SELECT * FROM ProdXDist WHERE IdProv = '" & idprov & "' AND Producto = '" & cmbProd.Text & "'", cnn, adOpenStatic, adLockOptimistic
cmbCod.ListIndex = cmbProd.ListIndex
lblMarc.Caption = rs!marca
lblMod.Caption = rs!modelo
lblPrecioU.Caption = rs!preciou
cmbMon.ListIndex = rs!Moneda - 1
cmbMon.Refresh
cmbIVA.ListIndex = rs!IVA
cmbIVA.Refresh
txtMargen.Text = rs!Margen
rs.Close

End If
End Sub

Private Sub cmdAddPed_Click()

If (LVP.ListItems.Count <> 0) Then
Dim i As Integer
Dim query As String

If modificara = False Then
query = "INSERT INTO Pedidos (Distribuidor,Estado,PDolar) VALUES ('" & lblProv.Caption & "','" & cmbEst.Text & "','" & pdolar & "')"
cnn.Execute query
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT id FROM pedidos WHERE Distribuidor = '" & lblProv.Caption & "'"
rs.MoveLast
id = rs!id
rs.Close
i = 1
query = "INSERT INTO ProdXPed (IdPed, Codigo, Producto, Marca, Modelo, PrecioU, Cantidad, IVA, margen) VALUES "


While i <> LVP.ListItems.Count + 1

query = query & "('" & id & "','" & LVP.ListItems(i) & "','" & LVP.ListItems(i).ListSubItems(1) & "','" & LVP.ListItems(i).ListSubItems(2) & "','" & LVP.ListItems(i).ListSubItems(3) & "','" & LVP.ListItems(i).ListSubItems(4) & "','" & LVP.ListItems(i).ListSubItems(5) & "','" & LVP.ListItems(i).ListSubItems(6) & "','" & LVP.ListItems(i).ListSubItems(7) & "')"
cnn.Execute query


query = "INSERT INTO ProdXPed (IdPed, Codigo, Producto, Marca, Modelo, PrecioU, Cantidad, IVA, margen) VALUES "
i = i + 1
Wend
Else

query = "Update Pedidos SET Distribuidor = '" & lblProv.Caption & "', Estado = '" & cmbEst.Text & "',PDolar='" & pdolar & "' WHERE Id = " & id
cnn.Execute query
If rs.State = adStateOpen Then rs.Close
If (hubocambio = True) Then
query = "DELETE FROM ProdXPed WHERE IdPed = '" & id & "'"
cnn.Execute query

i = 1
query = "INSERT INTO ProdXPed (IdPed, Codigo, Producto, Marca, Modelo, PrecioU, Cantidad, IVA, margen) VALUES "


While i <> LVP.ListItems.Count + 1

query = query & "('" & id & "','" & LVP.ListItems(i) & "','" & LVP.ListItems(i).ListSubItems(1) & "','" & LVP.ListItems(i).ListSubItems(2) & "','" & LVP.ListItems(i).ListSubItems(3) & "','" & LVP.ListItems(i).ListSubItems(4) & "','" & LVP.ListItems(i).ListSubItems(5) & "','" & LVP.ListItems(i).ListSubItems(6) & "','" & LVP.ListItems(i).ListSubItems(7) & "')"
cnn.Execute query


query = "INSERT INTO ProdXPed (IdPed, Codigo, Producto, Marca, Modelo, PrecioU, Cantidad, IVA, margen) VALUES "
i = i + 1

Wend
End If

End If

If (cmbEst.Text = "Entregado") Then
 If MsgBox("¿Desea Ingresar el pedido al Stock?", _
                 vbExclamation + vbYesNo, "Agregar a stock") = vbYes Then
  frmAddPed2Dep.Show vbModal
  Me.Enabled = False
  '             i = 1
  '  ReDim queryg(LVP.ListItems.Count + 1) As String
  '                queryg(i) = "INSERT INTO Productos (Producto, Marca, Modelo, Distribuidor, Cantidad, PrecioU,FechaDeAlta , IVA, Margen) VALUES "

             '      While i <> LVP.ListItems.Count + 1

              '     queryg(i) = query & "('" & LVP.ListItems(i).ListSubItems(1) & "','" & LVP.ListItems(i).ListSubItems(2) & "','" & LVP.ListItems(i).ListSubItems(3) & "','" & lblProv.Caption & "','" & LVP.ListItems(i).ListSubItems(5) & "','" & LVP.ListItems(i).ListSubItems(4) & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & LVP.ListItems(i).ListSubItems(6) & "', '" & LVP.ListItems(i).ListSubItems(7) & "')"
                ' cnn.Execute query

             '    i = i + 1

     '              queryg(i) = "INSERT INTO Productos (Producto, Marca, Modelo, Distribuidor, Cantidad, PrecioU,FechaDeAlta , IVA, Margen) VALUES "
 
' Wend

  '               rs.Open "SELECT * FROM Productos", cnn, adOpenStatic, adLockOptimistic
   '              Call CargarListView2(frmMain.LV2, rs)
                 
                 
                End If
End If

rs.Open "SELECT * FROM Pedidos", cnn, adOpenStatic, adLockOptimistic
Call CargarListView7(frmProveedores.LVP, rs)
rs.Close
Unload Me
Else
MsgBox "No se pidio nada"
Exit Sub
End If

End Sub

Private Sub cmdAddTo_Click()
If (Trim(txtCant.Text) <> "" And IsNumeric(Trim(txtCant)) = True) And Trim(txtCant) <> "0" Then
hubocambio = True
 Set ObjItem = LVP.ListItems.Add(, , cmbCod.Text)
                
           
           ObjItem.SubItems(1) = cmbProd.Text
           ObjItem.SubItems(2) = lblMarc.Caption
           ObjItem.SubItems(3) = lblMod.Caption
           If (cmbMon.ListIndex = 0) Then
           ObjItem.SubItems(4) = lblPrecioU.Caption
           Else
           ObjItem.SubItems(4) = lblPrecioU.Caption * pdolar
           
           End If
           ObjItem.SubItems(5) = txtCant.Text
        ObjItem.SubItems(6) = cmbIVA.ListIndex
           ObjItem.SubItems(7) = txtMargen.Text

Else
MsgBox ("La cantidad es incorrecta")
Exit Sub
End If

End Sub



Private Sub cmdBuscPed_Click()
With frmBusqPed
.idprovBusq = id
.Show vbModal

End With
End Sub

Private Sub cmdEl_Click()

    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LVP.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LVP.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
    
    hubocambio = True
Dim codigo() As String
Dim producto() As String
Dim marca() As String
Dim modelo() As String
Dim preciou() As String
Dim cantidad() As String
Dim IVA() As String
Dim Margen() As String
ReDim codigo(LVP.ListItems.Count) As String
ReDim producto(LVP.ListItems.Count) As String
ReDim marca(LVP.ListItems.Count) As String
ReDim modelo(LVP.ListItems.Count) As String
ReDim preciou(LVP.ListItems.Count) As String
ReDim cantidad(LVP.ListItems.Count) As String
ReDim IVA(LVP.ListItems.Count) As String
ReDim Margen(LVP.ListItems.Count) As String
Dim xa As Integer, xb As Integer
xa = 0
xb = 0

While Not xa = LVP.ListItems.Count
If (LVP.ListItems(xa + 1).Index <> LVP.SelectedItem.Index) Then
codigo(xb) = LVP.ListItems(xa + 1)
producto(xb) = LVP.ListItems(xa + 1).ListSubItems(1)
marca(xb) = LVP.ListItems(xa + 1).ListSubItems(2)
modelo(xb) = LVP.ListItems(xa + 1).ListSubItems(3)
preciou(xb) = LVP.ListItems(xa + 1).ListSubItems(4)
cantidad(xb) = LVP.ListItems(xa + 1).ListSubItems(5)
IVA(xb) = LVP.ListItems(xa + 1).ListSubItems(6)
Margen(xb) = LVP.ListItems(xa + 1).ListSubItems(7)
xb = xb + 1
End If
xa = xa + 1
Wend
xa = 0
LVP.ListItems.Clear
While Not xa = xb


 Set ObjItem = LVP.ListItems.Add(, , codigo(xa))
                
           
           ObjItem.SubItems(1) = producto(xa)
           ObjItem.SubItems(2) = marca(xa)
           ObjItem.SubItems(3) = modelo(xa)
           ObjItem.SubItems(4) = preciou(xa)
           ObjItem.SubItems(5) = cantidad(xa)
           ObjItem.SubItems(6) = IVA(xa)
           ObjItem.SubItems(7) = Margen(xa)
xa = xa + 1
Wend
End Sub

Private Sub cmdExcelSE_Click(Index As Integer)
Call Exportar2
End Sub
Sub Exportar2()


'EXPORT to EXCEL
Dim xlApp As Excel.Application
Dim xlSh As Excel.Worksheet
Dim i As Long
Dim p As Long
Set xlApp = New Excel.Application

xlApp.Visible = True
xlApp.Workbooks.Add
Set xlSh = xlApp.Workbooks(1).Worksheets(1)

  
 For i = 1 To LVP.ListItems.Count
 For p = 1 To LVP.ListItems(i).ListSubItems.Count
 If (i = 1) Then
 xlSh.Cells(i, p).Font.Bold = True
   xlSh.Cells(i, p).Value = LVP.ColumnHeaders(p).Text
   If (p = LVP.ListItems(i).ListSubItems.Count) Then
    xlSh.Cells(i, p + 1).Font.Bold = True
    xlSh.Cells(i, p + 1).Value = LVP.ColumnHeaders(p + 1).Text
    End If
   ' xlSh.Cells(i, p).Value = LV3.ListItems.Item(p)
End If
    If (p = 1) Then
    xlSh.Cells(i + 1, p).Value = LVP.ListItems(i).Text
    End If
    xlSh.Cells(i + 1, p + 1).Value = LVP.ListItems(i).ListSubItems(p).Text
       
          Next p
        Next i
 

Set xlSh = Nothing
Set xlApp = Nothing

End Sub


Private Sub cmdExcelSI_Click(Index As Integer)

End Sub

Private Sub Form_Load()
On Error Resume Next
hubocambio = False
If rs.State = adStateOpen Then rs.Close
 rs.Open "select * from Configuracion", cnn, adOpenStatic, adLockOptimistic
 cmbIVA.AddItem ("No")
    cmbIVA.AddItem (rs!IVA1 & " %")
    cmbIVA.AddItem (rs!IVA2 & " %")
  
    If (rs!IVA3 <> " ") Then
   
    cmbIVA.AddItem (rs!IVA3 & " %")
    End If
     If (rs!IVA4 <> " ") Then
   
    cmbIVA.AddItem (rs!IVA4 & " %")
    End If
     If (rs!IVA5 <> " ") Then
   
    cmbIVA.AddItem (rs!IVA5 & " %")
    End If
    rs.Close
    
    
    If (modificara = False) Then
    cmbIVA.ListIndex = 0
    cmbEst.ListIndex = 0
rs.Open "SELECT * FROM ProdXDist WHERE IdProv = '" & id & "'", cnn, adOpenStatic, adLockOptimistic
While Not (rs.EOF)
cmbCod.AddItem (rs!codigo)
cmbProd.AddItem (rs!producto)
rs.MoveNext
Wend
rs.Close
cmbCod.ListIndex = 0
cmbProd.ListIndex = 0
rs.Open "SELECT * FROM ProdXDist WHERE IdProv = '" & id & "' AND Codigo = '" & cmbCod.Text & "'", cnn, adOpenStatic, adLockOptimistic
lblMarc.Caption = rs!marca
lblMod.Caption = rs!modelo
lblPrecioU.Caption = rs!preciou
cmbMon.ListIndex = rs!Moneda - 1
cmbMon.Refresh
cmbIVA.ListIndex = rs!IVA
cmbIVA.Refresh

txtMargen.Text = rs!Margen
rs.Close


Else


rs.Open "SELECT * FROM pedidos where Id = " & id, cnn, adOpenStatic, adLockOptimistic

If (cmbEst.List(0) = rs!Estado) Then
cmbEst.ListIndex = 0
ElseIf (cmbEst.List(1) = rs!Estado) Then
cmbEst.ListIndex = 1
Else
cmbEst.ListIndex = 2
End If
rs.Close

rs.Open "SELECT id From Distribuidores where Nombre = '" & ProvN & "'", cnn, adOpenStatic, adLockOptimistic
idprov = rs!id
rs.Close



rs.Open "SELECT * FROM ProdXDist WHERE IdProv = '" & idprov & "'", cnn, adOpenStatic, adLockOptimistic
While Not (rs.EOF)
cmbCod.AddItem (rs!codigo)
cmbProd.AddItem (rs!producto)
rs.MoveNext
Wend
rs.Close
cmbCod.ListIndex = 0
cmbProd.ListIndex = 0
rs.Open "SELECT * FROM ProdXDist WHERE IdProv = '" & idprov & "' AND Codigo = '" & cmbCod.Text & "'", cnn, adOpenStatic, adLockOptimistic
lblMarc.Caption = rs!marca
lblMod.Caption = rs!modelo
lblPrecioU.Caption = rs!preciou
cmbMon.ListIndex = rs!Moneda - 1
cmbMon.Refresh
cmbIVA.ListIndex = rs!IVA
cmbIVA.Refresh
rs.Close



rs.Open "SELECT * FROM ProdXPed where IdPed = '" & id & "'", cnn, adOpenStatic, adLockOptimistic


While Not rs.EOF
 Set ObjItem = LVP.ListItems.Add(, , rs!codigo)
                
           
           ObjItem.SubItems(1) = rs!producto
           ObjItem.SubItems(2) = rs!marca
           ObjItem.SubItems(3) = rs!modelo
           ObjItem.SubItems(4) = rs!preciou
           ObjItem.SubItems(5) = rs!cantidad
           ObjItem.SubItems(6) = rs!IVA
           ObjItem.SubItems(7) = rs!Margen
rs.MoveNext
Wend
rs.Close







End If


End Sub



