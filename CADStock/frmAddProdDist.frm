VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmAddProdDist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agregar Productos al distribuidor"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8010
   Icon            =   "frmAddProdDist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   4800
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ChamaleonButton.ChameleonBtn cmdAddPed 
      Height          =   855
      Left            =   3120
      TabIndex        =   24
      Top             =   8280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Agregar Productos"
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
      MICON           =   "frmAddProdDist.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtMinStock 
      Height          =   285
      Left            =   1320
      TabIndex        =   20
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txtPrecioU 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtModelo 
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtMarca 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtProd 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtCod 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddTo 
      Caption         =   "Agregar a Distribuidor"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.ComboBox cmbIVA 
      Height          =   315
      ItemData        =   "frmAddProdDist.frx":05A6
      Left            =   1320
      List            =   "frmAddProdDist.frx":05A8
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdEl 
      Caption         =   "-"
      Height          =   255
      Left            =   7680
      TabIndex        =   2
      Top             =   5160
      Width           =   255
   End
   Begin VB.TextBox txtMargen 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      ItemData        =   "frmAddProdDist.frx":05AA
      Left            =   2640
      List            =   "frmAddProdDist.frx":05B4
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2880
      Width           =   855
   End
   Begin MSComctlLib.ListView LVP 
      Height          =   3015
      Left            =   480
      TabIndex        =   5
      Top             =   5160
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
      NumItems        =   9
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
         Text            =   "Moneda"
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
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "MinStock"
         Object.Width           =   2540
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelVI 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   9000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Importar Productos"
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
      MICON           =   "frmAddProdDist.frx":05C0
      PICN            =   "frmAddProdDist.frx":05DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelS 
      Height          =   735
      Index           =   1
      Left            =   1440
      TabIndex        =   23
      Top             =   9000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Exportar Productos"
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
      MICON           =   "frmAddProdDist.frx":5DB2
      PICN            =   "frmAddProdDist.frx":5DCE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Stock Minimo:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Proveedor:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblProv 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Producto:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Marca:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Modelo:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Precio Unitario:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "IVA:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Margen:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddProdDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public modificar As Boolean
Public pdolar As Single
Public id As Integer
Dim hubocambio As Boolean
Private Sub cmdAddPed_Click()


If (LVP.ListItems.Count <> 0) Then
Dim i As Integer
Dim query As String

If modificar = False Then
query = "INSERT INTO Pedidos (Distribuidor,Estado,PDolar) VALUES ('" & lblProv.Caption & "','" & cmbEst.Text & "','" & pdolar & "')"
cnn.Execute query
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT id FROM pedidos WHERE Distribuidor = '" & lblProv.Caption & "'"
rs.MoveLast
id = rs!id
rs.Close
i = 1
query = "INSERT INTO ProdXDist (IdProv, Codigo, Producto, Marca, Modelo, PrecioU, moneda, IVA, margen, MinStock) VALUES "


While i <> LVP.ListItems.Count + 1

query = query & "('" & id & "','" & LVP.ListItems(i) & "','" & LVP.ListItems(i).ListSubItems(1) & "','" & LVP.ListItems(i).ListSubItems(2) & "','" & LVP.ListItems(i).ListSubItems(3) & "','" & LVP.ListItems(i).ListSubItems(4) & "','" & LVP.ListItems(i).ListSubItems(5) & "','" & LVP.ListItems(i).ListSubItems(6) & "','" & LVP.ListItems(i).ListSubItems(7) & "','" & LVP.ListItems(i).ListSubItems(8) & "')"
cnn.Execute query


query = "INSERT INTO ProdXDist (IdProv, Codigo, Producto, Marca, Modelo, PrecioU, moneda, IVA, margen, MinStock) VALUES "
i = i + 1
Wend
Else


If rs.State = adStateOpen Then rs.Close
If (hubocambio = True) Then
query = "DELETE FROM ProdXDist WHERE IdProv = '" & id & "'"
cnn.Execute query

i = 1
query = "INSERT INTO ProdXDist (IdProv, Codigo, Producto, Marca, Modelo, PrecioU, moneda, IVA, margen, MinStock) VALUES "


While i <> LVP.ListItems.Count + 1

query = query & "('" & id & "','" & LVP.ListItems(i) & "','" & LVP.ListItems(i).ListSubItems(1) & "','" & LVP.ListItems(i).ListSubItems(2) & "','" & LVP.ListItems(i).ListSubItems(3) & "','" & LVP.ListItems(i).ListSubItems(4) & "','" & LVP.ListItems(i).ListSubItems(5) & "','" & LVP.ListItems(i).ListSubItems(6) & "','" & LVP.ListItems(i).ListSubItems(7) & "','" & LVP.ListItems(i).ListSubItems(8) & "')"
cnn.Execute query


query = "INSERT INTO ProdXDist (IdProv, Codigo, Producto, Marca, Modelo, PrecioU, moneda, IVA, margen, MinStock) VALUES "
i = i + 1

Wend
End If

End If


Unload Me
Else
MsgBox "No se pidio nada"
Exit Sub
End If





End Sub

Private Sub cmdAddTo_Click()
If (Trim(txtCod.Text) = "" Or Trim(txtProd.Text) = "" Or Trim(txtMarca.Text) = "" Or Trim(txtPrecioU.Text) = "" Or Trim(txtMargen.Text) = "" Or Trim(txtMinStock.Text) = "") Then
MsgBox "Alguno de los campos no esta completado"
Exit Sub
End If
hubocambio = True
Set ObjItem = LVP.ListItems.Add(, , txtCod.Text)
                
           
           ObjItem.SubItems(1) = txtProd.Text
           ObjItem.SubItems(2) = txtMarca.Text
           ObjItem.SubItems(3) = txtModelo.Text
         
           ObjItem.SubItems(4) = txtPrecioU.Text
   
    
             If (cmbMon.ListIndex = 0) Then
      ObjItem.SubItems(5) = 1
      Else
      ObjItem.SubItems(5) = 2
      End If
      ObjItem.SubItems(6) = cmbIVA.ListIndex
        ObjItem.SubItems(7) = txtMargen.Text
ObjItem.SubItems(8) = txtMinStock.Text


'rs.Open "SELECT id FROM Distribuidores where nombre = '" & lblProv.Caption & "'", cnn, adOpenStatic, adLockOptimistic

'rs.Close

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
Dim Moneda() As String
Dim IVA() As String
Dim Margen() As String
Dim MinStock() As String
ReDim codigo(LVP.ListItems.Count) As String
ReDim producto(LVP.ListItems.Count) As String
ReDim marca(LVP.ListItems.Count) As String
ReDim modelo(LVP.ListItems.Count) As String
ReDim preciou(LVP.ListItems.Count) As String
ReDim Moneda(LVP.ListItems.Count) As String
ReDim IVA(LVP.ListItems.Count) As String
ReDim Margen(LVP.ListItems.Count) As String
ReDim MinStock(LVP.ListItems.Count) As String
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
Moneda(xb) = LVP.ListItems(xa + 1).ListSubItems(5)
IVA(xb) = LVP.ListItems(xa + 1).ListSubItems(6)
Margen(xb) = LVP.ListItems(xa + 1).ListSubItems(7)
MinStock(xb) = LVP.ListItems(xa + 1).ListSubItems(8)
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
           ObjItem.SubItems(5) = Moneda(xa)
           ObjItem.SubItems(6) = IVA(xa)
           ObjItem.SubItems(7) = Margen(xa)
           ObjItem.SubItems(8) = MinStock(xa)
           
xa = xa + 1
Wend
End Sub

Private Sub cmdExcelS_Click(Index As Integer)
Call Exportar
End Sub

Private Sub cmdExcelVI_Click(Index As Integer)
On Error GoTo ErrorSub





CommonDialog.Filter = "Archivos de Excel|*.xls"
CommonDialog.ShowOpen
If CommonDialog.FileName <> "" Then
'dimensiones
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
Dim lngUltimaColumna As Long
Dim SQL As String
Dim x As Long
Dim Y As Long
'abrir programa Excel
Set xlApp = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en otra carpeta)
Set xlLibro = xlApp.Workbooks.Open _
(CommonDialog.FileName, True, True, , "")
 
'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(CommonDialog.FileName, True, True, , "")
Set xlHoja = xlApp.Worksheets(1)
 
'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range("A1:C10").Value

'2. Si no conoces el rango

lngUltimaColumna = 1

lngUltimaFila = _
xlHoja.Columns("A:A").Range("A1000").End(xlUp).Row
While xlHoja.Cells(1, lngUltimaColumna) <> ""
lngUltimaColumna = lngUltimaColumna + 1
Wend



'SQL = "INSERT INTO Ventas (IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,FDP,Anotaciones,Cliente) VALUES ('"



varMatriz = xlHoja.Range(xlHoja.Cells(1, 1), _
xlHoja.Cells(lngUltimaFila, lngUltimaColumna))
'utilizamos los datos...
For x = 1 To lngUltimaFila
For Y = 1 To lngUltimaColumna - 1
If (x <> 1) Then
If (Y = 1) Then
SQL = SQL & id & "','"
End If
If (Y <> lngUltimaColumna - 1) Then
SQL = SQL & varMatriz(x, Y) & "','"
Else
SQL = SQL & varMatriz(x, Y) & "')"

End If

End If
 Next
 If (x <> 1) Then
 cnn.Execute SQL
 End If
 SQL = "INSERT INTO ProdXDist (IdProv, Codigo, Producto, Marca, Modelo, PrecioU, moneda, IVA, margen, MinStock) VALUES ('"

 Next
'cerramos el archivo Excel
Set varMatriz = Nothing

xlLibro.Close SaveChanges:=False, FileName:=CommonDialog.FileName
xlApp.UserControl = False


xlApp.Quit
 
'reset variables de los objetos

Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing








LVP.ListItems.Clear
rs.Open "SELECT * FROM ProdXDist WHERE idProv = '" & id & "'", cnn, adOpenStatic, adLockOptimistic
While Not rs.EOF
 Set ObjItem = LVP.ListItems.Add(, , rs!codigo)
                
           
           ObjItem.SubItems(1) = rs!producto
           ObjItem.SubItems(2) = rs!marca
           ObjItem.SubItems(3) = rs!modelo
           ObjItem.SubItems(4) = rs!preciou
           ObjItem.SubItems(5) = rs!Moneda
           ObjItem.SubItems(6) = rs!IVA
           ObjItem.SubItems(7) = rs!Margen
           ObjItem.SubItems(8) = rs!MinStock
           rs.MoveNext
Wend


rs.Close


   
End If
Exit Sub
ErrorSub:
'cerramos el archivo Excel
Set varMatriz = Nothing
xlLibro.Close SaveChanges:=False, FileName:=CommonDialog.FileName
xlApp.Quit
 
'reset variables de los objetos

Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing

MsgBox ("Error durante la importacion del archivo")


End Sub

Private Sub Form_Load()
On Error Resume Next
cmbMon.ListIndex = 0
hubocambio = False
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM configuracion", cnn, adOpenStatic, adLockOptimistic
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
cmbIVA.ListIndex = 0
If (modificar = True) Then

rs.Open "SELECT * FROM ProdXDist WHERE idProv = '" & id & "'", cnn, adOpenStatic, adLockOptimistic
While Not rs.EOF
 Set ObjItem = LVP.ListItems.Add(, , rs!codigo)
                
           
           ObjItem.SubItems(1) = rs!producto
           ObjItem.SubItems(2) = rs!marca
           ObjItem.SubItems(3) = rs!modelo
           ObjItem.SubItems(4) = rs!preciou
           ObjItem.SubItems(5) = rs!Moneda
           ObjItem.SubItems(6) = rs!IVA
           ObjItem.SubItems(7) = rs!Margen
           ObjItem.SubItems(8) = rs!MinStock
           rs.MoveNext
Wend


rs.Close


End If






End Sub

Sub Exportar()


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



