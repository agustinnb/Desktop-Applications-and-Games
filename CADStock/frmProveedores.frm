VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmProveedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Distribuidores"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12045
   Icon            =   "frmProveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ChamaleonButton.ChameleonBtn cmdHP 
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Hacer Pedido"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProveedores.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidos"
      Height          =   4935
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   6255
      Begin MSComctlLib.ListView LVP 
         Height          =   4575
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8070
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Distribuidor"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Precio Dolar"
            Object.Width           =   3175
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Distribuidores"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin MSComctlLib.ListView LVD 
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   8070
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Distribuidores"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Contacto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Telefono"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Direccion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Precio Dolar"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Nuevo"
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
      FCOL            =   16777152
      FCOLO           =   16777152
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProveedores.frx":05A6
      PICN            =   "frmProveedores.frx":05C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   615
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmProveedores.frx":095C
      PICN            =   "frmProveedores.frx":0978
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   615
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Eliminar"
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
      MICON           =   "frmProveedores.frx":0F12
      PICN            =   "frmProveedores.frx":0F2E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdElPed 
      Height          =   615
      Left            =   7320
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Cancelar Pedido"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProveedores.frx":12C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdMP 
      Height          =   615
      Left            =   9120
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Modificar Pedido"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmProveedores.frx":12E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelVI 
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Importar Distribuidores"
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
      MICON           =   "frmProveedores.frx":1300
      PICN            =   "frmProveedores.frx":131C
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
      Left            =   1560
      TabIndex        =   11
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Exportar Distribuidores"
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
      MICON           =   "frmProveedores.frx":6AF2
      PICN            =   "frmProveedores.frx":6B0E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelSE 
      Height          =   735
      Index           =   0
      Left            =   5760
      TabIndex        =   12
      Top             =   6120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Exportar Pedidos"
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
      MICON           =   "frmProveedores.frx":C2E4
      PICN            =   "frmProveedores.frx":C300
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdElPed_Click()
Call EliminarPed
End Sub



Private Sub cmdExcelS_Click(Index As Integer)
Call Exportar
End Sub

Private Sub cmdExcelSE_Click(Index As Integer)
Call Exportar2
End Sub

Private Sub cmdExcelVI_Click(Index As Integer)

 On Error GoTo ErrorSub
CommonDialog1.Filter = "Archivos de Excel|*.xls"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
'dimensiones
Dim xlApp As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja As Excel.Worksheet
Dim varMatriz As Variant
Dim lngUltimaFila As Long
Dim lngUltimaColumna As Long
Dim sql As String
Dim x As Long
Dim Y As Long
'abrir programa Excel
Set xlApp = New Excel.Application
'xl.Visible = True

'abrir el archivo Excel
'(archivo en otra carpeta)
Set xlLibro = xlApp.Workbooks.Open _
(CommonDialog1.FileName, True, True, , "")
 
'abrir el archivo Excel
'(archivo en la misma carpeta)
Set xlLibro = xlApp.Workbooks.Open(CommonDialog1.FileName, True, True, , "")
Set xlHoja = xlApp.Worksheets(1)
 
'1. Si conoces el rango a leer
'varMatriz = xlHoja.Range("A1:C10").Value

'2. Si no conoces el rango

lngUltimaColumna = 1

lngUltimaFila = _
xlHoja.Columns("A:A").Range("A65536").End(xlUp).Row
While xlHoja.Cells(1, lngUltimaColumna) <> ""
lngUltimaColumna = lngUltimaColumna + 1
Wend



'SQL = "INSERT INTO Ventas (IDProd,Producto,Marca,Modelo,Cantidad,VendidoPor,PrecioVenta,FechaDeBaja,FDP,Anotaciones,Cliente) VALUES ('"

sql = "INSERT INTO Distribuidores (Nombre, Contacto, Telefono, Direccion, PDolar) VALUES ('"


varMatriz = xlHoja.Range(xlHoja.Cells(1, 1), _
xlHoja.Cells(lngUltimaFila, lngUltimaColumna))
'utilizamos los datos...
For x = 1 To lngUltimaFila
For Y = 1 To lngUltimaColumna - 1
If (x <> 1) Then
If (Y <> 1) Then
If (Y <> lngUltimaColumna - 1) Then
sql = sql & varMatriz(x, Y) & "','"
Else
sql = sql & varMatriz(x, Y) & "')"

End If
End If
End If
 Next
 If (x <> 1) Then
 cnn.Execute sql
 End If
 'query =

 sql = "INSERT INTO Distribuidores (Nombre, Contacto, Telefono, Direccion, PDolar) VALUES ('"

 Next
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing



rs.Open "SELECT * FROM Distribuidores", cnn, adOpenStatic, adLockOptimistic
Call CargarListView6(LVD, rs)
rs.Close

   
End If
Exit Sub
ErrorSub:
MsgBox ("Error durante la importacion del archivo")
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing


End Sub

Private Sub cmdHP_Click()
frmAddPed.modificara = False
   If (LVD.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LVD.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
With frmAddPed
  .id = LVD.SelectedItem.Text
.lblProv = LVD.SelectedItem.ListSubItems(1).Text
.pdolar = LVD.SelectedItem.ListSubItems(5).Text
.Show vbModal
End With

End Sub

Private Sub cmdMP_Click()
frmAddPed.modificara = True
  If (LVP.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LVP.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
With frmAddPed
  .id = LVP.SelectedItem.Text
.lblProv = LVP.SelectedItem.ListSubItems(1).Text
.pdolar = LVP.SelectedItem.ListSubItems(3).Text
.Show vbModal
End With
End Sub

Private Sub cmdOpciones_Click(Index As Integer)
If (Index = 0) Then
frmAddProv.modificar = False
frmAddProv.Show

End If
If (Index = 1) Then
Call Editar

End If
If (Index = 2) Then
Call Eliminar
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
If Dist(0) = 0 Then
cmdOpciones(0).Visible = False
End If
If Dist(1) = 0 Then
cmdOpciones(1).Visible = False
End If
If Dist(2) = 0 Then
cmdOpciones(2).Visible = False
End If
If Dist(3) = 0 Then
cmdHP.Visible = False
End If
If Dist(4) = 0 Then
cmdMP.Visible = False
End If
If Dist(5) = 0 Then
cmdElPed.Visible = False
End If


If (rs.State = adStateOpen) Then rs.Close
rs.Open "SELECT * FROM Distribuidores", cnn, adOpenStatic, adLockOptimistic
Call CargarListView6(LVD, rs)
rs.Close
rs.Open "SELECT * FROM Pedidos", cnn, adOpenStatic, adLockOptimistic
Call CargarListView7(LVP, rs)
rs.Close
frmAddPed.ProvN = LVP.ListItems(1).ListSubItems(1).Text
End Sub



Private Sub Form_Unload(Cancel As Integer)


frmMain.Enabled = True
End Sub

Private Sub LVD_DblClick()
If Dist(1) = 0 Then
Exit Sub
End If
Call Editar

End Sub
Private Sub Editar()

    Dim i As Integer
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LVD.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LVD.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
    With frmAddProv
        ' obtiene el elemento seleccionado
        .lblID.Caption = LVD.SelectedItem.Text
      
            .txtAddProv = LVD.SelectedItem.ListSubItems(1).Text
        .txtCon = LVD.SelectedItem.ListSubItems(2).Text
         .txtTel = LVD.SelectedItem.ListSubItems(3).Text
        .txtDir = LVD.SelectedItem.ListSubItems(4).Text
        .txtPDolar = LVD.SelectedItem.ListSubItems(5).Text
        .modificar = True
        
        .Show vbModal
    End With

End Sub










' Elimina el registro actual seleccionado
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub Eliminar()

    

    If (LVD.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LVD.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LVD.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "Nombre " & .ListSubItems(1).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Distribuidores where Id = " & .Text & ""
            ' refresca el recordset
            rs.Open "select * from Distribuidores", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView6(LVD, rs)
        End If
    End With
    If rs.State = adStateOpen Then rs.Close
End Sub

Sub EliminarPed()

    If (LVP.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LVP.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LVP.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el pedido : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "Distribuidor " & .ListSubItems(1).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Pedidos where Id = " & .Text & ""
            cnn.Execute "delete from ProdXPed where IdPed = '" & .Text & "'"
            ' refresca el recordset
            rs.Open "select * from Pedidos", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView7(LVP, rs)
        End If
    End With
    If rs.State = adStateOpen Then rs.Close
End Sub



Private Sub LVP_Click()
  If (LVP.ListItems.Count <> 0) And Not (LVP.SelectedItem Is Nothing) Then
 frmAddPed.ProvN = LVP.SelectedItem.ListSubItems(1)
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

  
 For i = 1 To LVD.ListItems.Count
 For p = 1 To LVD.ListItems(i).ListSubItems.Count
 If (i = 1) Then
 xlSh.Cells(i, p).Font.Bold = True
   xlSh.Cells(i, p).Value = LVD.ColumnHeaders(p).Text
   If (p = LVD.ListItems(i).ListSubItems.Count) Then
    xlSh.Cells(i, p + 1).Font.Bold = True
    xlSh.Cells(i, p + 1).Value = LVD.ColumnHeaders(p + 1).Text
    End If
   ' xlSh.Cells(i, p).Value = LV3.ListItems.Item(p)
End If
    If (p = 1) Then
    xlSh.Cells(i + 1, p).Value = LVD.ListItems(i).Text
    End If
    xlSh.Cells(i + 1, p + 1).Value = LVD.ListItems(i).ListSubItems(p).Text
       
          Next p
        Next i
 

Set xlSh = Nothing
Set xlApp = Nothing

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
