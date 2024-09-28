VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Object = "{B389CD47-E20E-4D96-A4EC-576F2B1F43BF}#1.0#0"; "Hook-Menu-2.ocx"
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendedores"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11250
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5535
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9763
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
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Apellido"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Teléfono"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Comision"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Fecha de alta"
         Object.Width           =   2540
      EndProperty
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   2520
      Top             =   5760
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   5
      Bmp:1           =   "FrmPrincipal.frx":058A
      Key:1           =   "#mnuSalir"
      Bmp:2           =   "FrmPrincipal.frx":09B2
      Key:2           =   "#mnuAgregar"
      Bmp:3           =   "FrmPrincipal.frx":0DDA
      Key:3           =   "#mnuEditarRegistro"
      Bmp:4           =   "FrmPrincipal.frx":1202
      Key:4           =   "#mnuEliminarReg"
      Bmp:5           =   "FrmPrincipal.frx":162A
      Key:5           =   "#mnuImprimir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
      MICON           =   "FrmPrincipal.frx":1A52
      PICN            =   "FrmPrincipal.frx":1A6E
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
      Left            =   120
      TabIndex        =   2
      Top             =   840
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
      MICON           =   "FrmPrincipal.frx":1E08
      PICN            =   "FrmPrincipal.frx":1E24
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
      Left            =   120
      TabIndex        =   3
      Top             =   1560
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
      MICON           =   "FrmPrincipal.frx":23BE
      PICN            =   "FrmPrincipal.frx":23DA
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
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      FOCUSR          =   0   'False
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPrincipal.frx":2774
      PICN            =   "FrmPrincipal.frx":2790
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Buscar"
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
      MICON           =   "FrmPrincipal.frx":2D2A
      PICN            =   "FrmPrincipal.frx":2D46
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "FrmPrincipal.frx":32E0
      PICN            =   "FrmPrincipal.frx":32FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdExcelVe 
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Exportar Vendedores"
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
      MICON           =   "FrmPrincipal.frx":3896
      PICN            =   "FrmPrincipal.frx":38B2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
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
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Importar Vendedores"
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
      MICON           =   "FrmPrincipal.frx":9088
      PICN            =   "FrmPrincipal.frx":90A4
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
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdExcelVe_Click(Index As Integer)
'EXPORT to EXCEL
Dim xlApp As Excel.Application
Dim xlSh As Excel.Worksheet
Dim i As Long
Dim p As Long
Set xlApp = New Excel.Application

xlApp.Visible = True
xlApp.Workbooks.Add
Set xlSh = xlApp.Workbooks(1).Worksheets(1)

  
 For i = 1 To LV.ListItems.Count
 For p = 1 To LV.ListItems(i).ListSubItems.Count
 If (i = 1) Then
 xlSh.Cells(i, p).Font.Bold = True
   xlSh.Cells(i, p).Value = LV.ColumnHeaders(p).Text
   If (p = LV.ListItems(i).ListSubItems.Count) Then
    xlSh.Cells(i, p + 1).Font.Bold = True
    xlSh.Cells(i, p + 1).Value = LV.ColumnHeaders(p + 1).Text
    End If
   ' xlSh.Cells(i, p).Value = LV3.ListItems.Item(p)
End If
    If (p = 1) Then
    xlSh.Cells(i + 1, p).Value = LV.ListItems(i).Text
    End If
    xlSh.Cells(i + 1, p + 1).Value = LV.ListItems(i).ListSubItems(p).Text
       
          Next p
        Next i
 

Set xlSh = Nothing
Set xlApp = Nothing


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
 sql = "INSERT INTO Vendedores (Nombre,Apellido,Telefono,Comision,FechaDeAlta) VALUES ('"

 Next
'cerramos el archivo Excel
xlLibro.Close SaveChanges:=False
xlApp.Quit
 
'reset variables de los objetos
Set xlHoja = Nothing
Set xlLibro = Nothing
Set xlApp = Nothing





    rs.Open "select * from Vendedores", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView(LV, rs)
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

' Botones de opción
''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOpciones_Click(Index As Integer)
    Select Case Index
        Case 0: Call Agregar
        Case 1: Call Editar
        Case 2: Call Eliminar
        Case 3: Unload Me
        Case 4: frmFilter.Show , Me
        Case 5: Call mnuImprimir_Click
    End Select
End Sub


'Abre el formulario para Editar el registro seleccionado en el ListView
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Editar()

    Dim i As Integer
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV.ListItems.Count = 0) Then
       MsgBox "No hay ningún regisro para editar", vbInformation
       Exit Sub
    End If
    If (LV.SelectedItem Is Nothing) Then
       MsgBox "Debe seleccionar previamente un registro para poder editarlo", vbInformation
       Exit Sub
    End If
    
    With FrmEdit
        ' obtiene el elemento seleccionado
        .lblID = LV.SelectedItem.Text
        For i = 1 To 3
            .Text1(i).Text = LV.SelectedItem.ListSubItems(i).Text
        Next
        .txtCom(0) = LV.SelectedItem.ListSubItems(4).Text
         .lblFecha = LV.SelectedItem.ListSubItems(5).Text
        .IdRegistro = LV.SelectedItem.Text
        .ACCION = EDITAR_REGISTRO
        
        .Show vbModal
    End With

End Sub

' Elimina el registro actual seleccionado
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub Eliminar()

    

    If (LV.ListItems.Count = 0) Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
    End If
    
    ' verifica que hay datos en el ListView y que hay uno seleccionado
    If (LV.SelectedItem Is Nothing) Then
        MsgBox "No hay registro seleccionado para eliminar", vbInformation
        Exit Sub
    End If
    
    
    With LV.SelectedItem
        ' pregunta
        If MsgBox("Se va a eliminar el registro : " & vbNewLine & _
                 String(50, "-") & vbNewLine & _
                 "ID: " & .Text & vbNewLine & _
                 "Nombre " & .ListSubItems(1).Text & vbNewLine & _
                 "Apellido: " & .ListSubItems(2).Text, _
                 vbExclamation + vbYesNo, "Eliminar") = vbYes Then
            ' Elimina
            cnn.Execute "delete from Vendedores where Id = " & .Text & ""
            ' refresca el recordset
            rs.Open "select * from Vendedores", cnn, adOpenStatic, adLockOptimistic
            ' vuelve a cargar los datos en el ListView
            Call CargarListView(LV, rs)
        End If
    End With
    If rs.State = adStateOpen Then rs.Close
End Sub


Sub Agregar()
    
    ' Acción
    FrmEdit.ACCION = AGREGAR_REGISTRO
    
    FrmEdit.lblFecha = Format(Date, "mm/dd/yyyy")
    ' Abre el Form
    FrmEdit.Show 1
End Sub

Sub Salir()
    If rs.State = adStateOpen Then rs.Close
    Unload Me
    End
End Sub


Private Sub Form_Load()
If Ven(0) = 0 Then
cmdOpciones(0).Visible = False
cmdExcelVI(0).Visible = False
End If
If Ven(1) = 0 Then
cmdOpciones(1).Visible = False
End If
If Ven(2) = 0 Then
cmdOpciones(2).Visible = False
End If
frmMain.Enabled = False
    ' llena el ListView
   If rs.State = adStateOpen Then rs.Close
      rs.Open "select * from Vendedores", cnn, adOpenStatic, adLockOptimistic
    Call CargarListView(LV, rs)
    If rs.State = adStateOpen Then rs.Close

End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub

Private Sub LV_DblClick()
If Ven(1) = 0 Then
Exit Sub
End If
    Call Editar
End Sub



Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim Item As ListItem
    
    Set Item = LV.HitTest(x, Y)
    
    If Not Item Is Nothing And Button = vbRightButton Then
       Item.Selected = True
   
    End If
End Sub

' menues
'''''''''''''''''''''''''''''

Private Sub mnuAgregar_Click()
    Call Agregar
End Sub

Private Sub mnuEditarRegistro_Click()
    Call Editar
End Sub

Private Sub mnuEliminarReg_Click()
    Call Eliminar
End Sub

Private Sub mnuImprimir_Click()
    If rs.State = adStateOpen Then rs.Close
   rs.Open "select * from Vendedores", cnn, adOpenStatic, adLockOptimistic
    Set DataReport1.DataSource = rs
    DataReport1.Show 1
       If rs.State = adStateOpen Then rs.Close
End Sub

' salir

''''''''''''''''''''''''
Private Sub mnuSalir_Click()
   Unload Me
End Sub
