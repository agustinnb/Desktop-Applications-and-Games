VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form frmBusq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de ventas"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14100
   Icon            =   "frmBusq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheckEC 
      Caption         =   "Restar egresos de caja"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   2775
   End
   Begin VB.ComboBox cmbFDP 
      Height          =   315
      ItemData        =   "frmBusq.frx":058A
      Left            =   1440
      List            =   "frmBusq.frx":0591
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cmbBP 
      Height          =   315
      ItemData        =   "frmBusq.frx":059C
      Left            =   5040
      List            =   "frmBusq.frx":05BB
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtBP 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox cmbFDB 
      Height          =   315
      ItemData        =   "frmBusq.frx":0615
      Left            =   1440
      List            =   "frmBusq.frx":061C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox cmbVP 
      Height          =   315
      ItemData        =   "frmBusq.frx":0627
      Left            =   7440
      List            =   "frmBusq.frx":062E
      TabIndex        =   2
      Text            =   "cmbVP"
      Top             =   240
      Width           =   1575
   End
   Begin MSComctlLib.ListView LV3 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   14055
      _ExtentX        =   24791
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID Producto"
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
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cantidad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Vendido Por"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Precio de Venta"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Fecha de Baja"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Seña"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Forma de Pago"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Cliente"
         Object.Width           =   2540
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdOpciones 
      Height          =   375
      Index           =   3
      Left            =   12600
      TabIndex        =   12
      Top             =   4200
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
      MICON           =   "frmBusq.frx":0639
      PICN            =   "frmBusq.frx":0655
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblCom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   5055
   End
   Begin VB.Label lblFDP 
      Alignment       =   1  'Right Justify
      Caption         =   "Forma de Pago:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblEf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha de baja:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblVP 
      Alignment       =   1  'Right Justify
      Caption         =   "Vendido por:"
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmBusq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vendedores() As String

Private Sub CheckEC_Click()
Call sumatoria
End Sub

Private Sub cmbBP_Click()
Call Buscar
End Sub

Private Sub cmbFDB_Click()
Call Buscar
End Sub

Private Sub cmbFDP_Click()
Call Buscar
End Sub

Private Sub cmbVP_Click()
Call Buscar
End Sub

Private Sub cmdOpciones_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
lblCom.Caption = ""
Dim X As Integer
Dim s As Integer
Dim fecha() As String
frmMain.Enabled = False
If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM Ventas", cnn, adOpenStatic, adLockOptimistic
Call CargarListView3(LV3, rs)
rs.MoveFirst
X = 0
ReDim fecha(X) As String
While (rs.EOF = False)
fecha(X) = rs!FechaDeBaja
X = X + 1
ReDim Preserve fecha(X) As String
rs.MoveNext
Wend
For s = 0 To X - 1
If (s = 0) Then
cmbFDB.AddItem fecha(s)
Else
If (fecha(s) <> fecha(s - 1)) Then
cmbFDB.AddItem fecha(s)
End If
End If
Next





rs.MoveFirst
X = 0
ReDim vendedores(X) As String
While (rs.EOF = False)
vendedores(X) = rs!vendidopor
X = X + 1
ReDim Preserve vendedores(X) As String
rs.MoveNext
Wend
Call Ordenar
For s = 0 To X - 1
If (s = 0) Then
If (vendedores(s) <> "") Then
cmbVP.AddItem vendedores(s)
End If
Else
If (vendedores(s) <> vendedores(s - 1)) Then
If (vendedores(s) <> "") Then
cmbVP.AddItem vendedores(s)
End If
End If
End If
Next

rs.Close

cmbFDP.AddItem "Efectivo"
cmbFDP.AddItem "Tarjeta"
cmbFDP.AddItem "Cheque"
cmbFDP.AddItem "Debito"
cmbFDP.AddItem "Otro"




cmbFDP.ListIndex = 0
cmbFDB.ListIndex = 0
cmbBP.ListIndex = 0
cmbVP.ListIndex = 0
End Sub



 Private Sub Ordenar()

     Dim iMin    As Long
     Dim iMax    As Long
     Dim Vectemp As String                   ' -- variable temporal
     Dim Pos     As Long
     Dim i       As Long

    iMin = LBound(vendedores)
     iMax = UBound(vendedores)

     While iMax > iMin
         Pos = iMin
         For i = iMin To iMax - 1
If vendedores(i) > vendedores(i + 1) Then
             Vectemp = vendedores(i + 1)
             vendedores(i + 1) = vendedores(i)
             vendedores(i) = Vectemp
             Pos = i
End If
         Next i
         iMax = Pos
     Wend
 End Sub


Private Sub Buscar()
 On Error Resume Next

Dim vendidopor As Boolean
Dim FDB As Boolean
Dim BP As Boolean
Dim FDP As Boolean
Dim query As String
vendidopor = False
FDB = False
BP = False
FDP = False
If (cmbVP.Text = "Todos") Then
vendidopor = False
Else
vendidopor = True
End If
If (cmbFDB.Text = "Todos") Then
FDB = False
Else
FDB = True
End If
If (cmbBP.Text = "Todos") Then
BP = False
Else
BP = True
End If
If (cmbFDP.Text = "Todos") Then
FDP = False
Else
FDP = True
End If



query = "SELECT * FROM Ventas"
If (vendidopor = True) Then
query = query & " WHERE VendidoPor LIKE '%" & cmbVP.Text & "%'"
End If


If (FDB = True) Then
If (vendidopor = True) Then
query = query & " AND FechaDeBaja LIKE '%" & cmbFDB.Text & "%'"
Else
query = query & " WHERE FechaDeBaja LIKE '%" & cmbFDB.Text & "%'"
End If
End If

If (FDP = True) Then
If (FDB = True) Or (vendidopor = True) Then
query = query & " AND FDP LIKE '%" & cmbFDP.Text & "%'"
Else
query = query & " WHERE FDP LIKE '%" & cmbFDP.Text & "%'"
End If
End If


If (Trim(txtBP.Text) <> "") Then
If (vendidopor = False) And (FDB = False) And (FDP = False) Then
If BP = False Then
query = query & " WHERE Id LIKE '" & txtBP.Text & "' OR IdProd LIKE '" & txtBP.Text & "' " & _
"OR Producto LIKE '%" & txtBP.Text & "%' OR Modelo LIKE '%" & txtBP.Text & "%' OR Marca LIKE '%" & txtBP.Text & "%' OR Cantidad LIKE '" & txtBP.Text & "%' OR (PrecioVenta*Cantidad) LIKE '" & txtBP.Text & "%' OR Clientes LIKE '%" & txtBP.Text & "%'"
Else
If cmbBP.ListIndex = 1 Then
query = query & " WHERE Id LIKE '" & txtBP.Text & "'"
End If
If cmbBP.ListIndex = 2 Then
query = query & " WHERE IdProd LIKE '" & txtBP.Text & "'"
End If
If cmbBP.ListIndex = 3 Then
query = query & " WHERE Producto LIKE '%" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 4 Then
query = query & " WHERE Marca LIKE '%" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 5 Then
query = query & " WHERE Modelo LIKE '%" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 6 Then
query = query & " WHERE Cantidad LIKE '" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 7 Then
query = query & " WHERE (PrecioVenta*Cantidad) LIKE '" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 8 Then
query = query & " WHERE Clientes LIKE '%" & txtBP.Text & "%'"
End If


End If




Else




If BP = False Then
query = query & " AND Id LIKE '" & txtBP.Text & "' OR IdProd LIKE '" & txtBP.Text & "' " & _
"OR Producto LIKE '%" & txtBP.Text & "%' OR Modelo LIKE '%" & txtBP.Text & "%' OR Marca LIKE '%" & txtBP.Text & "%' OR Cantidad LIKE '" & txtBP.Text & "%' OR (PrecioVenta*Cantidad) LIKE '" & txtBP.Text & "%' OR Clientes LIKE '%" & txtBP.Text & "%'"
Else


If cmbBP.ListIndex = 1 Then
query = query & " AND Id LIKE '" & txtBP.Text & "'"
End If
If cmbBP.ListIndex = 2 Then
query = query & " AND IdProd LIKE '" & txtBP.Text & "'"
End If
If cmbBP.ListIndex = 3 Then
query = query & " AND Producto LIKE '%" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 4 Then
query = query & " AND Marca LIKE '%" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 5 Then
query = query & " AND Modelo LIKE '%" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 6 Then
query = query & " AND Cantidad LIKE '" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 7 Then
query = query & " AND (PrecioVenta*Cantidad) LIKE '" & txtBP.Text & "%'"
End If
If cmbBP.ListIndex = 8 Then
query = query & " AND Clientes LIKE '%" & txtBP.Text & "%'"
End If

End If



End If
End If

rs.Open query, cnn, adOpenStatic, adLockOptimistic
Call CargarListView3(LV3, rs)
rs.Close


Call sumatoria

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub




Private Sub txtBP_Change()
Call Buscar
End Sub


Private Sub sumatoria()
On Error Resume Next
Dim Y As Integer
Dim sumar As Long
Dim query As String
Dim com As Single
Dim egresos As Single
egresos = 0
sumar = 0
 For Y = 1 To LV3.ListItems.Count
 sumar = sumar + LV3.ListItems(Y).ListSubItems(7).Text
  Next

  
  
If (CheckEC.Value = 1) Then

rs.Open "SELECT * FROM Egresos", cnn, adOpenStatic, adLockOptimistic
While Not (rs.EOF)
egresos = egresos + rs!monto
rs.MoveNext
Wend
rs.Close
End If
  
  
  
  
  
  
  
  
  If (CheckEC.Value = 0) Then
lblEf.Caption = "Sumatoria de ventas listadas: " & Format(sumar, "#,##0.00")
  
ElseIf (CheckEC.Value = 1) Then
  lblEf.Caption = "Sumatoria de ventas listadas: " & Format(sumar - egresos, "#,##0.00")

End If









If (cmbVP <> "Todos") Then
Dim vendedores() As String
p = CountWords(cmbVP.Text)
vendedores = Split(cmbVP.Text, " ")

query = "SELECT * FROM Vendedores WHERE "
For Y = 0 To p - 1
If (Y = 0) Then
query = query & "Nombre LIKE '%" & vendedores(Y) & "%' OR Apellido LIKE '%" & vendedores(Y) & "%'"
Else
query = query & " OR Nombre LIKE '%" & vendedores(Y) & "%' OR Apellido LIKE '%" & vendedores(Y) & "%'"
End If
Next

If rs.State = adStateOpen Then rs.Close
rs.Open query, cnn, adOpenStatic, adLockOptimistic
com = rs!Comision
rs.Close


lblCom.Caption = "Comision del vendedor: " & (Format((sumar * com) / 100, "#,##0.00"))
Else
lblCom.Caption = ""
End If



End Sub


Private Function CountWords(strText As String) As Long
    Dim X As Long
    Dim words As Long
    Dim lenStr As Long
    Dim lenPrevWord As Long
    Dim currentAscii As Integer
    
    ' Number of chars in string
    lenStr = Len(strText)
    ' Number of chars in the last word processed
    lenPrevWord = 0
    
    For X = 1 To lenStr
        ' ASCII value of the char being processed
        currentAscii = Asc(Mid$(strText, X, 1))
        Select Case currentAscii
            ' If current char is space (ASCII value 32)...
            Case 32
                ' ...and we have processed at least a char before
                ' then we have found a word
                If lenPrevWord > 0 Then
                    words = words + 1
                    ' Clear count of characters in previous word
                    lenPrevWord = 0
                End If
            Case Else
                ' If current char is not a space, then add 1 to the
                ' count of chars of the current word.
                lenPrevWord = lenPrevWord + 1
        End Select

        ' Test if last char is other than a space.
        ' If so then there is a word.
        If X = lenStr And currentAscii <> 32 Then
            words = words + 1
        End If
    Next X
    
    CountWords = words
End Function
