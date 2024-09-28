VERSION 5.00
Begin VB.Form frmAddPed2Dep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hacer pedido"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmAddPed2Dep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Elija el deposito al que desea integrar el pedido"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmAddPed2Dep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
Dim query As String
i = 1
query = "INSERT INTO Productos (Producto, Marca, Modelo, Distribuidor, Cantidad, PrecioU,FechaDeAlta , IVA, Margen, Deposito) VALUES "

While i <> frmAddPed.LVP.ListItems.Count + 1
  query = query & "('" & frmAddPed.LVP.ListItems(i).ListSubItems(1) & "','" & frmAddPed.LVP.ListItems(i).ListSubItems(2) & "','" & frmAddPed.LVP.ListItems(i).ListSubItems(3) & "','" & frmAddPed.lblProv.Caption & "','" & frmAddPed.LVP.ListItems(i).ListSubItems(5) & "','" & frmAddPed.LVP.ListItems(i).ListSubItems(4) & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & frmAddPed.LVP.ListItems(i).ListSubItems(6) & "', '" & frmAddPed.LVP.ListItems(i).ListSubItems(7) & "','" & Combo1.Text & "')"
  
                 cnn.Execute query
     i = i + 1

                   query = "INSERT INTO Productos (Producto, Marca, Modelo, Distribuidor, Cantidad, PrecioU,FechaDeAlta , IVA, Margen, Deposito) VALUES "
 
Wend

                 rs.Open "SELECT * FROM Productos", cnn, adOpenStatic, adLockOptimistic
                 Call CargarListView2(frmMain.LV2, rs)
                 frmAddPed.Enabled = True
                 Unload Me
                 

End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim query As String

If MsgBox("Se ingresara el pedido sin deposito. ¿Desea continuar?", _
                 vbExclamation + vbYesNo, "Agregar a stock") = vbYes Then
i = 1
query = "INSERT INTO Productos (Producto, Marca, Modelo, Distribuidor, Cantidad, PrecioU,FechaDeAlta , IVA, Margen) VALUES "

While i <> frmAddPed.LVP.ListItems.Count + 1
  query = query & "('" & frmAddPed.LVP.ListItems(i).ListSubItems(1) & "','" & frmAddPed.LVP.ListItems(i).ListSubItems(2) & "','" & frmAddPed.LVP.ListItems(i).ListSubItems(3) & "','" & frmAddPed.lblProv.Caption & "','" & frmAddPed.LVP.ListItems(i).ListSubItems(5) & "','" & frmAddPed.LVP.ListItems(i).ListSubItems(4) & "', '" & Format(Date, "dd/mm/yyyy") & "', '" & frmAddPed.LVP.ListItems(i).ListSubItems(6) & "', '" & frmAddPed.LVP.ListItems(i).ListSubItems(7) & "')"
  
                 cnn.Execute query
     i = i + 1

                   query = "INSERT INTO Productos (Producto, Marca, Modelo, Distribuidor, Cantidad, PrecioU,FechaDeAlta , IVA, Margen) VALUES "
 
Wend

                 rs.Open "SELECT * FROM Productos", cnn, adOpenStatic, adLockOptimistic
                 Call CargarListView2(frmMain.LV2, rs)
End If
                 frmAddPed.Enabled = True
                 Unload Me
End Sub

Private Sub Form_Load()
rs.Open "SELECT * FROM Depositos", cnn, adOpenStatic, adLockOptimistic
While Not rs.EOF
Combo1.AddItem rs!Nombre
rs.MoveNext
Wend
Combo1.ListIndex = 0
rs.Close
End Sub
