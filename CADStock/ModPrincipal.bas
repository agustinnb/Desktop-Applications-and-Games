Attribute VB_Name = "ModPrincipal"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32" ()

' variables para la conexión y el recordset
''''''''''''''''''''''''''''''''''''''''''''
Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public ES As Boolean
Public ObjItem As ListItem
Public ProdId As Integer
Public pdolar As Single
Public EditarProductos As Integer
Public Prin(16) As Integer
Public Dist(6) As Integer
Public Dep(3) As Integer
Public Cli(3) As Integer
Public Ven(3) As Integer
Public IVA(5) As Single
Public firsttime As Boolean


Sub Main()
    On Error Resume Next
    Call InitCommonControls
    Err.Clear
    frmLogin.Show
End Sub

' abre
Public Sub IniciarConexion()
Set cnn = New ADODB.Connection
    With cnn
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .Properties("Jet OLEDB:Database Password") = "a17n333b110*"
        .Properties("Data Source") = App.Path & "\datos.mdb"
        .Open
    '    .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
     '          & ";Persist Security Info=False"
    End With

End Sub

Public Sub IniciarIVA()
rs.Open "SELECT * FROM Configuracion", cnn, adOpenStatic, adLockOptimistic
IVA(0) = 0
IVA(1) = rs!IVA1
IVA(2) = rs!IVA2
IVA(3) = rs!IVA3
IVA(4) = rs!IVA4
IVA(5) = rs!IVA5
rs.Close

End Sub

Public Sub CargarListView(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           
           ObjItem.SubItems(1) = rs!Nombre
           ObjItem.SubItems(2) = rs!Apellido
           ObjItem.SubItems(3) = rs!Telefono
           ObjItem.SubItems(4) = rs!Comision & " %"
           ObjItem.SubItems(5) = rs!fechadealta
           
            ' siguiente registro
            rs.MoveNext

        
        Wend
        
    End If
    
   ' Call ForeColorColumn(&H8000&, 0, FrmPrincipal.LV)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub




Public Sub CargarListView2(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
     Dim p As Integer
         Dim s As Integer
                Dim depo As String
                    Dim gMargen As Single
                        Dim gIVA As Single
    'limpia el LV
    LV.ListItems.Clear
   s = 1
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
   
           
           ObjItem.SubItems(1) = rs!producto
           ObjItem.SubItems(2) = rs!marca
           ObjItem.SubItems(3) = rs!modelo
           ObjItem.SubItems(4) = rs!Distribuidor
              If (Prin(1) = 1) Then
              Dim cantidad As Integer
           ObjItem.SubItems(5) = rs!cantidad
           cantidad = val(rs!cantidad)
           Else
            ObjItem.SubItems(5) = "--"
           End If
           If (Prin(16) = 1) Then
               ObjItem.SubItems(6) = rs!preciou
              ObjItem.SubItems(7) = Format(rs!preciou / pdolar, "#,##0.00")
             Else
              ObjItem.SubItems(6) = "--"
              ObjItem.SubItems(7) = "--"
            
             End If
              If (Prin(0) = 1) Then
             ObjItem.SubItems(8) = rs!Margen
             Else
             ObjItem.SubItems(8) = "--"
             End If
           gMargen = 1 + (rs!Margen / 100)
           gIVA = 1 + (IVA(rs!IVA) / 100)
             ObjItem.SubItems(9) = Format(rs!preciou * gMargen * gIVA, "#,##0.00")
             
             ObjItem.SubItems(10) = rs!fechadealta
           depo = rs!Deposito
           
             If depo = "" Then
              ObjItem.SubItems(11) = "No definido"
             Else
             ObjItem.SubItems(11) = rs!Deposito
            End If
            depo = ""
              ObjItem.SubItems(12) = rs!IVA
              ObjItem.SubItems(13) = rs!anotaciones
    
    If cantidad = 0 Then
    frmMain.LV2.ListItems(s).ForeColor = vbRed
    Dim l As Integer
    For l = 1 To 12
    frmMain.LV2.ListItems(s).ListSubItems(l).ForeColor = vbRed
    Next l
    ElseIf cantidad < 4 Then
     frmMain.LV2.ListItems(s).ForeColor = RGB(223, 185, 32)
    For l = 1 To 12
    frmMain.LV2.ListItems(s).ListSubItems(l).ForeColor = RGB(223, 185, 32)
    Next l
    Else
      frmMain.LV2.ListItems(s).ForeColor = vbBlue
    For l = 1 To 12
    frmMain.LV2.ListItems(s).ListSubItems(l).ForeColor = vbBlue
    Next l

    End If
    s = s + 1
    
'                         If (rs!Cantidad < 5) Then
'       frmMain.LV2.ListItems(i).ForeColor = vbRed
'    For s = 1 To frmMain.LV2.ListItems.Item(i).ListSubItems.Count
'     frmMain.LV2.ListItems.Item(i).ListSubItems(s).ForeColor = vbRed
'     Next s
  
   
'    End If
'    i = i + 1
            ' siguiente registro
            rs.MoveNext
        
        Wend
        
    
    rs.Close
        
       
      '   p = 1
      '   s = 1
  
 '   While p < LV.ListItems.Count
    
 '     rs.Open "SELECT * FROM ProdXDist", cnn, adOpenStatic, adLockOptimistic
 '   While Not rs.EOF
 ' If (LV.ListItems(p).ListSubItems(1) = rs!producto) And _
 '  (LV.ListItems(p).ListSubItems(2) = rs!marca) And _
 ' (LV.ListItems(p).ListSubItems(3) = rs!modelo) And _
 ' (LV.ListItems(p).ListSubItems(5) < rs!MinStock) Then
 '   LV.ListItems(p).ForeColor = vbRed
 '  While s < LV.ListItems.Item(p).ListSubItems.Count
 '    LV.ListItems.Item(p).ListSubItems(s).ForeColor = vbRed
 '    s = s + 1
 '    Wend
' s = 1
' End If
'1
'2 marca
'3 modelo
'  rs.MoveNext
  
'    Wend
   '   rs.Close
   '   p = p + 1

  
   ' Wend
    
    
 
        
        LV.Refresh
        
    End If
   
  '  Call ForeColorRow(vbRed, 0, frmMain.LV2)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub




Public Sub CargarListView3(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           
           ObjItem.SubItems(1) = rs!IDProd
           ObjItem.SubItems(2) = rs!producto
           ObjItem.SubItems(3) = rs!marca
           ObjItem.SubItems(4) = rs!modelo
             ObjItem.SubItems(5) = rs!cantidad
              ObjItem.SubItems(6) = rs!vendidopor
                  If (Prin(3) = 1) Then
             ObjItem.SubItems(7) = Format(rs!PrecioVenta * rs!cantidad, "#,##0.00")
             Else
              ObjItem.SubItems(7) = "--"
            
             End If
             ObjItem.SubItems(8) = rs!FechaDeBaja
             If (Prin(4) = 1) Then
              ObjItem.SubItems(9) = rs!Sena
              Else
              ObjItem.SubItems(9) = "--"
              End If
              If (Prin(2) = 1) Then
              ObjItem.SubItems(10) = rs!FDP
              Else
              ObjItem.SubItems(10) = "--"
              End If
                   ObjItem.SubItems(11) = rs!cliente
            ' siguiente registro
            rs.MoveNext
    DoEvents
        
        Wend
   
    End If
'    Call ForeColorColumn(&H8000&, 0, frmMain.LV2)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    DoEvents
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub



Public Sub CargarListView4(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           
           ObjItem.SubItems(1) = rs!Nombre
           ObjItem.SubItems(2) = rs!Apellido
           ObjItem.SubItems(3) = rs!Email
             ObjItem.SubItems(4) = rs!Telefono
              ObjItem.SubItems(5) = rs!fechadealta
          
            ' siguiente registro
            rs.MoveNext
    DoEvents
        
        Wend
        
    End If
  '  Call ForeColorColumn(&H8000&, 0, frmMain.LV2)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    
    Exit Sub
    DoEvents
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub

Public Sub CargarListView5(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           
           ObjItem.SubItems(1) = rs!monto
           ObjItem.SubItems(2) = rs!Motivo
                    
            ' siguiente registro
            rs.MoveNext
    DoEvents
        
        Wend
        
    End If
   ' Call ForeColorColumn(&H8000&, 0, frmMain.LV2)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    DoEvents
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub

Public Sub CargarListView6(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           
           ObjItem.SubItems(1) = rs!Nombre
           ObjItem.SubItems(2) = rs!Contacto
              ObjItem.SubItems(3) = rs!Telefono
                 ObjItem.SubItems(4) = rs!Direccion
                     ObjItem.SubItems(5) = rs!pdolar
            ' siguiente registro
            rs.MoveNext
    DoEvents
        
        Wend
        
    End If
   ' Call ForeColorColumn(&H8000&, 2, frmMain.LV2)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    DoEvents
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub

Public Sub CargarListView7(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           
           ObjItem.SubItems(1) = rs!Distribuidor
           ObjItem.SubItems(2) = rs!Estado
              ObjItem.SubItems(3) = rs!pdolar
                    
            ' siguiente registro
            rs.MoveNext
    DoEvents
        
        Wend
        
    End If
   ' Call ForeColorColumn(&H8000&, 2, frmMain.LV2)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    DoEvents
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub

Public Sub CargarListView8(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           ObjItem.SubItems(1) = rs!CUIT
            ObjItem.SubItems(2) = rs!rs
            
           ObjItem.SubItems(3) = rs!Nombre
           ObjItem.SubItems(4) = rs!Apellido
           ObjItem.SubItems(5) = rs!Domicilio
           ObjItem.SubItems(6) = rs!Email
           
           ObjItem.SubItems(7) = rs!Telefono
           ObjItem.SubItems(8) = rs!RI
           ObjItem.SubItems(9) = rs!DNI
           ObjItem.SubItems(10) = rs!fechadealta
           
            ' siguiente registro
            rs.MoveNext
    DoEvents
        
        Wend
        
    End If
   ' Call ForeColorColumn(&H8000&, 0, FrmPrincipal.LV)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    DoEvents
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub

Public Sub CargarListView9(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           ObjItem.SubItems(1) = rs!codigo
            ObjItem.SubItems(2) = rs!producto
            
           ObjItem.SubItems(3) = rs!marca
           ObjItem.SubItems(4) = rs!modelo

           
            ' siguiente registro
            rs.MoveNext
    DoEvents
        
        Wend
        
    End If
   ' Call ForeColorColumn(&H8000&, 0, FrmPrincipal.LV)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    DoEvents
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub



Public Sub CargarListView10(LV As ListView, rs As ADODB.Recordset)
    
    On Error GoTo ErrorSub
    
    Dim i As Integer
    'limpia el LV
    LV.ListItems.Clear
    
    ' si hay registros
    If rs.RecordCount > 0 Then
        
        ' recorre el recordset
        While Not rs.EOF
            ' añade los datos
            Set ObjItem = LV.ListItems.Add(, , rs(0))
                
           
           ObjItem.SubItems(1) = rs!Nombre
           ObjItem.SubItems(2) = rs!Direccion
              ObjItem.SubItems(3) = rs!Telefono
                ObjItem.SubItems(4) = rs!Encargado
                  ObjItem.SubItems(5) = rs!Email
            ' siguiente registro
            rs.MoveNext
    DoEvents
        
        Wend
        
    End If
   ' Call ForeColorColumn(&H8000&, 2, frmMain.LV2)
    'Call ForeColorColumn(vbRed, 6, FrmPrincipal.LV)
    DoEvents
    
    Exit Sub
    
ErrorSub:
    
    If Err.Number = 94 Then Resume Next
    
End Sub

Public Sub KillProcess(ByVal processName As String)
On Error GoTo ErrHandler
    Dim oWMI
    Dim ret
    Dim sService
    Dim oWMIServices
    Dim oWMIService
    Dim oServices
    Dim oService
    Dim servicename

    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")

    For Each oService In oServices
        servicename = _
            LCase(Trim(CStr(oService.Name) & ""))

        If InStr(1, servicename, _
            LCase(processName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If
    Next

    Set oServices = Nothing
    Set oWMI = Nothing
    Exit Sub
ErrHandler:
    Err.Clear
End Sub

Public Sub submitquery(query As String)
cnn.Execute query
End Sub

' cierra
Sub Desconectar()
    On Local Error Resume Next
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub

Public Function doLogin(User As String, pass As String) As Boolean
doLogin = False
If Trim(User) = "" Then
MsgBox "Tiene que ingresar un nombre de usuario"
Exit Function
End If
If Trim(pass) = "" Then
MsgBox "Tiene que ingresar una password"
Exit Function
End If
  If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT * FROM Perfiles WHERE Usuarios = '" & User & "'", cnn, adOpenStatic, adLockOptimistic
If rs.RecordCount = 0 Then
MsgBox "El usuario ingresado es invalido"
Exit Function
End If
If rs!Passwords <> pass Or IsNull(rs!Passwords) Then
MsgBox "El Password ingresado es invalido"
Exit Function
End If










Prin(0) = rs!Prin0
Prin(1) = rs!Prin1
Prin(2) = rs!Prin2
Prin(3) = rs!Prin3
Prin(4) = rs!Prin4
Prin(5) = rs!Prin5
Prin(6) = rs!Prin6
Prin(7) = rs!Prin7
Prin(8) = rs!Prin8
Prin(9) = rs!Prin9
Prin(10) = rs!Prin10
Prin(11) = rs!Prin11
Prin(12) = rs!Prin12
Prin(13) = rs!Prin13
Prin(14) = rs!Prin14
Prin(15) = rs!Prin15
Prin(16) = rs!Prin16

Dist(0) = rs!Dist0
Dist(1) = rs!Dist1
Dist(2) = rs!Dist2
Dist(3) = rs!Dist3
Dist(4) = rs!Dist4
Dist(5) = rs!Dist5

Dep(0) = rs!Dep0
Dep(1) = rs!Dep1
Dep(2) = rs!Dep2

Cli(0) = rs!Cli0
Cli(1) = rs!Cli1
Cli(2) = rs!Cli2

Ven(0) = rs!Ven0
Ven(1) = rs!Ven1
Ven(2) = rs!Ven2









rs.Close


















doLogin = True
End Function







    Public Function Get_Numero_Serie(ByVal s_Drive As String) As Long
         
          Dim o_Fso As Scripting.FileSystemObject
          Dim o_Drive As Drive
             
          ' Creamos un nuevo objeto de tipo Scripting FileSystemObject
          Set o_Fso = New Scripting.FileSystemObject
             
          ' Si el Drive no es un vbnullstring
          If s_Drive <> "" Then
              ' Recuperamos el Drive para poder acceder _
               en las siguientes lineas
              Set o_Drive = o_Fso.GetDrive(s_Drive)
          End If
             
          With o_Drive
                 
              ' Si está disponible
              If .IsReady Then
                  Get_Numero_Serie = Not .SerialNumber
              Else
                  MsgBox " No se puede acceder a la unidad ", vbCritical
                  Get_Numero_Serie = -1
              End If
          End With
             
          ' Eliminamos los objetos instanciados
          Set o_Drive = Nothing
          Set o_Fso = Nothing
             
      End Function

