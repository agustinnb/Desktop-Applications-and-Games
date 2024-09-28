Attribute VB_Name = "ModColorearColumnasLV"
 Option Explicit


' Sub que cambia el color de la fuente de una columna
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ForeColorColumn(Color, Columna As Integer, LV As ListView)
    Dim Item As ListItem
    Dim i As Integer
    
    ' Verifica que el control contenga items
    If LV.ListItems.Count = 0 Then
       Exit Sub
    End If
    
    ' Verifica que el Listview esté en vista de reporte
    If LV.View <> lvwReport Then
       MsgBox "el listview debe estar en vista reporte", vbQuestion
       Exit Sub
    End If
    
    ' chequea el índice de la columna que sea válido
    If (Columna + 1) > LV.ColumnHeaders.Count Then
        MsgBox "El número de columna está fuera del intervalo"
        Exit Sub
    End If
    
    ' Color de fuente Para los items
    If Columna = 0 Then
        ' recorre la lista
        For i = 1 To LV.ListItems.Count
            ' cambia el color de la fuente de la columna indicada
            LV.ListItems(i).ForeColor = Color
        Next
    ' Color de fuente para los Subitems
    Else
        ' recorre
        For i = 1 To LV.ListItems.Count
            ' cambia el color de la fuente de la columna indicada
            LV.ListItems(i).ListSubItems(Columna).ForeColor = Color
        Next
    End If
    
    ' refresca el control
    LV.Refresh
    
End Sub



Sub ForeColorRow(Color, Fila As Integer, LV As ListView)
    Dim Item As ListItem
    Dim i As Integer
    
    ' Verifica que el control contenga items
    If LV.ListItems.Count = 0 Then
       Exit Sub
    End If
    
    ' Verifica que el Listview esté en vista de reporte
    If LV.View <> lvwReport Then
       MsgBox "el listview debe estar en vista reporte", vbQuestion
       Exit Sub
    End If
    
    ' chequea el índice de la columna que sea válido
    If (Fila + 1) > LV.ColumnHeaders.Count Then
        MsgBox "El número de Fila está fuera del intervalo"
        Exit Sub
    End If
    
    ' Color de fuente Para los items
 '   If Fila = 0 Then
        ' recorre la lista
        For i = 1 To LV.ListItems.Item(Fila).ListSubItems.Count
            ' cambia el color de la fuente de la columna indicada
            LV.ListItems.Item(Fila).ListSubItems(i).ForeColor = Color
        Next
    ' Color de fuente para los Subitems
'    Else
        ' recorre
  '      For i = 1 To LV.ListItems.Count
   '         ' cambia el color de la fuente de la columna indicada
    '        LV.ListItems(i).ListSubItems(Columna).ForeColor = Color
    '    Next
   ' End If
    
    ' refresca el control
    LV.Refresh
    
End Sub
