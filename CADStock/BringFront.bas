Attribute VB_Name = "BringFront"

#If Win16 Then
    Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Integer, _
        ByVal hWndInsertAfter As Integer, ByVal X As Integer, _
        ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, _
        ByVal wFlags As Integer)
    
#Else
    Declare Function SetWindowPos Lib "User32" (ByVal _
        hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X _
        As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy _
        As Long, ByVal wFlags As Long) As Long

#End If


Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Sub Make_Window_TopMost(ByRef rfrmForm As Form, _
    ByVal vboolOn As Boolean)


#If Win32 Then
Dim lngReturn As Long
#End If

If (vboolOn = True) Then
    'make window topmost
    #If Win16 Then
    
        '16 bit
        SetWindowPos rfrmForm.hWnd, HWND_TOPMOST, _
            0, 0, 0, 0, FLAGS
    #Else
 
        '32 bit
        lngReturn = SetWindowPos(rfrmForm.hWnd, HWND_TOPMOST, _
            0, 0, 0, 0, FLAGS)
    #End If
    
Else
    'turn off topmost effect
    #If Win16 Then
    
        '16 bit
        SetWindowPos rfrmForm.hWnd, HWND_NOTOPMOST, _
        0, 0, 0, 0, FLAGS
    #Else
    
        '32 bit
        lngReturn = SetWindowPos(rfrmForm.hWnd, HWND_NOTOPMOST, _
            0, 0, 0, 0, FLAGS)
    #End If
End If

End Sub

