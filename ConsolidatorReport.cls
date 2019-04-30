VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub consolidatorPool_Click()
    Call consPool
End Sub

Private Sub consPool()
    Const path As String = "F:\EI\SEA\Distribution\S167\"
    Const path2 As String = "C:\Users\sea-brandons\Desktop\" 'DEBUG
    Dim iLastRow As Integer
    Dim sFound As String
    Dim temp As Integer
    sFound = Dir(path & "YARD CHECK*.xlsm")
    'sFound = Dir(path2 & "YARD CHECK*.xlsm") 'DEBUG
    Dim bFileOpen As Boolean
    bFileOpen = IsWorkBookOpen(path & sFound)
    'bFileOpen = IsWorkBookOpen(path2 & sFound) 'DEBUG
    If bFileOpen Then
        MsgBox sFound & " is open. If you really want to send out emails then close the yard first."
        Exit Sub
    End If
    bFileOpen = IsWorkBookOpen(path & "consolidater pool.xlsx")
    If bFileOpen Then
        MsgBox "consolidater pool is open. If you really want to send out emails then close the yard first."
        Exit Sub
    End If
    
    Dim aSCAC As Variant
    Dim aCELC(4) As Integer, aHUBG(4) As Integer, aISCO(4) As Integer, aJBHU(4) As Integer
    Dim aLEGS(4) As Integer, aGLTN(4) As Integer, aSCNN(4) As Integer, aSCDS(4) As Integer
    Dim aSWFC(4) As Integer, aWTGG(4) As Integer, aSWIF(4) As Integer, aXPOG(4) As Integer
    
    Set yardWkb = Workbooks.Open(path & sFound)
    'Set yardWkb = Workbooks.Open(path2 & sFound) 'DEBUG
    Set poolWkb = Workbooks.Open(path & "consolidater pool.xlsx")
    
    iLastRow = yardWkb.Sheets("YARD").Cells(Rows.Count, 3).End(xlUp).row
    
    For i = 1 To iLastRow
    
    'If statement for CELTIC
    If yardWkb.Sheets("YARD").Cells(i, 4).Value = "CELC" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aCELC(0)
            aCELC(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aCELC(1)
            aCELC(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aCELC(2)
            aCELC(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aCELC(3)
            aCELC(3) = temp + 1
        End If
        
    'If statement for HUB
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "HUBG" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aHUBG(0)
            aHUBG(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aHUBG(1)
            aHUBG(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aHUBG(2)
            aHUBG(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aHUBG(3)
            aHUBG(3) = temp + 1
        End If
        
    'If statement for Intermodal Sales
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "ISCO" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aISCO(0)
            aISCO(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aISCO(1)
            aISCO(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aISCO(2)
            aISCO(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aISCO(3)
            aISCO(3) = temp + 1
        End If
        
    'If statement for JB Hunt
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "HJBC" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aJBHU(0)
            aJBHU(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aJBHU(1)
            aJBHU(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aJBHU(2)
            aJBHU(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aJBHU(3)
            aJBHU(3) = temp + 1
        End If
        
    'If statement for LEGEND
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "LEGS" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aLEGS(0)
            aLEGS(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aLEGS(1)
            aLEGS(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aLEGS(2)
            aLEGS(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aLEGS(3)
            aLEGS(3) = temp + 1
        End If
        
    'If statement for Performance Team
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "GLTN" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aGLTN(0)
            aGLTN(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aGLTN(1)
            aGLTN(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aGLTN(2)
            aGLTN(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aGLTN(3)
            aGLTN(3) = temp + 1
        End If
        
    'If statement for Schneider OTR
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "SCNN" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aSCNN(0)
            aSCNN(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aSCNN(1)
            aSCNN(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aSCNN(2)
            aSCNN(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aSCNN(3)
            aSCNN(3) = temp + 1
        End If
        
    'If statement for Schneider Rail
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "SCDS" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aSCDS(0)
            aSCDS(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aSCDS(1)
            aSCDS(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aSCDS(2)
            aSCDS(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aSCDS(3)
            aSCDS(3) = temp + 1
        End If
        
    'If statement for SWIFT OTR
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "SWFC" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aSWFC(0)
            aSWFC(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aSWFC(1)
            aSWFC(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aSWFC(2)
            aSWFC(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aSWFC(3)
            aSWFC(3) = temp + 1
        End If
        
    'If statement for WTGG
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "WTGG" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aWTGG(0)
            aWTGG(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aWTGG(1)
            aWTGG(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aWTGG(2)
            aWTGG(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aWTGG(3)
            aWTGG(3) = temp + 1
        End If
        
    'If statement for SWIF
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "SWIF" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aSWIF(0)
            aSWIF(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aSWIF(1)
            aSWIF(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aSWIF(2)
            aSWIF(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aSWIF(3)
            aSWIF(3) = temp + 1
        End If
        
    'If statement for XPO
    ElseIf yardWkb.Sheets("YARD").Cells(i, 4).Value = "XPOG" Then
        'If STATUS is DROP element 0 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 11).Value = "DROP" Then
            temp = aXPOG(0)
            aXPOG(0) = temp + 1
        'If STATUS has % or LOADED and PENDING or REWORK then element 1 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Text Like "*%" Or _
                (yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value = "PENDING") Or _
                yardWkb.Sheets("YARD").Cells(i, 11).Value = "REWORK" Then
            temp = aXPOG(1)
            aXPOG(1) = temp + 1
        'If STATUS has LOADED and J number element 2 increments by 1
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value = "LOADED" And yardWkb.Sheets("YARD").Cells(i, 13).Value Like "J*" Then
            temp = aXPOG(2)
            aXPOG(2) = temp + 1
        'If STATUS is NOT EMPTY and FULL display message box that the STATUS is UNRECOGNIZED
        ElseIf yardWkb.Sheets("YARD").Cells(i, 11).Value <> "EMPTY" And yardWkb.Sheets("YARD").Cells(i, 11).Value <> "FULL" Then
            MsgBox "UNRECOGNIZED STATUS" & vbCrLf & vbCrLf & yardWkb.Sheets("YARD").Cells(i, 3) & _
                " in " & yardWkb.Sheets("YARD").Cells(i, 2) & " status " & yardWkb.Sheets("YARD").Cells(i, 11)
        End If
        
        'If Trailer has been in yard for 1 day or more then element 3 increments by 1
        If yardWkb.Sheets("YARD").Cells(i, 15) <> "" And Date - yardWkb.Sheets("YARD").Cells(i, 15) >= 1 Then
            temp = aXPOG(3)
            aXPOG(3) = temp + 1
        End If
    End If
    Next i
    
    aSCAC = Array(aCELC, aHUBG, aISCO, aJBHU, aLEGS, aGLTN, aSCNN, aSCDS, aSWFC, aWTGG, aSWIF, aXPOG)
    
    '______________________________________________________________________________________________________
    '               INPUT THE ARRAY INFORMATION INTO THE CONSOLIDATOR POOL EXCEL SHEET
    For i = 8 To 19
        For j = 0 To 3
            poolWkb.Sheets("Daily Worksheet").Cells(i, j + 6) = aSCAC(i - 8)(j)
        Next j
    Next i
End Sub

Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function
