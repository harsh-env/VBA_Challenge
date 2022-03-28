Attribute VB_Name = "Module1"
Sub Analyze_Stock_Data()

Dim i As Long
Dim Tname, Tname1 As String
Dim OPrice, Cprice, Yearlychange, Percentchange, SVolm1, SVolm As Double
Dim WsName As String

'Dim ws As Worksheet
For Each ws In Worksheets

i = 2
j = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row ' syntax from class work

WsName = ws.Name


For i = 2 To LastRow
    OPrice = ws.Cells(i, 3)
    SVolm1 = 0
    Do
        Tname = ws.Cells(i, 1)
        Tname1 = ws.Cells(i + 1, 1)
        SVolm = ws.Cells(i, 7)
        SVolm1 = SVolm1 + SVolm
        
        i = i + 1
               
    Loop While Tname1 = Tname
    
    Cprice = ws.Cells(i - 1, 6)
    Yearlychange = Cprice - OPrice
    Percentchange = Yearlychange / OPrice
    ws.Cells(j, 12) = Tname
    ws.Cells(j, 13) = Yearlychange
    ws.Cells(j, 14).Value = Percentchange
    ws.Cells(j, 15) = SVolm1
    
    
    

j = j + 1
i = i - 1

 Next
        ws.Range("N1").EntireColumn.NumberFormat = "0.00%"  ' syntax from class work and google search
        ws.Range("L1").Value = "Ticker"
        ws.Range("M1").Value = "Yearly Change"
        ws.Range("N1").Value = "Percent Change"
        ws.Range("O1").Value = "Total Stock Volume"
        
 
        
 Next ws
 
Dim WbName As String
Dim k As Long
Dim Ychng, xmin, xmax, maxTV As Double
Dim P, T As Range


    k = 2
 
    For Each wb In Worksheets

    LastRwb = wb.Cells(Rows.Count, 13).End(xlUp).Row

    WbName = wb.Name
    
        For k = 2 To LastRwb
    
            Ychng = wb.Cells(k, 13).Value
                If Ychng < 0 Then
                    wb.Cells(k, 13).Interior.ColorIndex = 3
                Else
                    wb.Cells(k, 13).Interior.ColorIndex = 4
                End If
               
        Next k
        
        Set P = wb.Range("N2:N" & Rows.Count)
        
        xmin = Application.WorksheetFunction.Min(P) '' syntax from google search - stack overflow
        xmax = Application.WorksheetFunction.Max(P)
        
        Rmax = WorksheetFunction.Match(xmax, P, 0) + P.Row - 1 ' syntax from Mrexcel
        Rmin = WorksheetFunction.Match(xmin, P, 0) + P.Row - 1
        'Mx = WorksheetFunction.Max(Rng)
        'Rw = WorksheetFunction.Match(Mx, Rng, 0) + Rng.Row - 1
        wb.Range("S3") = xmax
        wb.Range("S4") = xmin
        wb.Range("R3") = wb.Range("L" & Rmax)
        wb.Range("R4") = wb.Range("L" & Rmin)
        
        'Rmax = wb.Range("N2:N" & LastRwb).Find(WorksheetFunction.Max(Range("N2:N" & LastRwb))).Row
        
        'Range("S3:S4").Select
        wb.Cells(3, 19).NumberFormat = "0.00%"
        wb.Cells(4, 19).NumberFormat = "0.00%"
        'Selection.
        
        Set T = wb.Range("O2:O" & Rows.Count)
        maxTV = Application.WorksheetFunction.Max(T)
        wb.Range("S5") = maxTV
        RTVmax = WorksheetFunction.Match(maxTV, T, 0) + T.Row - 1
        wb.Range("R5") = wb.Range("L" & RTVmax)
        
        wb.Range("Q3") = "Greatest % Increase "
        wb.Range("Q4") = "Greatest % Decrease "
        wb.Range("Q5") = "Greatest Total Volume "
    Next wb
      
    
    
 
 End Sub


 

