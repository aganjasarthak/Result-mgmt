Sub percent()
    Dim s As Double
    Dim lastrow As Long  ' Use Long instead of Double for row counts
    Dim i As Double, j As Double
    Dim c As Double
    Dim rng As Range
    Dim cell As Range
    Dim colLimit As Integer
    Dim targetcell As Range
    Dim target As Range
    Dim csum As Double
    Dim i1 As Long
    Dim csum1 As Double
    
    colLimit = 33
    'fff
'    If Sheet2.Cells(13, 4).Value = "Yes" Or Sheet2.Cells(13, 4).Value = "No" Then
    If Sheet2.Cells(9, 4).Value > 0 Then
    If Sheet2.Cells(15, 4).Value > 0 Then

    
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Input")
     c = 3 + ws.Cells(1, 3).Value
    
    lastrow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
     Dim col77 As Range
    Dim conditionRange As Range
    Dim checkCell As Range
    
     
    
     Dim z As Double
    Dim targetColumn As String
    targetColumn = "G"  ' Starting column for insertion
    Dim x As Double
    x = 2
    If Sheet5.Cells(13, 1) < 1 Then
    
    For z = 1 To (Sheet2.Cells(15, 4).Value - 1)
        
        Dim currentColumn As String
        currentColumn = Split(Cells(1, Columns(targetColumn).Column + z - 1).Address, "$")(1)
        
        
        Columns(currentColumn & ":" & currentColumn).Insert Shift:=xlToLeft
        
        
       
        Columns(currentColumn & ":" & currentColumn).ColumnWidth = 5
        
        ws.Cells(2, (x + 5)) = x
        x = x + 1
        
    Next z
 End If

      ws.Cells.EntireRow.Hidden = False
ws.Cells.EntireColumn.Hidden = False
      
      
      Dim k As Double
      k = 2
      Dim zx As Double


     
 If Sheet5.Cells(13, 1) < 2 Then
 

                            For zx = 1 To (Sheet2.Cells(17, 4) - 1)
                                ws.Cells(2, (zx + Sheet2.Cells(15, 4) + 7)).EntireColumn.Insert
                                ws.Cells(2, (zx + Sheet2.Cells(15, 4) + 7)) = k

                                k = k + 1
                            Next zx

End If

   'two part and single part
   If Sheet5.Cells(2, 1).Value = 0 Then
   'single
    
      
                            For zx = (7 + Sheet2.Cells(15, 4)) To (7 + Sheet2.Cells(15, 4) + Sheet2.Cells(17, 4) + 1)
                                ws.Cells(2, zx).EntireColumn.Hidden = True
                            Next zx
                            
                   
                
                      
    End If

    If Sheet5.Cells(26, 5).Value >= 1 Then
   If Sheet5.Cells(2, 1).Value = 0 Then
   
                                ws.Cells(2, (7 + Sheet2.Cells(15, 4) + Sheet2.Cells(17, 4) + 1)).EntireColumn.Hidden = False
                                Sheet2.Cells(17, 4) = 0
                                

End If

        End If

                           
    
    Dim a As Double
    Dim b As Double
    
    a = Sheet2.Cells(15, 4)
    b = Sheet2.Cells(17, 4)
    Dim newsum As Double
    newsum = 0
    
 Sheet5.Cells(27, 1).Value = 1
 Dim totalmarkssum As Double
 Dim newsum1 As Double
 
 
'  TABBLE FOR INPUT SHEET
 
    Dim d12 As Double
    Dim rng2 As Range
    Dim colLimit12 As Double
    Dim cell12 As Range
    Dim endCol122 As Double
    

    ' Set the worksheet, replace "SheetName" with your worksheet name
    
    ' Initialize d12 with a value, replace 'c' with your specific value or calculation
    d12 = c
    
    ' Determine colLimit12 based on a condition
    If Sheet5.Cells(2, 1).Value = 1 Then
        colLimit12 = a + b + 9
    Else
        colLimit12 = 10 + a
    End If
    
    ' Set the range from A2 to the last row defined by d12
    Set rng2 = ws.Range("A2:A" & d12)
    
    ' Loop through each cell in the range
    For Each cell12 In rng2
        ' Calculate the end column based on the minimum of colLimit12 and the last non-empty column
        endCol122 = colLimit12
        
        
        ' Check if the end column is greater than or equal to 1
        If endCol122 >= 1 Then
            ' Apply borders to the range from the current cell to the determined end column
            With ws.Range(cell12, ws.Cells(cell12.row, endCol122))
                .Borders.LineStyle = xlContinuous
                .Borders.ColorIndex = xlAutomatic ' Optional: Set the border color
                .Borders.Weight = xlThin
            End With
        End If
    Next cell12
    
If Sheet5.Cells(29, 8).Value > 0 Then

   If Sheet5.Cells(2, 1).Value = 0 Then
   
    For i = 3 To c
     newsum = 0
     
       For j = 6 To (6 + a - 1)

       newsum = newsum + ws.Cells(i, j)
    Next j
     ws.Cells(i, (6 + a)) = newsum
     ws.Cells(i, (6 + a + 3)) = newsum
     totalmarkssum = ws.Cells(3, (6 + a))
     
     If totalmarkssum > 0 Then
     ws.Cells(i, (6 + a + 4)) = newsum / totalmarkssum
     Else
     End If
     
     Sheet3.Cells(i, 6).Value = ws.Cells(i, (6 + a + 4))
    Next i
    
    
    
    Else
          
          For i = 3 To c
     newsum = 0
     
       For j = 6 To (6 + a - 1)

       newsum = newsum + ws.Cells(i, j)
    Next j
     ws.Cells(i, (6 + a)) = newsum
     newsum1 = 0
       For j = (6 + a + 1) To (6 + a + 1 + b - 1)

       newsum1 = newsum1 + ws.Cells(i, j)
    Next j
     ws.Cells(i, (6 + a + 1 + b - 1 + 1)) = newsum1
     ws.Cells(i, (6 + a + 1 + b - 1 + 1 + 1)) = newsum1 + newsum
     totalmarkssum = ws.Cells(3, (6 + a + 1 + b - 1 + 1 + 1))
     
     If totalmarkssum > 0 Then
     ws.Cells(i, (6 + a + 1 + b - 1 + 1 + 1 + 1)) = ws.Cells(i, (6 + a + 1 + b - 1 + 1 + 1)) / totalmarkssum
     Else
     
     End If
     
     Sheet3.Cells(i, 6).Value = ws.Cells(i, (6 + a + 1 + b - 1 + 1 + 1 + 1))
     
     
    Next i
    
End If
End If ' for summing parts(29,8)



'for s.n's
Dim kk As Double
kk = 1
For i = 4 To c
ws.Cells(i, 1) = kk
kk = kk + 1


 Next i
 
If Sheet5.Cells(29, 8).Value > 0 Then

'copying 2 grade sheet (sheet3)
For i = 2 To c
                 If i = 3 Then
                  
                 
                 Else
                    
                
                
                   For j = 1 To 5
                   Sheet3.Cells(i, j) = ws.Cells(i, j)
                   
                   Next j
                   
                   
                   
                
                 End If
                 Next i
                 
    
    Sheet3.Cells(2, 6) = "Percent"
    Sheet3.Cells(2, 7) = "Pass/Check"
    Sheet3.Cells(2, 8) = "Grades"
    
    
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("Grade")


    Dim rng1 As Range
    Dim cell1 As Range
    Dim colLimit1 As Integer
    Dim endCol1 As Integer
    Dim d1 As Double


    d1 = c

    colLimit1 = 8 ' Column limit up to which borders will be applied

    ' Define the range to loop through. Example: from row 2 to row d1
    Set rng1 = ws1.Range("A2:A" & d1)

    For Each cell1 In rng1
        ' Find the last used column in the row of the current cell
        endCol1 = Application.WorksheetFunction.Min(colLimit1, ws1.Cells(cell1.row, ws1.Columns.count).End(xlToLeft).Column)

        If endCol1 >= 1 Then
            ' Apply borders to the range from the current cell to column colLimit1
            With ws1.Range(cell1, ws1.Cells(cell1.row, colLimit1))
                .Borders.LineStyle = xlContinuous
                .Borders.ColorIndex = xlAutomatic ' Optional: Set the border color
                .Borders.Weight = xlThin
            End With
        End If
    Next cell1
    
    
    
    Dim grade1 As Double
        Dim grade4 As Double
        Dim diff As Double
        Dim diffp As Double

        grade1 = Sheet2.Cells(13, 7).Value / totalmarkssum
        grade4 = Sheet2.Cells(15, 7).Value / totalmarkssum
    
    If Sheet2.Cells(13, 7) > 0 Then
    If Sheet2.Cells(15, 7) > 0 Then
    
    
    
    
    For i = 4 To c
      If Sheet3.Cells(i, 6).Value < grade4 Then
          Sheet3.Cells(i, 7).Value = "NOT PASS"
          Sheet3.Cells(i, 8).Value = 5
          Sheet3.Cells(i, 8).Interior.Color = RGB(255, 204, 204)

          ElseIf Sheet3.Cells(i, 5) = "Not Present" Then

            Sheet3.Cells(i, 8).Value = "N/A"
            Sheet3.Cells(i, 8).Interior.Color = RGB(135, 206, 235)

            ElseIf Sheet3.Cells(i, 6).Value >= grade4 Then
            Sheet3.Cells(i, 7).Value = "PASS"
            Sheet3.Cells(i, 8).Value = 4

End If
     
     If Sheet3.Cells(i, 5) = "Not Present" Then
            Sheet3.Cells(i, 6).Value = ""
            Sheet3.Cells(i, 7).Value = "N/A"
            Sheet3.Cells(i, 7).Interior.Color = RGB(135, 206, 235)
            Sheet3.Cells(i, 8).Value = "N/A"
            Sheet3.Cells(i, 8).Interior.Color = RGB(135, 206, 235)
            End If
            
 Next i
End If
End If

End If '29,8 part 2
 '

 
 'INTERPOLATION OF GRADES
 If Sheet2.Cells(13, 7) > 0 Then
    If Sheet2.Cells(15, 7) > 0 Then
 
 
 
 For i = 4 To c


        grade1 = Sheet2.Cells(13, 7).Value / totalmarkssum
        grade4 = Sheet2.Cells(15, 7).Value / totalmarkssum

        diff = (Sheet2.Cells(13, 7).Value - Sheet2.Cells(15, 7).Value) / 8
        diffp = diff / totalmarkssum
         

        If (Sheet3.Cells(i, 6).Value >= grade1) Then
            Sheet3.Cells(i, 8).Value = 1
        End If
        


        If (Sheet3.Cells(i, 6).Value > grade4 & Sheet3.Cells(i, 6).Value < grade1) Then


'        (Sheet3.Cells(i, 6).Value > grade4 + 0)

                                                If Sheet3.Cells(i, 6).Value > grade4 Then
                                            Sheet3.Cells(i, 8).Value = 3.7
                                        End If


                                        If Sheet3.Cells(i, 6).Value > (grade4 + diffp) Then
                                            Sheet3.Cells(i, 8).Value = 3.3
                                        End If

                                        If Sheet3.Cells(i, 6).Value > grade4 + 2 * diffp Then
                                            Sheet3.Cells(i, 8).Value = 3
                                        End If

                                        If Sheet3.Cells(i, 6).Value > grade4 + 3 * diffp Then
                                            Sheet3.Cells(i, 8).Value = 2.7
                                        End If
                                    '
                                        If Sheet3.Cells(i, 6).Value > grade4 + 4 * diffp Then
                                            Sheet3.Cells(i, 8).Value = 2.3
                                        End If

                                        If Sheet3.Cells(i, 6).Value > grade4 + 5 * diffp Then
                                            Sheet3.Cells(i, 8).Value = 2
                                        End If

                                        If Sheet3.Cells(i, 6).Value > grade4 + 6 * diffp Then
                                            Sheet3.Cells(i, 8).Value = 1.7
                                        End If

                                        If Sheet3.Cells(i, 6).Value > grade4 + 7 * diffp Then
                                            Sheet3.Cells(i, 8).Value = 1.3
                                        End If

                                        If Sheet3.Cells(i, 6).Value >= grade4 + 8 * diffp Then
                                            Sheet3.Cells(i, 8).Value = 1
                                        End If


   End If
   
    

    Next i

 'check once interpolation
 
 'interpolation finished
 'now count.....
    Dim c1 As Double
                Dim c2 As Double
                Dim c3 As Double
                Dim c4 As Double
                Dim c5 As Double
                Dim c6 As Double
                Dim c7 As Double
                Dim c8 As Double
                Dim c9 As Double
                Dim c10 As Double
                Dim c11 As Double
                Dim c12  As Double
                c12 = 0


                            c1 = 0
                            c2 = 0
                            c3 = 0
                            c4 = 0
                            c5 = 0
                            c6 = 0
                            c7 = 0
                            c8 = 0
                            c9 = 0
                            c10 = 0
                            c11 = 0
                For i = 2 To c
                        If Sheet3.Cells(i, 8).Value = 1 Then
                            c1 = 1 + c1
                        End If

                        If Sheet3.Cells(i, 8).Value = 1.3 Then
                            c2 = 1 + c2
                        End If

                        If Sheet3.Cells(i, 8).Value = 1.7 Then
                            c3 = 1 + c3
                        End If

                        If Sheet3.Cells(i, 8).Value = 2 Then
                            c4 = 1 + c4
                        End If

                        If Sheet3.Cells(i, 8).Value = 2.3 Then
                            c5 = 1 + c5
                        End If

                        If Sheet3.Cells(i, 8).Value = 2.7 Then
                            c6 = 1 + c6
                        End If

                        If Sheet3.Cells(i, 8).Value = 3 Then
                            c7 = 1 + c7
                        End If

                        If Sheet3.Cells(i, 8).Value = 3.3 Then
                            c8 = 1 + c8
                        End If

                        If Sheet3.Cells(i, 8).Value = 3.7 Then
                            c9 = 1 + c9
                        End If

                        If Sheet3.Cells(i, 8).Value = 4 Then
                            c10 = 1 + c10
                        End If

                        If Sheet3.Cells(i, 8).Value = 5 Then
                            c11 = 1 + c11
                        End If

                        If Sheet3.Cells(i, 8).Value = "N/A" Then
                            c12 = 1 + c12
                        End If




                Next i

                 Sheet4.Cells(4, 4).Value = c1
                        Sheet4.Cells(5, 4).Value = c2
                        Sheet4.Cells(6, 4).Value = c3
                        Sheet4.Cells(7, 4).Value = c4
                        Sheet4.Cells(8, 4).Value = c5
                        Sheet4.Cells(9, 4).Value = c6
                        Sheet4.Cells(10, 4).Value = c7
                        Sheet4.Cells(11, 4).Value = c8
                        Sheet4.Cells(12, 4).Value = c9
                        Sheet4.Cells(13, 4).Value = c10
                        Sheet4.Cells(14, 4).Value = c11
                        Sheet4.Cells(15, 4).Value = c12

       Dim ts As Double
       ts = Sheet4.Cells(16, 4)
       'relative
                        Sheet4.Cells(4, 5).Value = c1 / ts
                        Sheet4.Cells(5, 5).Value = c2 / ts
                        Sheet4.Cells(6, 5).Value = c3 / ts
                        Sheet4.Cells(7, 5).Value = c4 / ts
                        Sheet4.Cells(8, 5).Value = c5 / ts
                        Sheet4.Cells(9, 5).Value = c6 / ts
                        Sheet4.Cells(10, 5).Value = c7 / ts
                        Sheet4.Cells(11, 5).Value = c8 / ts
                        Sheet4.Cells(12, 5).Value = c9 / ts
                        Sheet4.Cells(13, 5).Value = c10 / ts
                        Sheet4.Cells(14, 5).Value = c11 / ts
                          Sheet4.Cells(15, 5).Value = c12 / ts
                        
 
 
 
 End If
 End If
 
 
 
 'For ref
 
 
 
 
 
 
'     If Sheet5.Cells(2, 1).Value = 0 Then
'
'
'
'
'    Else
'
'     For i = 3 To c
'     newsum1 = 0
'
'       For j = (6 + a + 1) To (6 + a + 1 + b)
'
'       newsum1 = newsum1 + ws.Cells(i, j)
'    Next j
'     ws.Cells(i, (6 + a)) = newsum
'
'    Next i
'
'   End If
'
   'single
'
    
    
'    ' Set the worksheet and range
'    Set conditionRange = ws.Range("Y2:AE2")
'
'    ' Check if the value in Sheet4.Cells(2, 1) is 1
'    If Sheet5.Cells(2, 1).Value = 0 Then
'
'                                                           csum = 0
'                                    ' Hide columns in the specified range
'                                    For Each col77 In conditionRange
'                                        ws.Columns(col77.Column).EntireColumn.Hidden = True
'                                    Next col77
'    End If
'
'
'
'
'
'
'
'   If Sheet5.Cells(2, 1).Value = 1 Then
'
''
'                            For Each col77 In conditionRange
'                                ws.Columns(col77.Column).EntireColumn.Hidden = False
'
'
'
'
'       Next col77
'End If
'
'
' If Sheet5.Cells(2, 1).Value = 0 Then
'Dim h As Double
'Dim cd As Double
'
'   cd = 0
'
''
'
'        'ws.Cells(3, 24).Value = csum1
''
''
'                  For j = 6 To 30
'
'
'                                 csum = csum + ws.Cells(3, j).Value
'
'                                        If j = 23 Then
'                                          j = j + 2
'                                          ws.Cells(3, 24).Value = csum
'                                          csum = 0
'                                        End If
'
'
'
'
'                                 Next j
'     ws.Cells(3, 31).Value = 0
'     ws.Cells(3, 32).Value = ws.Cells(3, 31).Value + ws.Cells(3, 24).Value
'
'
'
'    s = 0
'    ws.Cells(1, 7).Value = lastrow
'    c = 3 + ws.Cells(1, 3).Value
'    For i = 4 To c
'         s = 0
'
'        For j = 6 To 23
'            s = s + ws.Cells(i, j).Value
'        Next j
'        ws.Cells(i, 24).Value = s
'        s = 0
'    Next i
'
'   For i = 4 To c
'
'         s = 0
'
'        For j = 25 To 30
'            s = s + ws.Cells(i, j).Value
'        Next j
'        ws.Cells(i, 31).Value = s
'        s = 0
'    Next i
'
'                            For i = 4 To c
'
'                                ws.Cells(i, 32).Value = ws.Cells(i, 31).Value + ws.Cells(i, 24).Value
'
'                            Next i
'
'                             For i = 4 To c
'
'                                ws.Cells(i, 33).Value = ws.Cells(i, 32).Value / ws.Cells(3, 32).Value
'
'
'                            Next i
''
' End If
'
' If Sheet5.Cells(2, 1).Value = 1 Then
'
'
'   cd = 0
'
''
'
'        'ws.Cells(3, 24).Value = csum1
''
''
'                  For j = 6 To 30
'
'
'                                 csum = csum + ws.Cells(3, j).Value
'
'                                        If j = 23 Then
'                                          j = j + 2
'                                          ws.Cells(3, 24).Value = csum
'                                          csum = 0
'                                        End If
'
'
'
'
'                                 Next j
'     ws.Cells(3, 31).Value = csum
'     ws.Cells(3, 32).Value = ws.Cells(3, 31).Value + ws.Cells(3, 24).Value
'
'
'
'    s = 0
'    ws.Cells(1, 7).Value = lastrow
'    c = 3 + ws.Cells(1, 3).Value
'    For i = 4 To c
'         s = 0
'
'        For j = 6 To 23
'            s = s + ws.Cells(i, j).Value
'        Next j
'        ws.Cells(i, 24).Value = s
'        s = 0
'    Next i
'
'   For i = 4 To c
'
'         s = 0
'
'        For j = 25 To 30
'            s = s + ws.Cells(i, j).Value
'        Next j
'        ws.Cells(i, 31).Value = s
'        s = 0
'    Next i
'
'                            For i = 4 To c
'
'                                ws.Cells(i, 32).Value = ws.Cells(i, 31).Value + ws.Cells(i, 24).Value
'
'                            Next i
'
'                             For i = 4 To c
'
'                                ws.Cells(i, 33).Value = ws.Cells(i, 32).Value / ws.Cells(3, 32).Value
'
'
'                            Next i
''
' End If
' Sheet3.Cells(2, 2).Value = ws.Cells(2, 2).Value
'
'    For i = 2 To c
'
'       Sheet3.Cells(i, 2).Value = ws.Cells(i, 2).Value
'        Sheet3.Cells(i, 3).Value = ws.Cells(i, 3).Value
'        Sheet3.Cells(i, 4).Value = ws.Cells(i, 4).Value
'        Sheet3.Cells(i, 5).Value = ws.Cells(i, 5).Value
'        Sheet3.Cells(i, 6).Value = ws.Cells(i, 33).Value
'
'        If Sheet3.Cells(i, 6).Value >= Sheet2.Cells(13, 4).Value Then
'
'        Sheet3.Cells(i, 7).Value = "PASS"
'
'        Else
'
'        Sheet3.Cells(i, 7).Value = "not PASS"
'
'        End If
'
'
'      ' Sheet3.Cells(i, 15).Value = Sheet3.Cells(i, 6).Value
'      ' Sheet3.Cells(i, 16).Value = Sheet2.Cells(13, 4).Value
'       Sheet5.Cells(i, 15).Value = Sheet3.Cells(i, 6).Value
'       Sheet5.Cells(i, 16).Value = Sheet2.Cells(13, 4).Value
'
'
'        If Sheet5.Cells(i, 15).Value >= Sheet5.Cells(2, 16).Value Then
'
'        Sheet3.Cells(i, 7).Value = "PASS"
'
'        Else
'
'        Sheet3.Cells(i, 7).Value = "not PASS"
'
'        End If
'
'    If i = 2 Then
'        i = i + 1
'       End If
'
'    Next i
'
'
'     For i = 2 To c
'
'       If Sheet5.Cells(1, 1).Value = 1 Then
'    Select Case Sheet3.Cells(i, 6).Value
'        Case Is >= 0.895
'            Sheet3.Cells(i, 8).Value = 1
'        Case Is >= 0.83
'            Sheet3.Cells(i, 8).Value = 1.3
'        Case Is >= 0.755
'            Sheet3.Cells(i, 8).Value = 1.7
'        Case Is >= 0.685
'            Sheet3.Cells(i, 8).Value = 2
'        Case Is >= 0.62
'            Sheet3.Cells(i, 8).Value = 2.3
'        Case Is >= 0.545
'            Sheet3.Cells(i, 8).Value = 2.7
'        Case Is >= 0.475
'            Sheet3.Cells(i, 8).Value = 3
'        Case Is >= 0.41
'            Sheet3.Cells(i, 8).Value = 3.3
'        Case Is >= 0.335
'            Sheet3.Cells(i, 8).Value = 3.7
'        Case Is >= 0.3
'            Sheet3.Cells(i, 8).Value = 4
'        Case Is >= 0
'            Sheet3.Cells(i, 8).Value = 5
''             Sheet3.Cells(i, 8).Interior.Color = RGB(255, 204, 204)
'    End Select
'Else
'    Select Case Sheet3.Cells(i, 6).Value
'        Case Is >= 0.9
'            Sheet3.Cells(i, 8).Value = 1
'        Case Is >= 0.835
'            Sheet3.Cells(i, 8).Value = 1.3
'        Case Is >= 0.77
'            Sheet3.Cells(i, 8).Value = 1.7
'        Case Is >= 0.705
'            Sheet3.Cells(i, 8).Value = 2
'        Case Is >= 0.645
'            Sheet3.Cells(i, 8).Value = 2.3
'        Case Is >= 0.575
'            Sheet3.Cells(i, 8).Value = 2.7
'        Case Is >= 0.515
'            Sheet3.Cells(i, 8).Value = 3
'        Case Is >= 0.45
'            Sheet3.Cells(i, 8).Value = 3.3
'        Case Is >= 0.385
'            Sheet3.Cells(i, 8).Value = 3.7
'        Case Is >= 0.35
'            Sheet3.Cells(i, 8).Value = 4
'        Case Is >= 0
'            Sheet3.Cells(i, 8).Value = 5
''            Set targetcell = Sheet3.Cells(i, 8).Value
''         targetcell.Interior.Color = RGB(255, 204, 204)
'    End Select
'End If
'
'
'If i = 2 Then
'        Sheet3.Cells(i, 8).Value = "Grades"
'       End If
'
'       If i = 2 Then
'        i = i + 1
'       End If
'
'Next i
'
'
'
'For i = 4 To c
'Dim x As Double
'Dim x1 As Double
' If (Sheet3.Cells(i, 6).Value > Sheet2.Cells(15, 7).Value / ws.Cells(3, 32).Value) Then
'    Sheet3.Cells(i, 8).Value = 1
'
'
'
' ElseIf (Sheet3.Cells(i, 6).Value < Sheet2.Cells(17, 7).Value / ws.Cells(3, 32).Value) Then
'
'    Sheet3.Cells(i, 8).Value = 5
'
'
'  Else
'    Sheet5.Cells(i, 6).Value = Sheet3.Cells(i, 6).Value
'    Sheet5.Cells(15, 7).Value = Sheet2.Cells(15, 7).Value / 100
'   Sheet3.Cells(i, 8).Value = linear(Sheet3.Cells(i, 6).Value, Sheet2.Cells(15, 7).Value / ws.Cells(3, 32).Value, 1, Sheet2.Cells(17, 7).Value / ws.Cells(3, 32).Value, 4)
'
' End If
' Next i
'
'
'  For i = 4 To c
'        Dim grade1 As Double
'        Dim grade4 As Double
'        Dim diff As Double
'        Dim diffp As Double
'
'        grade1 = Sheet2.Cells(15, 7).Value / ws.Cells(3, 32).Value
'        grade4 = Sheet2.Cells(17, 7).Value / ws.Cells(3, 32).Value
'        diff = (Sheet2.Cells(15, 7).Value - Sheet2.Cells(17, 7).Value) / 8
'        diffp = diff / ws.Cells(3, 32).Value
'
'        If (Sheet3.Cells(i, 6).Value > grade1) Then
'            Sheet3.Cells(i, 8).Value = 1
'
'        ElseIf (Sheet3.Cells(i, 6).Value < grade4) Then
'            Sheet3.Cells(i, 8).Value = 5
'            Sheet3.Cells(i, 8).Interior.Color = RGB(255, 204, 204)
'
'            End If
'
'
'        If (Sheet3.Cells(i, 6).Value > grade4 & Sheet3.Cells(i, 6).Value < grade1) Then
'
'
''        (Sheet3.Cells(i, 6).Value > grade4 + 0)
'
'            If Sheet3.Cells(i, 6).Value > grade4 Then
'        Sheet3.Cells(i, 8).Value = 3.7
'    End If
'
'
'    If Sheet3.Cells(i, 6).Value > (grade4 + diffp) Then
'        Sheet3.Cells(i, 8).Value = 3.3
'    End If
'
'    If Sheet3.Cells(i, 6).Value > grade4 + 2 * diffp Then
'        Sheet3.Cells(i, 8).Value = 3
'    End If
'
'    If Sheet3.Cells(i, 6).Value > grade4 + 3 * diffp Then
'        Sheet3.Cells(i, 8).Value = 2.7
'    End If
'
'    If Sheet3.Cells(i, 6).Value > grade4 + 4 * diffp Then
'        Sheet3.Cells(i, 8).Value = 2.3
'    End If
'
'    If Sheet3.Cells(i, 6).Value > grade4 + 5 * diffp Then
'        Sheet3.Cells(i, 8).Value = 2
'    End If
'
'    If Sheet3.Cells(i, 6).Value > grade4 + 6 * diffp Then
'        Sheet3.Cells(i, 8).Value = 1.7
'    End If
'
'    If Sheet3.Cells(i, 6).Value > grade4 + 7 * diffp Then
'        Sheet3.Cells(i, 8).Value = 1.3
'    End If
'
'    If Sheet3.Cells(i, 6).Value > grade4 + 8 * diffp Then
'        Sheet3.Cells(i, 8).Value = 1
'    End If
'        End If
'    Next i
'    Sheet3.Cells(3, 26).Value = grade4 + 2 * diffp
'      Sheet3.Cells(3, 27).Value = grade4
'      Sheet3.Cells(3, 28).Value = diffp
'
'
'
'     Dim d12 As Double
'     Dim rng2 As Range
'     Dim colLimit12 As Double
'     Dim cell12 As Range
'     d12 = c
'      Dim endCol122 As Double
'
'     If Sheet5.Cells(2, 1).Value = 1 Then
'      colLimit12 = 27
'
'      Else
'      colLimit12 = 33
'
'     End If
'
''
''
''
'     'Define the range to loop through. Example: from row 2 to row d1
'    Set rng2 = ws.Range("A2:A" & d12)
'
'    For Each cell12 In rng2
'        ' Find the last used column in the row of the current cell
'        endCol122 = Application.WorksheetFunction.Min(colLimit12, ws.Cells(cell12.row, ws.Columns.Count).End(xlToLeft).Column)
'
'        If endCol22 >= 1 Then
'             'Apply borders to the range from the current cell to column colLimit1
'            With ws.Range(cell12, ws.Cells(cell12.row, colLimit12))
'                .Borders.LineStyle = xlContinuous
'                .Borders.ColorIndex = xlAutomatic ' Optional: Set the border color
'                .Borders.Weight = xlThin
'            End With
'        End If
'    Next cell12
'
'
'Set rng = ws.Range("A2:A" & (c))
'
'For Each cell In rng
'
'
'            Dim endCol As Integer
'            endCol = Application.WorksheetFunction.Min(colLimit, ws.Cells(cell.row, ws.Columns.Count).End(xlToLeft).Column)
'
'            If endCol >= 1 Then
'                With ws.Range(cell, ws.Cells(cell.row, 33))
'                    .Borders.LineStyle = xlContinuous
'                    .Borders.ColorIndex = xlAutomatic ' Optional: Set the border color
'                    .Borders.Weight = xlThin
'                End With
'            End If
'   Next cell
'
'    For i = 4 To c
'   If Sheet3.Cells(i, 5).Value = "Not Present" Then
'
'   Sheet3.Cells(i, 8) = "N/A"
'   Sheet3.Cells(i, 8).Interior.Color = RGB(135, 206, 235)
'
'
'   End If
'
'   Next i
'
'   Dim c1 As Double
'                Dim c2 As Double
'                Dim c3 As Double
'                Dim c4 As Double
'                Dim c5 As Double
'                Dim c6 As Double
'                Dim c7 As Double
'                Dim c8 As Double
'                Dim c9 As Double
'                Dim c10 As Double
'                Dim c11 As Double
'                Dim c12  As Double
'                c12 = 0
'
'
'                            c1 = 0
'                            c2 = 0
'                            c3 = 0
'                            c4 = 0
'                            c5 = 0
'                            c6 = 0
'                            c7 = 0
'                            c8 = 0
'                            c9 = 0
'                            c10 = 0
'                            c11 = 0
'                For i = 2 To c
'                        If Sheet3.Cells(i, 8).Value = 1 Then
'                            c1 = 1 + c1
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 1.3 Then
'                            c2 = 1 + c2
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 1.7 Then
'                            c3 = 1 + c3
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 2 Then
'                            c4 = 1 + c4
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 2.3 Then
'                            c5 = 1 + c5
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 2.7 Then
'                            c6 = 1 + c6
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 3 Then
'                            c7 = 1 + c7
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 3.3 Then
'                            c8 = 1 + c8
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 3.7 Then
'                            c9 = 1 + c9
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 4 Then
'                            c10 = 1 + c10
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = 5 Then
'                            c11 = 1 + c11
'                        End If
'
'                        If Sheet3.Cells(i, 8).Value = "N/A" Then
'                            c12 = 1 + c12
'                        End If
'
'
'
'
'                Next i
'
'                 Sheet4.Cells(4, 4).Value = c1
'                        Sheet4.Cells(5, 4).Value = c2
'                        Sheet4.Cells(6, 4).Value = c3
'                        Sheet4.Cells(7, 4).Value = c4
'                        Sheet4.Cells(8, 4).Value = c5
'                        Sheet4.Cells(9, 4).Value = c6
'                        Sheet4.Cells(10, 4).Value = c7
'                        Sheet4.Cells(11, 4).Value = c8
'                        Sheet4.Cells(12, 4).Value = c9
'                        Sheet4.Cells(13, 4).Value = c10
'                        Sheet4.Cells(14, 4).Value = c11
'                        Sheet4.Cells(15, 4).Value = c12
'
'       Dim ts As Double
'       ts = Sheet4.Cells(16, 4)
'       'relative
'                        Sheet4.Cells(4, 5).Value = c1 / ts
'                        Sheet4.Cells(5, 5).Value = c2 / ts
'                        Sheet4.Cells(6, 5).Value = c3 / ts
'                        Sheet4.Cells(7, 5).Value = c4 / ts
'                        Sheet4.Cells(8, 5).Value = c5 / ts
'                        Sheet4.Cells(9, 5).Value = c6 / ts
'                        Sheet4.Cells(10, 5).Value = c7 / ts
'                        Sheet4.Cells(11, 5).Value = c8 / ts
'                        Sheet4.Cells(12, 5).Value = c9 / ts
'                        Sheet4.Cells(13, 5).Value = c10 / ts
'                        Sheet4.Cells(14, 5).Value = c11 / ts
'
'                        Sheet4.Cells(15, 5).Value = c12 / ts
'
'
'
'   Dim ws1 As Worksheet
'    Set ws1 = ThisWorkbook.Worksheets("Grade")
'
'
'    Dim rng1 As Range
'    Dim cell1 As Range
'    Dim colLimit1 As Integer
'    Dim endCol1 As Integer
'    Dim d1 As Double
'
'
'    d1 = c
'
'    colLimit1 = 8 ' Column limit up to which borders will be applied
'
'    ' Define the range to loop through. Example: from row 2 to row d1
'    Set rng1 = ws1.Range("A2:A" & d1)
'
'    For Each cell1 In rng1
'        ' Find the last used column in the row of the current cell
'        endCol1 = Application.WorksheetFunction.Min(colLimit1, ws1.Cells(cell1.row, ws1.Columns.Count).End(xlToLeft).Column)
'
'        If endCol1 >= 1 Then
'            ' Apply borders to the range from the current cell to column colLimit1
'            With ws1.Range(cell1, ws1.Cells(cell1.row, colLimit1))
'                .Borders.LineStyle = xlContinuous
'                .Borders.ColorIndex = xlAutomatic ' Optional: Set the border color
'                .Borders.Weight = xlThin
'            End With
'        End If
'    Next cell1
'     Dim a As Double
'     a = 0
'     Dim b As Double
'     b = 0
'
'
'    For i = 2 To (c - 2)
'    a = a + 1
'    ws.Cells((i + 2), 1).Value = a
'    b = b + 1
'    Sheet3.Cells((i + 2), 1) = b
'
'    Next i
End If
End If

'
End Sub

Sub clear()
Dim ws As Worksheet
Dim i As Double
Dim result As VbMsgBoxResult

result = MsgBox("Are you sure", vbYesNo)



Set ws = ThisWorkbook.Worksheets("Input")
Dim c As Double
 c = 4 + 0
    
   If result = vbYes Then
   
   
Sheet3.Cells.clear
'Sheet5.Cells.clear
With ws
            .Range("A" & c, .Cells(.Rows.count, 3)).clear
            .Range("F" & c, .Cells(.Rows.count, .Columns.count)).clear
    End With
     For i = 4 To (100 + ws.Cells(1, 3).Value)
      Sheet1.Cells(i, 5) = ""
     Next i
     Sheet2.Cells(13, 7) = ""
     Sheet2.Cells(15, 7) = ""
     
     For i = 9 To 23
      
      Sheet2.Cells(i, 4).Value = ""
      
      i = i + 1
     
     Next i
     
     Dim j As Long
     Dim columnsToDelete As New Collection
    
   For j = 1 To 1000
        If IsNumeric(ws.Cells(2, j).Value) Then
            If ws.Cells(2, j).Value > 1 Then
                columnsToDelete.Add j
            End If
        End If
    Next j

    ' Delete columns from the end to the start
    For j = columnsToDelete.count To 1 Step -1
        ws.Columns(columnsToDelete(j)).Delete
    Next j
     ws.Cells.EntireRow.Hidden = False
ws.Cells.EntireColumn.Hidden = False

Dim clea As Double
For clea = 6 To 2000
ws.Cells(3, clea) = ""
Next clea

End If

Dim clea2 As Double

For clea2 = 4 To 14
Sheet4.Cells(clea2, 4).Value = 0
Sheet4.Cells(clea2, 5).Value = 0
Next clea2



    
End Sub

Sub ExportSheetToPDF()
    Dim ws4 As Worksheet
    Dim pdfFilePath As String
    
    ' Define the worksheet
    On Error Resume Next
    Set ws4 = ThisWorkbook.Sheets("Statistics")
    On Error GoTo 0
    
    
    pdfFilePath = "C:\Users\User\Documents\Statisticsstical Data.pdf"
    
    ' Export the sheet to PDF
    ws4.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard, _
                            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    MsgBox "PDF saved successfully at: " & pdfFilePath, vbInformation
    
    Dim ws5 As Worksheet
    
    
    ' Define the worksheet
    On Error Resume Next
    Set ws5 = ThisWorkbook.Sheets("Grade")
    On Error GoTo 0
    
    
    pdfFilePath = "C:\Users\User\Documents\Grades.pdf"
    
    ' Export the sheet to PDF
    ws5.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard, _
                            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    MsgBox "PDF2  saved successfully at: " & pdfFilePath, vbInformation
    
    
End Sub

Sub create_PDF()
'
' create_PDF Makro
'

'
    Sheets(Array("Statistics", "Grade")).Select
    Sheets("Statistics").Activate
    Application.Dialogs(xlDialogPrint).Show

End Sub


Sub ref()


       Dim s As Double
    Dim lastrow As Long  ' Use Long instead of Double for row counts
    Dim i As Double, j As Double
    Dim c As Double
    Dim rng As Range
    Dim cell As Range
    Dim colLimit As Integer
    Dim targetcell As Range
    Dim target As Range
    Dim csum As Double
    Dim i1 As Long
    Dim csum1 As Double
    Dim colcount As Double
    Dim countssss As Double
    
    colLimit = 33
    'fff
    
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Input")
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("References")
    
     c = 3 + ws.Cells(1, 3).Value
    
    lastrow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
     Dim col77 As Range
    Dim conditionRange As Range
    Dim checkCell As Range
    
 'Define variables
     Dim a As Double
    Dim b As Double
    
    a = Sheet2.Cells(15, 4)
    b = Sheet2.Cells(17, 4)
    Dim newsum As Double
    newsum = 0
    
 Sheet5.Cells(27, 1).Value = 1
 Dim totalmarkssum As Double
 Dim newsum1 As Double
 
   If Sheet5.Cells(2, 1).Value = 0 Then
   
    For i = 3 To c
     newsum = 0
     
       For j = 6 To (6 + a - 1)

       newsum = newsum + ws.Cells(i, j)
    Next j
     ws.Cells(i, (6 + a)) = newsum
     ws.Cells(i, (6 + a + 3)) = newsum
     totalmarkssum = ws.Cells(3, (6 + a))
      colcount = 6 + a
     
     If totalmarkssum > 0 Then
     ws.Cells(i, (6 + a + 4)) = newsum / totalmarkssum
     Else
     End If
     
     
    Next i
    
    
    
    Else
          
          For i = 3 To c
     newsum = 0
     
       For j = 6 To (6 + a - 1)

       newsum = newsum + ws.Cells(i, j)
    Next j
     ws.Cells(i, (6 + a)) = newsum
     newsum1 = 0
       For j = (6 + a + 1) To (6 + a + 1 + b - 1)

       newsum1 = newsum1 + ws.Cells(i, j)
    Next j
     ws.Cells(i, (6 + a + 1 + b - 1 + 1)) = newsum1
     ws.Cells(i, (6 + a + 1 + b - 1 + 1 + 1)) = newsum1 + newsum
     totalmarkssum = ws.Cells(3, (6 + a + 1 + b - 1 + 1 + 1))
     colcount = 6 + a + 1 + b - 1 + 1 + 1
     
     If totalmarkssum > 0 Then
     ws.Cells(i, (6 + a + 1 + b - 1 + 1 + 1 + 1)) = ws.Cells(i, (6 + a + 1 + b - 1 + 1 + 1)) / totalmarkssum
     Else
     
     End If
     
     
     
     
    Next i
    
End If
 If totalmarkssum > 0 Then
 If Sheet6.Cells(2, 2).Value >= 0 Then
 If Sheet6.Cells(3, 2).Value > 0 Then
 
 
 
 
 Dim forref As Range
    Set forref = ws.Range(ws.Cells(4, colcount), ws.Cells(ws.Rows.count, colcount))

    Dim criteria As String
    Dim startVal As Double, ed As Double, exi As Double, countVal As Double
    
    
    startVal = ThisWorkbook.Sheets("References").Cells(2, 2).Value
    exi = ThisWorkbook.Sheets("References").Cells(3, 2).Value
    ed = totalmarkssum
    Dim k As Double
    k = 6
    
    For i = startVal To ed Step exi
        criteria = ">" & i
        countVal = Application.WorksheetFunction.CountIf(forref, criteria)
        ThisWorkbook.Sheets("References").Cells((k + 1), 1).Value = "Greater than " & i
        ThisWorkbook.Sheets("References").Cells((k + 1), 2).Value = countVal
        

        If k = 6 Then
            ThisWorkbook.Sheets("References").Cells(k, 1).Value = "Greater-than"
            ThisWorkbook.Sheets("References").Cells(k, 2).Value = "count"
            ThisWorkbook.Sheets("References").Cells(k, 3).Value = "Cummulative%"
        End If

        k = k + 1
    Next i
 
  If Sheet2.Cells(9, 4).Value > 0 Then
'   Dim ws1 As Worksheet
'   Set ws1 = ThisWorkbook.Worksheets("References")
   Dim dummy As Double
   dummy = 7
   
     For i = startVal To ed Step exi
     
     ws1.Cells(dummy, 3) = ws1.Cells(dummy, 2) / Sheet2.Cells(9, 4).Value * 100
     dummy = dummy + 1
     countssss = countssss + 1
     
     
     Next i
     Sheet6.Cells(1, 2).Value = totalmarkssum
     
  
   End If
   
 End If
 End If
 End If
 
 
 
    Dim d12 As Double
Dim rng2 As Range
Dim colLimit12 As Double
Dim cell12 As Range
Dim endCol122 As Double

    

    ' Set the worksheet, replace "SheetName" with your worksheet name
    
    ' Initialize d12 with a value, replace 'c' with your specific value or calculation
    d12 = 6

' Initialize colLimit12 with a value
colLimit12 = 3

' Set the range from A6 to the last row defined by d12
Set rng2 = ws1.Range("A6:A" & (d12 + countssss))

' Loop through each cell in the range
For Each cell12 In rng2
    ' Calculate the end column based on the minimum of colLimit12 and the last non-empty column
    endCol122 = Application.WorksheetFunction.Min(colLimit12, ws1.Cells(cell12.row, ws1.Columns.count).End(xlToLeft).Column)
    
    ' Check if the end column is greater than or equal to 1
    If endCol122 >= 1 Then
        ' Apply borders to the range from the current cell to the determined end column
        With ws1.Range(cell12, ws1.Cells(cell12.row, endCol122))
            .Borders.LineStyle = xlContinuous
            .Borders.ColorIndex = xlAutomatic ' Optional: Set the border color
            .Borders.Weight = xlThin
        End With
    End If
Next cell12
 
 
End Sub

Sub CLEARREF()
  Dim ws1 As Worksheet
   Set ws1 = ThisWorkbook.Worksheets("References")
   Dim nthterm As Double
   Dim ts As Double
   
   If Sheet6.Cells(3, 2).Value > 0 Then
   
   nthterm = (ts - Sheet6.Cells(2, 2).Value) / Sheet6.Cells(3, 2).Value + 1
   Sheet6.Cells(1, 3).Value = nthterm
   
   End If
   nthterm = 6
   With ws1
            .Range("A" & nthterm, .Cells(.Rows.count, 3)).clear
    End With
   
End Sub
