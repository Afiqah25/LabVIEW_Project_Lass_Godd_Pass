Sub Step5()

Dim Technician As String
Dim DeviceID As String
Dim LotID As String
Dim Temp As Double
Dim ws As Worksheet

Dim Title As String
Dim SheetName As String

Dim Header(1 To 4) As String

'---------------------------

Technician = "Fadzly"
DeviceID = "ABC123DIODE"
LotID = "ABC123LOT"
Temp = 25

Header(1) = "Device No"
Header(2) = "  Pass(A)  "
Header(3) = "  Fail(A)  "
Header(4) = "Pass Ifsm_MV(V)"

SheetName = "Summary Table"
Title = "SUMMARY TABLE"

'---------------------------

'//Step 1////////////////////////////
Dim hCount As Integer
hCount = 1

Set ws = Sheets.Add
ws.Name = SheetName
ws.Select
ws.Activate

ws.Cells(1, 2).Value = Title
ws.Cells(1, 2).Font.Bold = True
ws.Cells(1, 2).Font.Size = 14

'//Step 1.5////////////////////////////

ws.Cells(2, 2).Value = "Device"
ws.Cells(2, 2).Font.Bold = True
ws.Cells(2, 3).Value = DeviceID
ws.Cells(2, 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
ws.Cells(2, 4).Borders(xlEdgeBottom).LineStyle = xlContinuous

ws.Cells(3, 2).Value = "Technician"
ws.Cells(3, 2).Font.Bold = True
ws.Cells(3, 3).Value = Technician
ws.Cells(3, 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
ws.Cells(3, 4).Borders(xlEdgeBottom).LineStyle = xlContinuous

ws.Cells(4, 2).Value = "LOT #"
ws.Cells(4, 2).Font.Bold = True
ws.Cells(4, 3).Value = LotID
ws.Cells(4, 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
ws.Cells(4, 4).Borders(xlEdgeBottom).LineStyle = xlContinuous

ws.Cells(2, 6).Value = "Sample Type"
ws.Cells(2, 6).Font.Bold = True
ws.Cells(2, 7).Value = "Diode_N"
ws.Cells(2, 7).Borders(xlEdgeBottom).LineStyle = xlContinuous


ws.Cells(3, 6).Value = "Wave"
ws.Cells(3, 6).Font.Bold = True
ws.Cells(3, 7).Value = "Sine"
ws.Cells(3, 7).Borders(xlEdgeBottom).LineStyle = xlContinuous


ws.Cells(2, 9).Value = "Temperature"
ws.Cells(2, 9).Font.Bold = True
ws.Cells(2, 10).Value = Temp & "C"
ws.Cells(2, 10).Borders(xlEdgeBottom).LineStyle = xlContinuous


ws.Cells(3, 9).Value = "Width"
ws.Cells(3, 9).Font.Bold = True
ws.Cells(3, 10).Value = "8.3ms"
ws.Cells(3, 10).Borders(xlEdgeBottom).LineStyle = xlContinuous

ws.Range("B2:J4").Columns.AutoFit

'//Step 2////////////////////////////

    For j = 2 To 5
        Cells(6, j).Value = Header(hCount)
        hCount = hCount + 1
    Next j


Range("B6:E6").Borders(xlEdgeBottom).Weight = xlThick
Range("B6:E6").Borders.LineStyle = xlContinuous
Range("B6:E6").Font.Bold = True
Range("B6:E6").Columns.AutoFit
Range("B6:E6").Interior.ColorIndex = 6 'Yellow


'//Step3 to 7////////////////////////////
'
Dim CountFront As Integer
Dim SheetN As Integer
Dim msg As String

SheetN = (Sheets.Count)
CountFront = 0
ReDim Status(1 To SheetN) As String


    '// Count how many available hastag
    For j = 1 To SheetN

             If Right(Sheets(j).Name, 1) = "#" Then

                                            If Len(Sheets(j).Name) = 2 Then
                                            'MsgBox Left(Sheets(j).Name, 1)
                                            Status(j) = Left(Sheets(j).Name, 1)

                                            ElseIf Len(Sheets(j).Name) = 3 Then
                                            'MsgBox Left(Sheets(j).Name, 2)
                                            Status(j) = Left(Sheets(j).Name, 2)

                                            ElseIf Len(Sheets(j).Name) = 4 Then
                                            'MsgBox Left(Sheets(j).Name, 3)
                                            Status(j) = Left(Sheets(j).Name, 3)

                                            End If
'            CountFront = CountFront + 1
            Else
                Status(j) = 0
            End If
    Next j



    '// Insert number(string) into array
    Dim RowResultValue As Integer
    Dim RowInputValue As Integer
    Dim DUT As String
    Dim StRange As String
    Dim StLine As String

    Dim StPass As String
    Dim StFail As String
    Dim StIfsmPass As String
    Dim St4Border As String

    RowResultValue = 7
    RowInputValue = 12
    DUT = " "



    For m = 1 To SheetN

        If Status(m) = 0 Then

        Else
                    If Sheets(m).Columns(7).Rows(RowInputValue).Value <> "FAIL" Then

                                Do While Sheets(m).Columns(7).Rows(RowInputValue).Value <> "FAIL"
                                RowInputValue = RowInputValue + 1
                                Loop
                    End If



                    If Sheets(m).Columns(7).Rows(RowInputValue).Value = "FAIL" Then

                    'MsgBox "Sheet Name :" & m & vbNewLine & Sheets(m).Columns(7).Rows(RowInputValue).Value & vbNewLine & "RowResultValue :" & RowResultValue & vbNewLine & "RowInputValue :" & RowInputValue


                    StPass = "C" & (RowInputValue - 1)
                    StFail = "C" & (RowInputValue)
                    StIfsmPass = "F" & (RowInputValue - 1)
                    St4Border = "B" & (RowResultValue) & ":E" & (RowResultValue)

                    ws.Columns(2).Rows(RowResultValue) = Status(m) 'DUT
                    ws.Columns(3).Rows(RowResultValue) = Sheets(m).Range(StPass) 'PASS
                    ws.Columns(4).Rows(RowResultValue) = Sheets(m).Range(StFail) 'FAIL
                    ws.Columns(5).Rows(RowResultValue) = Sheets(m).Range(StIfsmPass) 'Pass Ifsm_MV(V)
                    ws.Range(St4Border).Borders.LineStyle = xlContinuous

                    RowResultValue = RowResultValue + 1
                    RowInputValue = RowInputValue + 1

                    Else

                    MsgBox " No Fail "

                    End If


                    RowInputValue = 12
        End If




    Next m



End Sub

