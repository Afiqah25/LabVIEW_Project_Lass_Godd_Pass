Sub Step3()

Dim Technician As String
Dim DeviceID As String
Dim LotID As String
Dim Direction As String

Dim RqID As Double
Dim Temp As Double
Dim Freq As Double

Dim ws As Worksheet
Dim CurrentDT As Date

Dim Header(1 To 10) As String

'---------------------------

RqID = 122133
Temp = 25
Freq = 60

Technician = "Fadzly"
DeviceID = "ABC123DIODE"
LotID = "ABC123LOT"
Direction = "Forward PNP"

Header(1) = "DUT"
Header(2) = "I_ifsm(A)"
Header(3) = "VF(V)(@If=0.010A)"
Header(4) = "Ifsm_MI(A)"
Header(5) = "Ifsm_MV(V)"
Header(6) = "Ir(mA)(@Vr=15V)"
Header(7) = "Result"
Header(8) = "Vf_chk(V)"
Header(9) = "PeakW(W)"
Header(10) = "Energy (J)"


'---------------------------

'//Step 1////////////////////////////

Set ws = Sheets.Add
ws.Name = RqID & "_Surge_IFSM"
ws.Select
ws.Activate

CurrentDT = Now

'//Step 1.5////////////////////////////

'ws.Range("H2:O7").Merge
'ws.Range("H2:O7").Value = "NOTE: "

'ws.Columns(1).Rows(1) = "Date"
ws.Range("A1:D1").Merge
ws.Range("A1:D1").Value = "Date"
ws.Range("A1:D1").HorizontalAlignment = xlRight
ws.Range("A1:D1").Font.Bold = True

'ws.Columns(1).Rows(2) = "Time"
ws.Range("A2:D2").Merge
ws.Range("A2:D2").Value = "Time"
ws.Range("A2:D2").HorizontalAlignment = xlRight
ws.Range("A2:D2").Font.Bold = True

'ws.Columns(1).Rows(3) = "Technician"
ws.Range("A3:D3").Merge
ws.Range("A3:D3").Value = "Technician"
ws.Range("A3:D3").HorizontalAlignment = xlRight
ws.Range("A3:D3").Font.Bold = True

'ws.Columns(1).Rows(4) = "Device #"
ws.Range("A4:D4").Merge
ws.Range("A4:D4").Value = "Device #"
ws.Range("A4:D4").HorizontalAlignment = xlRight
ws.Range("A4:D4").Font.Bold = True

'ws.Columns(1).Rows(5) = "Characterization #"
ws.Range("A5:D5").Merge
ws.Range("A5:D5").Value = "Characterization #"
ws.Range("A5:D5").HorizontalAlignment = xlRight
ws.Range("A5:D5").Font.Bold = True

'ws.Columns(1).Rows(6) = "Temperature"
ws.Range("A6:D6").Merge
ws.Range("A6:D6").Value = "Temperature"
ws.Range("A6:D6").HorizontalAlignment = xlRight
ws.Range("A6:D6").Font.Bold = True

'ws.Columns(1).Rows(7) = "Surge Type"
ws.Range("A7:D7").Merge
ws.Range("A7:D7").Value = "Surge Type"
ws.Range("A7:D7").HorizontalAlignment = xlRight
ws.Range("A7:D7").Font.Bold = True

'ws.Columns(1).Rows(8) = "Surge Direction"
ws.Range("A8:D8").Merge
ws.Range("A8:D8").Value = "Surge Direction"
ws.Range("A8:D8").HorizontalAlignment = xlRight
ws.Range("A8:D8").Font.Bold = True

'ws.Columns(2).Rows(1) = Format(Now, "mm/dd/yyyy")
ws.Range("E1:G1").Merge
ws.Range("E1:G1").Value = Format(Now, "mm/dd/yyyy")
ws.Range("E1:G1").HorizontalAlignment = xlLeft

'ws.Columns(2).Rows(2) = Format(Now, "HH:mm")
ws.Range("E2:G2").Merge
ws.Range("E2:G2").Value = Format(Now, "HH:mm")
ws.Range("E2:G2").HorizontalAlignment = xlLeft

'ws.Columns(2).Rows(3) = Technician
ws.Range("E3:G3").Merge
ws.Range("E3:G3").Value = Technician
ws.Range("E3:G3").HorizontalAlignment = xlLeft

'ws.Columns(2).Rows(4) = DeviceID
ws.Range("E4:G4").Merge
ws.Range("E4:G4").Value = DeviceID
ws.Range("E4:G4").HorizontalAlignment = xlLeft

'ws.Columns(2).Rows(5) = RqID
ws.Range("E5:G5").Merge
ws.Range("E5:G5").Value = RqID
ws.Range("E5:G5").HorizontalAlignment = xlLeft

'ws.Columns(2).Rows(6) = "+" & Temp & "C"
ws.Range("E6:G6").Merge
ws.Range("E6:G6").Value = "+" & Temp & "C"
ws.Range("E6:G6").HorizontalAlignment = xlLeft

'ws.Columns(2).Rows(7) = Freq & "Hz to Destruction (IFSM) "
ws.Range("E7:G7").Merge
ws.Range("E7:G7").Value = Freq & "Hz to Destruction (IFSM) "
ws.Range("E7:G7").HorizontalAlignment = xlLeft

'ws.Columns(2).Rows(8) = Direction
ws.Range("E8:G8").Merge
ws.Range("E8:G8").Value = Direction
ws.Range("E8:G8").HorizontalAlignment = xlLeft

'ws.Range("A1:G8").Columns.AutoFit
'ws.Range("A1:G8").HorizontalAlignment = xlRight

ws.Range("K2:R7").Select
ws.Range("K2:R7").Merge
ws.Range("K2:R7").Value = "NOTE: "
ws.Range("K2:R7").Borders.LineStyle = xlContinuous
ws.Range("K2:R7").Borders.Weight = xlThick

ws.Range("K2:R7").Font.Bold = True
ws.Range("K2:R7").HorizontalAlignment = xlLeft
ws.Range("K2:R7").VerticalAlignment = xlVAlignTop

'//Step2////////////////////////////

    For j = 1 To 10
        Cells(9, j).Value = Header(j)
    Next j


Range("A9:J9").Borders(xlEdgeBottom).Weight = xlThick
Range("A9:J9").Font.Bold = True
Range("A9:J9").Columns.AutoFit

''//Step3 to 7////////////////////////////
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


    RowResultValue = 10
    RowInputValue = 12
    DUT = " "



    For m = 1 To SheetN

        If Status(m) = 0 Then

        Else
                    If Sheets(m).Columns(7).Rows(RowInputValue).Value <> "FAIL" Then

                                Do While Sheets(m).Columns(7).Rows(RowInputValue).Value <> "FAIL"

                                StRange = "B" & RowInputValue & ":J" & RowInputValue

                                ws.Columns(1).Rows(RowResultValue) = Status(m)
                                Sheets(m).Range(StRange).Copy ws.Columns(2).Rows(RowResultValue)


                                RowResultValue = RowResultValue + 1
                                RowInputValue = RowInputValue + 1
                                Loop
                    End If



                    If Sheets(m).Columns(7).Rows(RowInputValue).Value = "FAIL" Then

                    'MsgBox "Sheet Name :" & m & vbNewLine & Sheets(m).Columns(7).Rows(RowInputValue).Value & vbNewLine & "RowResultValue :" & RowResultValue & vbNewLine & "RowInputValue :" & RowInputValue


                    StRange = "B" & RowInputValue & ":J" & RowInputValue
                    ws.Columns(1).Rows(RowResultValue) = Status(m)
                    Sheets(m).Range(StRange).Copy ws.Columns(2).Rows(RowResultValue)

                    StLine = "A" & RowResultValue & ":J" & RowResultValue
                    Range(StLine).Borders(xlEdgeBottom).Weight = xlThick

                    RowResultValue = RowResultValue + 1
                    RowInputValue = RowInputValue + 1

                    Else

                    MsgBox " No Fail "

                    End If

                    
                    RowInputValue = 12
        End If




    Next m




End Sub




