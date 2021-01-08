Sub Step4()

Dim RqID As Double
Dim ws As Worksheet
Dim Header(1 To 10) As String
Dim SheetName As String
Dim Title As String

'---------------------------

RqID = 122133

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

SheetName = RqID & "_Surge_IFSM"

Title = "Summary of Last Good Pass of Each Piece"

'---------------------------


'//Step 1////////////////////////////
Dim hCount As Integer
hCount = 1

Set ws = Worksheets(SheetName)
ws.Select
ws.Activate

Cells(9, 12).Value = Title
Cells(9, 12).Font.Bold = True


'//Step 2////////////////////////////

    For j = 12 To 21
        Cells(10, j).Value = Header(hCount)
        hCount = hCount + 1
    Next j


Range("L10:U10").Borders(xlEdgeBottom).Weight = xlThick
Range("L10:U10").Font.Bold = True
Range("L10:U10").Columns.AutoFit



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


    RowResultValue = 11
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


                    StRange = "B" & (RowInputValue - 1) & ":J" & (RowInputValue - 1)
                    ws.Columns(12).Rows(RowResultValue) = Status(m)
                    Sheets(m).Range(StRange).Copy ws.Columns(13).Rows(RowResultValue)

                    RowResultValue = RowResultValue + 1
                    RowInputValue = RowInputValue + 1

                    Else

                    MsgBox " No Fail "

                    End If

                    
                    RowInputValue = 12
        End If




    Next m



End Sub




