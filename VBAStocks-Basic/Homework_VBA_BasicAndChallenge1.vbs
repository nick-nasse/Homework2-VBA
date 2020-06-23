Sub RealCoder()
    Dim Last_Row as Variant
    Dim I as Variant
    
    Dim Vticker as String
    Dim Volume as Variant 
    Dim Opening as Double
    Dim Closing as Double

    Dim Yearly_Change as Double
    Dim Percent_Change as Variant
    
    Dim J as Integer

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly-Change"
    Cells(1, 11).Value = "Percent-Change"
    Cells(1, 12).Value = "Total-Stock-Volume"
    
    Opening = Range("C2").Value
    Vticker = Range("A2").Value
    J = 2


    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row

    For I = 2 to Last_Row 

        Volume = Volume + Cells(I, 7).Value

        If Cells(I, 1).Value <> Cells(I, 1).Offset(1, 0).Value Then

            Closing = Cells(I, 6).Value
            
            Yearly_Change = Closing - Opening

            If Yearly_Change = 0 Or Opening = 0 Then 

                Percent_Change = 0 

            Else 

                Percent_Change = Yearly_Change / Opening

            End If

            Cells(J, 9).Value = Vticker
            Cells(J, 10).Value = Yearly_Change
            Cells(J, 11).Value = Percent_Change
            Cells(J, 12).Value = Volume

            Vticker = Cells(I, 1).Offset(1, 0).Value 
            Opening = Cells(I, 3).Offset(1, 0).Value
            Volume = 0
            J = J + 1

        End If

    Next I

    Dim Coloring_Row as Integer
    Dim P as Integer

    Coloring_Row = Cells(Rows.Count, 10).End(xlUp).Row

    For P = 2 to Coloring_Row

        If Cells(P, 10).value > 0 then 

            Cells(P, 10).Interior.Color = vbGreen

        ElseIf Cells(P, 10).value < 0 then

            Cells(P, 10).Interior.Color = vbRed

        End If 

    Next P

    Range(Cells(2, 11), Cells(Coloring_Row, 11)).NumberFormat = "0.00%"

    Dim L
    Dim M
    Dim N

    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    Dim GI_N as Variant 
    Dim GI_T as String

    Dim GD_N as Variant 
    Dim GD_T as String

    Dim GV_N as Variant 
    Dim GV_T as String

    GI_N = 0
    GD_N = 0
    GV_N = 0 

    For L = 2 to Coloring_Row

        If Cells(L, 11).Value > GI_N Then 

            GI_N = Cells(L, 11).Value
            GI_T = Cells(L, 11).Offset(0, -2).Value
    
        ElseIf Cells(L, 11).Value < GD_N Then

            GD_N = Cells(L, 11).Value
            GD_T = Cells(L, 11).Offset(0, -2).Value
    
        End If

    Next L 

    Range("P2").Value = GI_T
    Range("P3").Value = GD_T

    Range("Q2").Value = GI_N
    Range("Q3").Value = GD_N

    For M = 2 to Coloring_Row

        If Cells(M, 12).Value > GV_N Then 

            GV_N = Cells(M, 12).Value
            GV_T = Cells(M, 12).Offset(0, -3).Value

        End If 

    Next M

    Range("P4").Value = GV_T
    Range("Q4").Value = GV_N

    Range("Q2:Q3").NumberFormat = "0.00%"

End Sub