Sub stockmkt_moderate_Cons()
' Declare the variable where the Stock Volume will be stored
    Dim volume_tot as double
    Dim open_yrl as double
    Dim close_yrl as double
' Print the headers on Cells I1 & J1
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    ' Find the last row - using Cells Rows.Count and the column "A"_Data is located on this column
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    ' Loop until the last row is found starting on row #2 (to avoid the header)
        For i = 2 to lastrow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                open_yrl = open_yrl + Cells(i,3).Value
                close_yrl = close_yrl + Cells(i,6).Value
                volume_tot = volume_tot + Cells(i,7).Value
                Range("I" & 2 + j).Value = Cells(i,1).Value
                Range("J" & 2 + j).Value = open_yrl - close_yrl
                If open_yrl >= close_yrl then
                    Range("J" & 2 + j).Interior.ColorIndex = 4
                Else
                    Range("J" & 2 + j).Interior.ColorIndex = 3
                End if
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("K" & 2 + j).Value = 1 - (close_yrl / open_yrl)
                Range("L" & 2 + j).Value = volume_tot
                volume_tot = 0
                open_yrl = 0
                close_yrl = 0
                j = j + 1
            Else
                volume_tot = volume_tot + Cells(i,7).Value
                open_yrl = open_yrl + Cells(i,3).Value
                close_yrl = close_yrl + Cells(i,6).Value
            End if
        next i
end sub
