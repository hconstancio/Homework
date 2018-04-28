Sub stockmkt_easy()
' Declare the variable where the Stock Volume will be stored
 Dim volume_tot as double
 ' Print the headers on Cells I1 & J1
 Range("I1").Value = "Ticker"
 Range("J1").Value = "Total Stock Volume"
 ' Find the last row - using Cells Rows.Count and the column "A"_Data is located on this column
 lastrow = Cells(Rows.Count, "A").End(xlUp).Row
 ' Loop until the last row is found starting on row #2 (to avoid the header)
    For i = 2 to lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            volume_tot = volume_tot + Cells(i,7).Value
            Range("I" & 2 + j).Value = Cells(i,1).Value
            Range("J" & 2 + j).Value = volume_tot
            volume_tot = 0
            j = j + 1
        Else
            volume_tot = volume_tot + Cells(i,7).Value
        End if
    next i
end sub