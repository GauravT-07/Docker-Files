# Docker-Files
This Repo contains Docker files
Sub MergeSameRequirements()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim startRow As Long
    Dim i As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    startRow = 2   'Assumes row 1 is the header

    For i = 3 To lastRow + 1

        If ws.Cells(i, 1).Value <> ws.Cells(startRow, 1).Value Then

            If i - startRow > 1 Then
                With ws.Range(ws.Cells(startRow, 1), ws.Cells(i - 1, 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            End If

            startRow = i

        End If

    Next i

End Sub

