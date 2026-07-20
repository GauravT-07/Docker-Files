Dear [Manager's Name],

I hope you are doing well.

Please accept this email as my formal resignation from my position at Collins Aerospace. After careful consideration, I have decided to accept an opportunity with another organization. As per my notice period, my last working day will be in accordance with the company policy.

This was not an easy decision. One of the primary reasons for this move is that my new role is based in Gurgaon, which is much closer to my hometown. Being nearer to my family will allow me to better support and take care of my parents and fulfill my family responsibilities while continuing my professional growth.

I would like to sincerely thank you and the entire team for the opportunities, guidance, and support I have received during my time at Collins Aerospace. Working here has been a rewarding experience, and I am grateful for the trust placed in me to lead projects and work alongside such talented colleagues. The knowledge and experience I have gained will remain invaluable throughout my career.

Over the notice period, I will do my best to ensure a smooth transition by completing my pending responsibilities, documenting my work, and assisting with the knowledge transfer process.

Thank you once again for your mentorship and support. I wish you and the entire team continued success, and I hope our paths cross again in the future.

Sincerely,

Gaurav Totla
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

