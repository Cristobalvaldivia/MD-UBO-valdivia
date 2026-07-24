Sub Geocodificar()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dir As String
    Dim http As Object
    Dim url As String
    Dim resp As String
    Dim lat As String
    Dim lon As String
    Dim p1 As Long
    Dim p2 As Long

    On Error GoTo fin

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        dir = ws.Cells(i, 1).Value

        If dir <> "" Then
            ws.Cells(i, 4).Value = "Buscando..."
            DoEvents

            Set http = CreateObject("MSXML2.XMLHTTP.6.0")
            url = "https://nominatim.openstreetmap.org/search?q=" & dir & "&format=json&limit=1"

            http.Open "GET", url, False
            http.Send
            resp = http.responseText

            If InStr(resp, """lat""") > 0 Then
                p1 = InStr(resp, """lat""") + 7
                p2 = InStr(p1, resp, ",")
                lat = Mid(resp, p1, p2 - p1 - 1)

                p1 = InStr(resp, """lon""") + 7
                p2 = InStr(p1, resp, ",")
                lon = Mid(resp, p1, p2 - p1 - 1)

                ws.Cells(i, 2).Value = lat
                ws.Cells(i, 3).Value = lon
                ws.Cells(i, 4).Value = "OK"
                ws.Cells(i, 4).Interior.Color = RGB(0, 176, 80)
            Else
                ws.Cells(i, 4).Value = "No"
                ws.Cells(i, 4).Interior.Color = RGB(192, 0, 0)
            End If

            Application.Wait Now + TimeValue("00:00:02")
        End If
    Next i

fin:
    MsgBox "Listo"
End Sub
