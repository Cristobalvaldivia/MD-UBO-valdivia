Sub Geocodificar()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dir As String
    Dim lat As String
    Dim lon As String

    On Error GoTo fin

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        dir = ws.Cells(i, 1).Value

        If dir <> "" Then
            ws.Cells(i, 4).Value = "Buscando..."
            DoEvents

            If BuscarDireccion(dir, lat, lon) Then
                ws.Cells(i, 2).Value = lat
                ws.Cells(i, 3).Value = lon
                ws.Cells(i, 4).Value = "OK"
                ws.Cells(i, 4).Interior.Color = RGB(0, 176, 80)
            Else
                ws.Cells(i, 2).Value = "N/A"
                ws.Cells(i, 3).Value = "N/A"
                ws.Cells(i, 4).Value = "No"
                ws.Cells(i, 4).Interior.Color = RGB(192, 0, 0)
            End If

            Application.Wait Now + TimeValue("00:00:01")
        End If
    Next i

fin:
    MsgBox "Geocodificacion completa"
End Sub

Function BuscarDireccion(dir As String, ByRef lat As String, ByRef lon As String) As Boolean
    Dim http As Object
    Dim url As String
    Dim resp As String
    Dim p1 As Long
    Dim p2 As Long
    Dim busquedas As Variant
    Dim i As Long

    On Error GoTo ErrorBuscar

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    busquedas = Array(dir, ExtractCity(dir), "Puerto Varas Chile", "Los Lagos Chile")

    For i = LBound(busquedas) To UBound(busquedas)
        If busquedas(i) <> "" Then
            url = "https://nominatim.openstreetmap.org/search?q=" & busquedas(i) & "&format=json&limit=1"

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

                BuscarDireccion = True
                Exit Function
            End If
        End If
    Next i

    BuscarDireccion = False
    Exit Function

ErrorBuscar:
    BuscarDireccion = False
End Function

Function ExtractCity(dir As String) As String
    Dim partes As Variant
    Dim i As Long
    Dim ciudad As String

    partes = Split(dir, ",")

    For i = LBound(partes) To UBound(partes)
        If InStr(partes(i), "Puerto Varas") > 0 Or _
           InStr(partes(i), "Los Muermos") > 0 Or _
           InStr(partes(i), "Nueva Braunau") > 0 Then
            ExtractCity = Trim(partes(i))
            Exit Function
        End If
    Next i

    ExtractCity = ""
End Function
