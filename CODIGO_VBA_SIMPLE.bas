'=======================================================
' VERSIÓN ULTRA SIMPLE - Sin manejo de JSON
' Usa búsqueda por texto simple
'=======================================================

Option Explicit

Sub Geocodificar()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim direccion As String
    Dim contador As Long

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets("Geocodificación")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    contador = 0

    For row = 2 To lastRow
        direccion = Trim(ws.Cells(row, 1).Value)

        If Len(direccion) > 0 Then
            ws.Cells(row, 4).Value = "⏳ Buscando..."
            DoEvents

            If ObtenerCoordenadas(direccion, _
                ws.Cells(row, 2), ws.Cells(row, 3)) Then
                ws.Cells(row, 4).Value = "✓ OK"
                ws.Cells(row, 4).Interior.Color = RGB(0, 176, 80)
                contador = contador + 1
            Else
                ws.Cells(row, 2).Value = "N/A"
                ws.Cells(row, 3).Value = "N/A"
                ws.Cells(row, 4).Value = "✗ No encontrada"
                ws.Cells(row, 4).Interior.Color = RGB(192, 0, 0)
            End If

            Application.Wait Now + TimeValue("00:00:02")
        End If
    Next row

    MsgBox "Listo: " & contador & " de " & (lastRow - 1), vbInformation

    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Function ObtenerCoordenadas(dir As String, LatCell As Range, LonCell As Range) As Boolean
    Dim xmlhttp As Object
    Dim url As String
    Dim response As String
    Dim latStart As Long
    Dim lonStart As Long
    Dim latEnd As Long
    Dim lonEnd As Long
    Dim lat As String
    Dim lon As String

    On Error GoTo ErrorFunc

    Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")

    ' URL con dirección codificada
    url = "https://nominatim.openstreetmap.org/search?q=" & dir & "&format=json&limit=1"

    xmlhttp.Open "GET", url, False
    xmlhttp.SetRequestHeader "User-Agent", "Excel"
    xmlhttp.Send

    response = xmlhttp.responseText

    ' Si no hay respuesta
    If InStr(response, """lat""") = 0 Then
        ObtenerCoordenadas = False
        Exit Function
    End If

    ' Extraer latitud
    latStart = InStr(response, """lat"":""") + 8
    latEnd = InStr(latStart, response, """")
    If latEnd > latStart Then
        lat = Mid(response, latStart, latEnd - latStart)
        LatCell.Value = lat
    End If

    ' Extraer longitud
    lonStart = InStr(response, """lon"":""") + 8
    lonEnd = InStr(lonStart, response, """")
    If lonEnd > lonStart Then
        lon = Mid(response, lonStart, lonEnd - lonStart)
        LonCell.Value = lon
    End If

    If Len(lat) > 0 And Len(lon) > 0 Then
        ObtenerCoordenadas = True
    Else
        ObtenerCoordenadas = False
    End If

    Exit Function
ErrorFunc:
    ObtenerCoordenadas = False
End Function

Sub Auto_Open()
    On Error Resume Next
    Application.OnKey "+^G", "Geocodificar"
End Sub
