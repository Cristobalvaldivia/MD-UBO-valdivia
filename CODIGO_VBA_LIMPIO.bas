'=========================================================
' GEOCODIFICADOR EXCEL - VERSION LIMPIA SIN ACENTOS
' Simple y sin caracteres especiales
' Presiona Ctrl+Shift+G para ejecutar
'=========================================================

Option Explicit

Sub Geocodificar()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim direccion As String
    Dim lat As String
    Dim lon As String
    Dim contador As Long

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets("Geocodificacion")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    contador = 0

    For row = 2 To lastRow
        direccion = Trim(ws.Cells(row, 1).Value)

        If Len(direccion) > 0 Then
            ws.Cells(row, 4).Value = "Procesando..."
            ws.Cells(row, 4).Interior.Color = RGB(255, 255, 0)
            DoEvents

            If BuscarCoordenadas(direccion, lat, lon) Then
                ws.Cells(row, 2).Value = lat
                ws.Cells(row, 3).Value = lon
                ws.Cells(row, 4).Value = "OK"
                ws.Cells(row, 4).Interior.Color = RGB(0, 176, 80)
                contador = contador + 1
            Else
                ws.Cells(row, 2).Value = "N/A"
                ws.Cells(row, 3).Value = "N/A"
                ws.Cells(row, 4).Value = "No encontrada"
                ws.Cells(row, 4).Interior.Color = RGB(192, 0, 0)
            End If

            Application.Wait Now + TimeValue("00:00:02")
        End If
    Next row

    MsgBox "Geocodificacion completada. Encontradas: " & contador & " de " & (lastRow - 1), vbInformation

    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Function BuscarCoordenadas(direccion As String, ByRef latitud As String, ByRef longitud As String) As Boolean
    Dim xmlhttp As Object
    Dim url As String
    Dim response As String
    Dim latPos As Long
    Dim lonPos As Long
    Dim latValue As String
    Dim lonValue As String

    On Error GoTo ErrorBuscar

    Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")

    url = "https://nominatim.openstreetmap.org/search?q=" & direccion & "&format=json&limit=1"

    xmlhttp.Open "GET", url, False
    xmlhttp.SetRequestHeader "User-Agent", "Mozilla/5.0"
    xmlhttp.Send

    response = xmlhttp.responseText

    If Len(response) = 0 Or response = "[]" Then
        BuscarCoordenadas = False
        Exit Function
    End If

    latPos = InStr(response, """lat"":")
    If latPos > 0 Then
        latValue = ExtractValor(response, latPos)
        If Len(latValue) > 0 Then
            latitud = latValue
        Else
            BuscarCoordenadas = False
            Exit Function
        End If
    Else
        BuscarCoordenadas = False
        Exit Function
    End If

    lonPos = InStr(response, """lon"":")
    If lonPos > 0 Then
        lonValue = ExtractValor(response, lonPos)
        If Len(lonValue) > 0 Then
            longitud = lonValue
        Else
            BuscarCoordenadas = False
            Exit Function
        End If
    Else
        BuscarCoordenadas = False
        Exit Function
    End If

    BuscarCoordenadas = True
    Exit Function

ErrorBuscar:
    BuscarCoordenadas = False
End Function

Function ExtractValor(jsonStr As String, startPos As Long) As String
    Dim i As Long
    Dim valor As String
    Dim charCount As Long
    Dim inQuote As Boolean
    Dim char As String

    i = startPos

    Do While i < Len(jsonStr)
        i = i + 1
        char = Mid(jsonStr, i, 1)

        If char = """" Then
            inQuote = True
            Exit Do
        ElseIf (char >= "0" And char <= "9") Or char = "-" Then
            Exit Do
        End If
    Loop

    valor = ""
    charCount = 0

    Do While i < Len(jsonStr) And charCount < 20
        char = Mid(jsonStr, i, 1)

        If inQuote Then
            If char = """" Then Exit Do
        Else
            If char = "," Or char = "}" Or char = "]" Then Exit Do
        End If

        valor = valor & char
        i = i + 1
        charCount = charCount + 1
    Loop

    valor = Trim(Replace(valor, """", ""))
    ExtractValor = valor
End Function

Sub Auto_Open()
    On Error Resume Next
    Application.OnKey "+^G", "Geocodificar"
End Sub
