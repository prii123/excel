# Combinar Hojas de Excel

Sub CombinaHojas()
    Dim ws As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFila As Long
    Dim ultimaFilaDestino As Long
    Dim i As Integer
    
    ' Crea una nueva hoja llamada "Combinada" para combinar todas las hojas
    Set wsDestino = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDestino.Name = "Combinada"
    ultimaFilaDestino = 1
    
    ' Recorre todas las hojas en el libro
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> wsDestino.Name Then ' Evita combinar la hoja "Combinada" consigo misma
            ' Encuentra la última fila ocupada en la hoja de destino
            ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Copia los datos desde la hoja actual a la hoja de destino
            'ws.Range("A1").EntireRow.Copy Destination:=wsDestino.Cells(ultimaFilaDestino, 1)
            ws.Range("A1:Z100").EntireRow.Copy Destination:=wsDestino.Cells(ultimaFilaDestino, 1)
            
            ' Actualiza la última fila de la hoja de destino
            ultimaFilaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1
        End If
    Next ws
    
    ' Ajusta el ancho de las columnas en la hoja de destino para que se vean bien los datos
    wsDestino.Columns.AutoFit
    
    MsgBox "¡Se han combinado todas las hojas en una sola hoja llamada 'Combinada'!", vbInformation
End Sub
