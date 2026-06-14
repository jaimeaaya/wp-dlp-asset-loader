' ============================================================
' ConsolidarHojas — Macro VBA para wp-dlp-asset-loader
' ============================================================
' Consolida todas las hojas de marca en la hoja "Consolidado".
'
' ESTRUCTURA DE HOJAS FUENTE (Hero, TheraBreath, Batiste, etc.)
'   Col A  UBICACION EN DB (provisional)  — se omite
'   Col B  NOMBRE
'   Col C  CATEGORIA
'   Col D  SUBCATEGORIA (Solo si aplica)
'   Col E  DESCRIPCION
'   Col F  FORMATO
'   Col G  IMAGEN PRINCIPAL (DROPBOX)
'   Col H  PREVIEW CONTENIDOS (DROPBOX)
'   Col I  EDITABLES
'   Col J  PIEZAS
'   Col K  CONSOLIDADO
'
' ESTRUCTURA DE HOJA CONSOLIDADO (resultado)
'   Col A  NOMBRE
'   Col B  CATEGORIA
'   Col C  SUBCATEGORIA (Solo si aplica)
'   Col D  DESCRIPCION
'   Col E  FORMATO
'   Col F  IMAGEN PRINCIPAL (DROPBOX)
'   Col G  PREVIEW CONTENIDOS (DROPBOX)
'   Col H  EDITABLES
'   Col I  PIEZAS
'   Col J  CONSOLIDADO
'   Col K  Hoja             (nombre de la hoja fuente — agrega la macro)
'   Col L  COMPARACION EDITABLE PIEZA CONSOLIDADO (formula opcional)
'
' USO:
'   1. Abrir el archivo .xlsm en Excel
'   2. Ir a Desarrollador > Macros > ConsolidarHojas > Ejecutar
'      O presionar Alt+F8 y seleccionar ConsolidarHojas
'   3. La hoja "Consolidado" se actualiza limpiando datos previos
'
' NOTAS:
'   - Hojas excluidas: "Consolidado" y "CRONOGRAMA"
'   - Se procesan todas las demas hojas en el orden en que aparecen
'   - Filas completamente vacias se omiten
'   - Guardar el archivo como .xlsm para preservar la macro
' ============================================================

Sub ConsolidarHojas()

    Dim wsConsolidado As Worksheet
    Dim wsOrigen      As Worksheet
    Dim ultimaFilaOrigen As Long
    Dim ultimaFilaDest   As Long
    Dim filaOrigen       As Long
    Dim i                As Long

    ' Hojas que se excluyen del consolidado
    Dim hojasExcluidas(1) As String
    hojasExcluidas(0) = "Consolidado"
    hojasExcluidas(1) = "CRONOGRAMA"

    ' Apagar actualizacion de pantalla para mayor velocidad
    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual

    Set wsConsolidado = ThisWorkbook.Worksheets("Consolidado")

    ' -- Limpiar datos previos (mantener cabecera en fila 1) ----------------
    If wsConsolidado.Cells(2, 1).Value <> "" Then
        Dim ultimaFilaLimpieza As Long
        ultimaFilaLimpieza = wsConsolidado.Cells(wsConsolidado.Rows.Count, 1).End(xlUp).Row
        If ultimaFilaLimpieza >= 2 Then
            wsConsolidado.Rows("2:" & ultimaFilaLimpieza).Delete Shift:=xlUp
        End If
    End If

    ' -- Escribir cabecera si no existe -------------------------------------
    Dim cabecera(11) As String
    cabecera(0)  = "NOMBRE"
    cabecera(1)  = "CATEGORIA"
    cabecera(2)  = "SUBCATEGORIA (Solo si aplica)"
    cabecera(3)  = "DESCRIPCION"
    cabecera(4)  = "FORMATO"
    cabecera(5)  = "IMAGEN PRINCIPAL (DROPBOX)"
    cabecera(6)  = "PREVIEW CONTENIDOS (DROPBOX)"
    cabecera(7)  = "EDITABLES"
    cabecera(8)  = "PIEZAS"
    cabecera(9)  = "CONSOLIDADO"
    cabecera(10) = "Hoja"
    cabecera(11) = "COMPARACION EDITABLE PIEZA CONSOLIDADO"

    For i = 0 To 11
        wsConsolidado.Cells(1, i + 1).Value = cabecera(i)
    Next i

    ' -- Recorrer cada hoja fuente -----------------------------------------
    For Each wsOrigen In ThisWorkbook.Worksheets

        ' Saltar hojas excluidas
        Dim esExcluida As Boolean
        esExcluida = False
        For i = 0 To UBound(hojasExcluidas)
            If wsOrigen.Name = hojasExcluidas(i) Then
                esExcluida = True
                Exit For
            End If
        Next i
        If esExcluida Then GoTo SiguienteHoja

        ' Encontrar ultima fila con datos en la hoja fuente
        ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, 2).End(xlUp).Row
        If ultimaFilaOrigen < 2 Then GoTo SiguienteHoja  ' hoja vacia

        ' Copiar fila a fila (saltando la cabecera de la hoja fuente)
        For filaOrigen = 2 To ultimaFilaOrigen

            ' Saltar filas completamente vacias
            If WorksheetFunction.CountA(wsOrigen.Rows(filaOrigen)) = 0 Then
                GoTo SiguienteFila
            End If

            ' Proxima fila libre en Consolidado
            ultimaFilaDest = wsConsolidado.Cells(wsConsolidado.Rows.Count, 1).End(xlUp).Row + 1

            ' Copiar columnas B a K de la fuente -> A a J del Consolidado
            wsConsolidado.Cells(ultimaFilaDest, 1).Value  = wsOrigen.Cells(filaOrigen, 2).Value   ' NOMBRE
            wsConsolidado.Cells(ultimaFilaDest, 2).Value  = wsOrigen.Cells(filaOrigen, 3).Value   ' CATEGORIA
            wsConsolidado.Cells(ultimaFilaDest, 3).Value  = wsOrigen.Cells(filaOrigen, 4).Value   ' SUBCATEGORIA
            wsConsolidado.Cells(ultimaFilaDest, 4).Value  = wsOrigen.Cells(filaOrigen, 5).Value   ' DESCRIPCION
            wsConsolidado.Cells(ultimaFilaDest, 5).Value  = wsOrigen.Cells(filaOrigen, 6).Value   ' FORMATO
            wsConsolidado.Cells(ultimaFilaDest, 6).Value  = wsOrigen.Cells(filaOrigen, 7).Value   ' IMAGEN PRINCIPAL
            wsConsolidado.Cells(ultimaFilaDest, 7).Value  = wsOrigen.Cells(filaOrigen, 8).Value   ' PREVIEW CONTENIDOS
            wsConsolidado.Cells(ultimaFilaDest, 8).Value  = wsOrigen.Cells(filaOrigen, 9).Value   ' EDITABLES
            wsConsolidado.Cells(ultimaFilaDest, 9).Value  = wsOrigen.Cells(filaOrigen, 10).Value  ' PIEZAS
            wsConsolidado.Cells(ultimaFilaDest, 10).Value = wsOrigen.Cells(filaOrigen, 11).Value  ' CONSOLIDADO
            wsConsolidado.Cells(ultimaFilaDest, 11).Value = wsOrigen.Name                         ' Hoja (nombre fuente)

SiguienteFila:
        Next filaOrigen

SiguienteHoja:
    Next wsOrigen

    ' -- Restaurar configuracion -------------------------------------------
    Application.Calculation    = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Consolidado actualizado." & vbNewLine & _
           "Filas: " & (wsConsolidado.Cells(wsConsolidado.Rows.Count, 1).End(xlUp).Row - 1), _
           vbInformation, "ConsolidarHojas"

End Sub
