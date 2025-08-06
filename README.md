# 📊 Automatización de Reportes en Excel con Macros (VBA)

<img width="1773" height="1010" alt="image" src="https://github.com/user-attachments/assets/6665c213-3d48-4b94-8567-cc5eabc21fe0" />

Este proyecto consiste en una herramienta desarrollada en **Excel con Macros (VBA)** que permite automatizar completamente el proceso de:

- Consolidar múltiples archivos de ventas.
- Calcular KPIs clave de forma automática.
- Generar una tabla dinámica y un gráfico.
- Ejecutar todo el proceso con un solo botón.

Diseñada como una solución práctica para **Control de Gestión**, esta herramienta es especialmente útil para reportes mensuales o semanales.

---

## 🗂️ Datos utilizados

Se trabajó con tres archivos de ventas ficticios correspondientes a los meses de **enero, febrero y marzo**, ubicados en la carpeta `/Datos_Ventas`.

---

## ⚙️ Paso a paso del proyecto

### 1. Macro de Consolidación de Archivos

Se desarrolló una macro que recorre automáticamente todos los archivos de la carpeta y **consolida la información en una sola hoja de cálculo**, unificando las ventas de los tres meses.

```vba
Sub ConsolidarVentas()
    Dim Carpeta As String, Archivo As String
    Dim LibroOrigen As Workbook
    Dim HojaDestino As Worksheet
    Dim UltimaFila As Long, FilaInicio As Long

    ' Ruta de la carpeta
    Carpeta = ThisWorkbook.Path & "\Datos_Ventas\"
    Archivo = Dir(Carpeta & "*.xlsm")

    ' Hoja donde irá todo
    Set HojaDestino = ThisWorkbook.Sheets("Ventas")
    HojaDestino.Cells.Clear
    FilaInicio = 2

    ' Copiar encabezado
    HojaDestino.Range("A1:G1").Value = Array("Fecha", "Región", "Producto", "Unidades", "Precio", "Descuento", "Costo Unitario")

    Do While Archivo <> ""
        Set LibroOrigen = Workbooks.Open(Carpeta & Archivo)
        With LibroOrigen.Sheets("Ventas")
            UltimaFila = .Cells(.Rows.Count, "A").End(xlUp).Row
            .Range("A2:G" & UltimaFila).Copy HojaDestino.Cells(FilaInicio, 1)
            FilaInicio = HojaDestino.Cells(HojaDestino.Rows.Count, "A").End(xlUp).Row + 1
        End With
        LibroOrigen.Close SaveChanges:=False
        Archivo = Dir
    Loop
    
        ' Agregar encabezados
    HojaDestino.Range("H1").Value = "Margen"
    HojaDestino.Range("I1").Value = "Ingreso Total"
    HojaDestino.Range("J1").Value = "Descuento Total"
    
    ' Aplicar fórmulas
    Dim UltimaFilaDatos As Long
    UltimaFilaDatos = HojaDestino.Cells(HojaDestino.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To UltimaFilaDatos
        ' Margen
        HojaDestino.Cells(i, "H").Formula = "=(E" & i & " - G" & i & ") / E" & i
        ' Ingreso Total
        HojaDestino.Cells(i, "I").Formula = "=E" & i & "*D" & i
        ' Descuento Total
        HojaDestino.Cells(i, "J").Formula = "=E" & i & "*F" & i & "*D" & i
    Next i

    MsgBox "Consolidación completada", vbInformation
    
End Sub
```

<img width="1452" height="776" alt="image" src="https://github.com/user-attachments/assets/ab465ec8-b543-4492-aba2-c39d07a5f9a1" />


### 2. Macro de Tabla Dinámica

Se creó una macro que genera **una tabla dinámica automáticamente** con las métricas relevantes a partir de los datos consolidados. Esta tabla sirve como base para el análisis.

```vba
Sub CrearTablaDinamica()
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim wsDatos As Worksheet
    Dim wsResumen As Worksheet
    Dim UltimaFila As Long
    Dim RangoDatos As Range

    ' Set hojas
    Set wsDatos = ThisWorkbook.Sheets("Ventas")

    ' Crear hoja para resumen
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Resumen").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsResumen = ThisWorkbook.Sheets.Add
    wsResumen.Name = "Resumen"

    ' Determinar última fila
    UltimaFila = wsDatos.Cells(wsDatos.Rows.Count, "A").End(xlUp).Row
    Set RangoDatos = wsDatos.Range(wsDatos.Cells(1, 1), wsDatos.Cells(UltimaFila, 10)) ' Columnas A a J

    ' Crear caché
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=RangoDatos)

    ' Crear tabla dinámica
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=wsResumen.Range("A3"), _
        TableName:="VentasPorRegion_" & Format(Now, "hhmmss"))

    ' Esperar breve momento (opcional, para estabilidad)
    Application.Wait Now + TimeValue("0:00:01")

    ' Configurar tabla dinámica
    With pt
        .PivotFields("Región").Orientation = xlRowField
        With .PivotFields("Ingreso Total")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "#,##0"
        End With
    End With
End Sub
```

<img width="1497" height="816" alt="image" src="https://github.com/user-attachments/assets/dbc73ba9-cd14-48ba-be4a-fe8994aaf98c" />


### 3. Macro de Gráfico Dinámico

A partir de la tabla dinámica, se automatizó la **generación de un gráfico de barras**, lo que permite visualizar los resultados de forma clara e inmediata.

```vba
Sub CrearGraficoDinamico()
    Dim wsResumen As Worksheet
    Dim pt As PivotTable
    Dim grafico As ChartObject

    Set wsResumen = ThisWorkbook.Sheets("Resumen")
    Set pt = wsResumen.PivotTables(1) ' Usa la primera tabla dinámica

    ' Borrar gráficos existentes si los hay
    For Each grafico In wsResumen.ChartObjects
        grafico.Delete
    Next grafico

    ' Insertar nuevo gráfico
    Set grafico = wsResumen.ChartObjects.Add(Left:=300, Top:=30, Width:=500, Height:=300)
    With grafico.Chart
        .SetSourceData Source:=pt.TableRange1
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Ventas Totales por Región"
        .ApplyLayout (4)
    End With
End Sub
```

<img width="1125" height="508" alt="image" src="https://github.com/user-attachments/assets/92660e90-3ca9-4f01-a893-bcd24c7e5606" />


### 4. Macro de Automatización Total

Se desarrolló una macro principal que permite ejecutar todo el proceso anterior con **un solo botón interactivo**, facilitando el uso para usuarios sin conocimientos de VBA.

```vba
Sub GenerarReporteCompleto()
    Call ConsolidarVentas
    Call CrearTablaDinamica
    Call CrearGraficoDinamico
End Sub
```

<img width="431" height="239" alt="image" src="https://github.com/user-attachments/assets/9704589b-f20c-49d2-b932-cf751a01d2d2" />


---

## 🧠 Aplicación práctica

Este proyecto demuestra cómo aplicar **automatización con VBA en tareas reales de control de gestión**, permitiendo:

- Ahorro significativo de tiempo.
- Reducción de errores manuales.
- Estandarización de reportes.

---

## 📁 Archivos incluidos

- `macro_reportes.xlsm`: archivo principal con macros y botón de ejecución.
- Carpeta `/Datos_Ventas`: contiene los archivos de ventas mensuales.
- `README.md`: descripción detallada del proyecto.

---

## 🧰 Tecnologías utilizadas

- Microsoft Excel
- VBA (Visual Basic for Applications)
