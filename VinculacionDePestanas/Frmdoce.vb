Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Win32
Public Class Frmdoce
    Dim RutaArchivo As String
    Dim consulta, cadena As String
    Dim longitud As Double
    'Dim Nfila, Folio, Magnitud, Informe, ClaveEmpresa, Empresa, ClaveContacto, Contacto, Usuario,
    'FechaRecepcion, FechaCalibracion, FechaEmision, ServCatalogo1, ServCatalogo2, ServiciosAdicionales, PUCalib, PULab, PUFacturado, Tipo, Alcance, Marca, Modelo, Serie, ID, Accesorios, Puntos,
    'FuncionesCalibradas, Etiquetas, DatosdelInforme, Observaciones, Calibro, NumCot, Status, RealizoMedicion, EmpresaEmision, Calle, Colonia, Ciudad, Estado, Pais, CP, Calendarizacion, Politicadeajuste, Intervalodecalibracion,
    'Evaluaciondelaconformidad, Curvadeajuste, Mantenimiento, Idioma, SVAD10, FechadeRegistro, FechadeRecepcionLab,
    'patron1, patron2, patron3, patron4, patron5, patron6, patron7, patron8, patron9, patron10, firma, observacionStatus As String

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            Dim ofd As OpenFileDialog = New OpenFileDialog
            ofd.DefaultExt = "txt"
            ofd.FileName = "defaultname"
            ofd.InitialDirectory = "c:\"
            ofd.Filter = "All files|*.*|Text files|*.txt"
            ofd.Title = "Select file"
            If ofd.ShowDialog() <> DialogResult.Cancel Then
                RutaArchivo = ofd.FileName
                Label2.Text = RutaArchivo
            End If
        Catch ex As Exception
            MsgBox("No se pudo leer el archivo.", MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            MetodoMetasInf2019()
            Dim R As String
            R = "SELECT TOP 1 * FROM [INFORMES-SERVICIOS]"
            Dim dAdap As New SqlDataAdapter(R, conexion2019)
            Dim ds As New DataSet
            dAdap.Fill(ds)
            For Each dc As DataColumn In ds.Tables(0).Columns
                DataGridView1.Rows.Add(dc.ColumnName)
            Next
        Catch ex As Exception
            MsgBox("Error de lectura de datos.", MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Try
            Dim ofd As OpenFileDialog = New OpenFileDialog
            ofd.DefaultExt = "txt"
            ofd.FileName = "defaultname"
            ofd.InitialDirectory = "c:\"
            ofd.Filter = "All files|*.*|Text files|*.txt"
            ofd.Title = "Select file"
            If ofd.ShowDialog() <> DialogResult.Cancel Then
                RutaArchivo = ofd.FileName
                Label2.Text = RutaArchivo
            End If
        Catch ex As Exception
            MsgBox("No se pudo leer el archivo.", MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ''----------------------exportar consulta a excel------------------------------------------------------------------
            Dim Aplicacion As New Excel.Application
            Dim Libro As Excel.Workbook
            Dim Hoja As Excel.Worksheet
            Aplicacion = New Excel.Application
            Libro = Aplicacion.Workbooks.Open(RutaArchivo)
            Hoja = Libro.Worksheets("Importados")
            ''se escribe en excel, el nombre de las columnas de SQL---------------
            For i = 1 To DataGridView1.Rows.Count - 2
                Hoja.Cells(i + 4, 8) = "[" & DataGridView1.Item(0, i).Value & "]"
            Next i
            ''--------------------------------------------------------------------
            ''-------consultar los campos que laboratorio agrego al excel en la columna A---
            For i = 1 To 96
                ' DataGridView2.Rows.Add(Libro.Sheets(1).Cells(i + 4, 1).Value)
                'DataGridView2.Rows.Add(Hoja.Cells(i + 4, 1).Value)
                DataGridView2.Rows.Add(Libro.Sheets("Importados").Cells(i + 4, 1).Value)
            Next i
            ''-------------------------------------------------------------------------------
            ''Codigo para hacer el recorrido de los campos que se requieren-------------------------------------------
            Dim conteo As Integer
            For i = 0 To DataGridView2.Rows.Count - 2
                If DataGridView2.Item(0, i).Value = "" Then
                    'MsgBox("Espacio en blanco, fin ")
                    'Exit For
                    cadena = cadena
                Else
                    conteo = conteo + 1
                    cadena = cadena & DataGridView2.Item(0, i).Value & ","
                End If
            Next i
            comando2019 = conexion2019.CreateCommand()
            cadena = cadena.Substring(0, Len(cadena) - 1)
            cadena = "select " + cadena + " from [INFORMES-SERVICIOS] where MAGNITUD='" & txtMagnitud.Text & "' and INFORME='" & txtInforme.Text & "'"
            'MsgBox(cadena)
            comando2019.CommandText = cadena
            lector2019 = comando2019.ExecuteReader
            lector2019.Read()
            Dim renglon As Integer = 0
            For i = 0 To DataGridView2.Rows.Count - 2

                If Not DataGridView2.Item(0, i).Value = "" Then
                    If renglon <= conteo Then
                        DataGridView2.Item(1, i).Value = lector2019(renglon)
                        renglon = renglon + 1
                    End If
                End If
            Next i
            '---------------------------------------------------------------------------------------------------------------
            ''se escribe en excel, el resultado de la consulta SQL--------------
            For i = 0 To DataGridView2.Rows.Count - 2
                Hoja.Cells(i + 5, 2) = DataGridView2.Item(1, i).Value
            Next i
            ''--------------------------------------------------------------------

            ''---------------------------------
            '' Codigo que fuese
            ''---------------------------------
            ''CerrarCOM(oExcel)
            ''Hoja = Nothing
            ''Libro = Nothing
            ''Aplicacion = Nothing
            ''GC.Collect()
            ''ClearMemory()
            ''Process.Start(oFileName)


            ''CerrarCOM(oExcel)
            ''xlSheet = Nothing
            ''xlBook = Nothing
            ''oExcel = Nothing
            ''GC.Collect()
            ''ClearMemory()

            ''----------------------------------
            Dim ms_Excel As New Excel.Application()
            Dim wbook As Excel.Workbook = ms_Excel.Workbooks.Open(RutaArchivo) ' abre el libro por ende su proceso

            ms_Excel.Visible = False

            wbook.Saved = False
            wbook.Close() ' cierra el clibro y el proceso que este genero.
            ms_Excel.Quit()
            ms_Excel = Nothing
            '--------------------------------------------------------

            Aplicacion.Sheets(1).Protect(password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True) 'Quita la Protección del Archivo
            Aplicacion.DisplayAlerts = False 'Elimina Los Mensajes De Alerta
            Aplicacion.ActiveWorkbook.Save() 'Guarda los Cambios 
            Aplicacion.Visible = True
            Aplicacion.ActiveWorkbook.ActiveSheet.PrintPreview()
            Aplicacion.Quit()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(Aplicacion)
            Aplicacion = Nothing
            ''----


            Libro.Close()
            Aplicacion.Quit()
            'releaseObject(Aplicacion)
            'releaseObject(Libro)
            'releaseObject(Hoja)
            MsgBox("Datos cargados correctamente.", MsgBoxStyle.Information)
            ''-------------------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            MsgBox("Palabras reservadas no encotradas, verifica tu lista de palabras.", MsgBoxStyle.Critical)
        End Try
    End Sub

    ''    Friend Function CerrarCOM(ByRef oComObject As Object) As Boolean
    ''        Try
    ''            If oComObject Is Nothing = False Then
    ''                System.Runtime.InteropServices.Marshal.ReleaseComObject(oComObject)
    ''            End If
    ''            Return True
    ''        Catch ex As Exception
    ''            Return False
    ''        End Try
    ''    End Function


    ''#Region "Liberar Memoria no utilizada"
    ''    'Funcion de liberacion de memoria
    ''    Public Sub ClearMemory()
    ''        Try
    ''            Dim Mem As Process
    ''            Mem = Process.GetCurrentProcess()
    ''            SetProcessWorkingSetSize(Mem.Handle, -1, -1)
    ''        Catch ex As Exception
    ''            'Control de errores
    ''        End Try
    ''    End Sub

    ''#End Region

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim Aplicacion As New Excel.Application
            Dim Libro As Excel.Workbook
            Dim Hoja As Excel.Worksheet
            Aplicacion = New Excel.Application
            Libro = Aplicacion.Workbooks.Open(RutaArchivo)
            Hoja = Libro.Worksheets("Importados")

            ''se escribe en excel, el nombre de las columnas de SQL---------------
            For i = 1 To DataGridView1.Rows.Count - 2
                Hoja.Cells(i + 4, 8) = "[" & DataGridView1.Item(0, i).Value & "]"
            Next i
            ''--------------------------------------------------------------------

            Libro.Close()
            Aplicacion.Quit()
            'releaseObject(Aplicacion)
            'releaseObject(Libro)
            'releaseObject(Hoja)
            MsgBox("Columnas Actualizadas", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("No se pudo leer el archivo.", MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Me.Dispose()
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    ''Public Sub busquedaDeDatos()
    ''    Try
    ''        Dim R As String
    ''        R = "select * from [INFORMES-SERVICIOS] where MAGNITUD='" & txtMagnitud.Text & "'and INFORME='" & txtInforme.Text & "'"
    ''        Dim comando As New SqlCommand(R, conexion2019)
    ''        Dim lector As SqlDataReader
    ''        lector = comando.ExecuteReader
    ''        lector.Read()

    ''        If ((lector(0) Is DBNull.Value) OrElse (lector(0) Is Nothing)) Then
    ''            Nfila = 0
    ''        Else
    ''            Nfila = lector(0)
    ''        End If

    ''        If ((lector(1) Is DBNull.Value) OrElse (lector(1) Is Nothing)) Then
    ''            Folio = 0
    ''        Else
    ''            Folio = lector(1)
    ''        End If

    ''        If ((lector(2) Is DBNull.Value) OrElse (lector(2) Is Nothing)) Then
    ''            Magnitud = ""
    ''        Else
    ''            Magnitud = lector(2)
    ''        End If

    ''        If ((lector(3) Is DBNull.Value) OrElse (lector(3) Is Nothing)) Then

    ''            Informe = ""
    ''        Else
    ''            Informe = lector(3)
    ''        End If

    ''        If ((lector(5) Is DBNull.Value) OrElse (lector(5) Is Nothing)) Then
    ''            ClaveEmpresa = 0
    ''        Else
    ''            ClaveEmpresa = lector(5)
    ''        End If

    ''        If ((lector(6) Is DBNull.Value) OrElse (lector(6) Is Nothing)) Then
    ''            Empresa = ""
    ''        Else
    ''            Empresa = lector(6)
    ''        End If

    ''        If ((lector(7) Is DBNull.Value) OrElse (lector(7) Is Nothing)) Then
    ''            ClaveContacto = 0
    ''        Else
    ''            ClaveContacto = lector(7)
    ''        End If

    ''        If ((lector(8) Is DBNull.Value) OrElse (lector(8) Is Nothing)) Then
    ''            Usuario = ""
    ''        Else
    ''            Usuario = lector(8)
    ''        End If

    ''        If ((lector(10) Is DBNull.Value) OrElse (lector(10) Is Nothing)) Then
    ''            FechaRecepcion = ""
    ''        Else
    ''            FechaRecepcion = lector(10)
    ''        End If

    ''        If ((lector(11) Is DBNull.Value) OrElse (lector(11) Is Nothing)) Then
    ''            FechaCalibracion = ""
    ''        Else
    ''            FechaCalibracion = lector(11)
    ''        End If

    ''        If ((lector(12) Is DBNull.Value) OrElse (lector(12) Is Nothing)) Then
    ''            FechaEmision = ""
    ''        Else
    ''            FechaEmision = lector(12)
    ''        End If

    ''        If ((lector(13) Is DBNull.Value) OrElse (lector(13) Is Nothing)) Then
    ''            ServCatalogo1 = ""
    ''        Else
    ''            ServCatalogo1 = lector(13)
    ''        End If

    ''        If ((lector(14) Is DBNull.Value) OrElse (lector(14) Is Nothing)) Then
    ''            ServCatalogo2 = ""
    ''        Else
    ''            ServCatalogo2 = lector(14)
    ''        End If

    ''        If ((lector(15) Is DBNull.Value) OrElse (lector(15) Is Nothing)) Then
    ''            ServiciosAdicionales = ""
    ''        Else
    ''            ServiciosAdicionales = lector(15)
    ''        End If

    ''        If ((lector(16) Is DBNull.Value) OrElse (lector(16) Is Nothing)) Then
    ''            PUCalib = 0
    ''        Else
    ''            PUCalib = lector(16)
    ''        End If

    ''        If ((lector(17) Is DBNull.Value) OrElse (lector(17) Is Nothing)) Then
    ''            PULab = 0
    ''        Else
    ''            PULab = lector(17)
    ''        End If

    ''        If ((lector(18) Is DBNull.Value) OrElse (lector(18) Is Nothing)) Then
    ''            PUFacturado = 0
    ''        Else
    ''            PUFacturado = lector(18)
    ''        End If

    ''        If ((lector(19) Is DBNull.Value) OrElse (lector(19) Is Nothing)) Then
    ''            Tipo = ""
    ''        Else
    ''            Tipo = lector(19)
    ''        End If

    ''        If ((lector(20) Is DBNull.Value) OrElse (lector(20) Is Nothing)) Then
    ''            Alcance = ""
    ''        Else
    ''            Alcance = lector(20)
    ''        End If

    ''        If ((lector(21) Is DBNull.Value) OrElse (lector(21) Is Nothing)) Then
    ''            Marca = ""
    ''        Else
    ''            Marca = lector(21)
    ''        End If

    ''        If ((lector(22) Is DBNull.Value) OrElse (lector(22) Is Nothing)) Then
    ''            Modelo = ""
    ''        Else
    ''            Modelo = lector(22)
    ''        End If

    ''        If ((lector(23) Is DBNull.Value) OrElse (lector(23) Is Nothing)) Then
    ''            Serie = ""
    ''        Else
    ''            Serie = lector(23)
    ''        End If

    ''        If ((lector(24) Is DBNull.Value) OrElse (lector(24) Is Nothing)) Then
    ''            ID = ""
    ''        Else
    ''            ID = lector(24)
    ''        End If

    ''        If ((lector(25) Is DBNull.Value) OrElse (lector(25) Is Nothing)) Then
    ''            Accesorios = ""
    ''        Else
    ''            Accesorios = lector(25)
    ''        End If

    ''        If ((lector(26) Is DBNull.Value) OrElse (lector(26) Is Nothing)) Then
    ''            Puntos = ""
    ''        Else
    ''            Puntos = lector(26)
    ''        End If

    ''        If ((lector(30) Is DBNull.Value) OrElse (lector(30) Is Nothing)) Then
    ''            FuncionesCalibradas = 0
    ''        Else
    ''            FuncionesCalibradas = lector(30)
    ''        End If

    ''        If ((lector(31) Is DBNull.Value) OrElse (lector(31) Is Nothing)) Then
    ''            Etiquetas = ""
    ''        Else
    ''            Etiquetas = lector(31)
    ''        End If

    ''        If ((lector(33) Is DBNull.Value) OrElse (lector(33) Is Nothing)) Then
    ''            DatosdelInforme = ""
    ''        Else
    ''            DatosdelInforme = lector(33)
    ''        End If

    ''        If ((lector(34) Is DBNull.Value) OrElse (lector(34) Is Nothing)) Then
    ''            Observaciones = ""
    ''        Else
    ''            Observaciones = lector(34)
    ''        End If

    ''        If ((lector(35) Is DBNull.Value) OrElse (lector(35) Is Nothing)) Then
    ''            Calibro = ""
    ''        Else
    ''            Calibro = lector(35)
    ''        End If

    ''        If ((lector(36) Is DBNull.Value) OrElse (lector(36) Is Nothing)) Then
    ''            NumCot = 0
    ''        Else
    ''            NumCot = lector(36)
    ''        End If


    ''        If ((lector(38) Is DBNull.Value) OrElse (lector(38) Is Nothing)) Then
    ''            Status = ""
    ''        Else
    ''            Status = lector(38)
    ''        End If

    ''        If ((lector(40) Is DBNull.Value) OrElse (lector(40) Is Nothing)) Then
    ''            RealizoMedicion = ""
    ''        Else
    ''            RealizoMedicion = lector(40)
    ''        End If

    ''        If ((lector(41) Is DBNull.Value) OrElse (lector(41) Is Nothing)) Then
    ''            EmpresaEmision = ""
    ''        Else
    ''            EmpresaEmision = lector(41)
    ''        End If

    ''        If ((lector(42) Is DBNull.Value) OrElse (lector(42) Is Nothing)) Then
    ''            Calle = ""
    ''        Else
    ''            Calle = lector(42)
    ''        End If

    ''        If ((lector(43) Is DBNull.Value) OrElse (lector(43) Is Nothing)) Then
    ''            Colonia = ""
    ''        Else
    ''            Colonia = lector(43)
    ''        End If

    ''        If ((lector(44) Is DBNull.Value) OrElse (lector(44) Is Nothing)) Then
    ''            Ciudad = ""
    ''        Else
    ''            Ciudad = lector(44)
    ''        End If

    ''        If ((lector(45) Is DBNull.Value) OrElse (lector(45) Is Nothing)) Then
    ''            Estado = ""
    ''        Else
    ''            Estado = lector(45)
    ''        End If

    ''        If ((lector(46) Is DBNull.Value) OrElse (lector(46) Is Nothing)) Then
    ''            Pais = ""
    ''        Else
    ''            Pais = lector(46)
    ''        End If

    ''        If ((lector(47) Is DBNull.Value) OrElse (lector(47) Is Nothing)) Then
    ''            CP = ""
    ''        Else
    ''            CP = lector(47)
    ''        End If

    ''        ''---Diferente validación para los ajustes--
    ''        Dim calendarizacionTexto, PoliticaTexto, IntervaloTexto, EvaluacionTexto, CurvaTexto, MantenimientoTexto, IdiomaTexto As String
    ''        If ((lector(54) Is DBNull.Value) OrElse (lector(54) Is Nothing)) Then
    ''            Calendarizacion = 0
    ''        Else
    ''            Calendarizacion = lector(54)
    ''            If Calendarizacion = 21 Then
    ''                calendarizacionTexto = "Normal"
    ''            ElseIf Calendarizacion = 22 Then
    ''                calendarizacionTexto = "Programado"
    ''            ElseIf Calendarizacion = 23 Then
    ''                calendarizacionTexto = "Urgente"
    ''            End If
    ''        End If

    ''        If ((lector(48) Is DBNull.Value) OrElse (lector(48) Is Nothing)) Then
    ''            Politicadeajuste = 0
    ''        Else
    ''            Politicadeajuste = lector(48)
    ''            If Politicadeajuste = 1 Then
    ''                PoliticaTexto = "Sin Ajuste"
    ''            ElseIf Politicadeajuste = 2 Then
    ''                PoliticaTexto = "Ajuste 50%"
    ''            ElseIf Politicadeajuste = 3 Then
    ''                PoliticaTexto = "Ajuste 100%"
    ''            End If
    ''        End If

    ''        If ((lector(51) Is DBNull.Value) OrElse (lector(51) Is Nothing)) Then
    ''            Intervalodecalibracion = 0
    ''        Else
    ''            Intervalodecalibracion = lector(51)
    ''            If Intervalodecalibracion = 13 Then
    ''                IntervaloTexto = "Sin calculo de intervalo de caibración"
    ''            ElseIf Intervalodecalibracion = 14 Then
    ''                IntervaloTexto = "Calculo de intervalo de calibración"
    ''            End If
    ''        End If

    ''        If ((lector(50) Is DBNull.Value) OrElse (lector(50) Is Nothing)) Then
    ''            Evaluaciondelaconformidad = 0
    ''        Else
    ''            Evaluaciondelaconformidad = lector(50)
    ''            If Evaluaciondelaconformidad = 10 Then
    ''                EvaluacionTexto = "Sin evaluación de la conformidad"
    ''            ElseIf Evaluaciondelaconformidad = 11 Then
    ''                EvaluacionTexto = "Con evaluación de la conformidad de los resultados finales, incluyendo incertidumbres"
    ''            ElseIf Evaluaciondelaconformidad = 12 Then
    ''                EvaluacionTexto = "Con evaluación de la conformidad Eléctrica"
    ''            End If
    ''        End If

    ''        If ((lector(52) Is DBNull.Value) OrElse (lector(52) Is Nothing)) Then
    ''            Curvadeajuste = 0
    ''        Else
    ''            Curvadeajuste = lector(52)
    ''            If Curvadeajuste = 15 Then
    ''                CurvaTexto = "Sin curva de ajuste"
    ''            ElseIf Curvadeajuste = 16 Then
    ''                CurvaTexto = "Curva de ajuste con residuales y evaluación de incertidumbre"
    ''            ElseIf Curvadeajuste = 17 Then
    ''                CurvaTexto = "Curva de ajuste Eléctrica"
    ''            End If
    ''        End If

    ''        If ((lector(49) Is DBNull.Value) OrElse (lector(49) Is Nothing)) Then
    ''            Mantenimiento = 0
    ''        Else
    ''            Mantenimiento = lector(49)
    ''            If Mantenimiento = 4 Then
    ''                MantenimientoTexto = "Pintado, despintado, rotulado y ajuste de pesa paralelepípeda"
    ''            ElseIf Mantenimiento = 5 Then
    ''                MantenimientoTexto = "Sin mantenimiento"
    ''            ElseIf Mantenimiento = 6 Then
    ''                MantenimientoTexto = "Mantenimiento preventivo"
    ''            End If
    ''        End If

    ''        If ((lector(53) Is DBNull.Value) OrElse (lector(53) Is Nothing)) Then
    ''            Idioma = 0
    ''        Else
    ''            Idioma = lector(53)
    ''            If Idioma = 18 Then
    ''                IdiomaTexto = "Español"
    ''            ElseIf Idioma = 19 Then
    ''                IdiomaTexto = "Portada bilingüe español + inglés"
    ''            ElseIf Idioma = 20 Then
    ''                IdiomaTexto = "Certificado completo bilingüe español + inglés"
    ''            End If
    ''        End If

    ''        If ((lector(55) Is DBNull.Value) OrElse (lector(55) Is Nothing)) Then
    ''            SVAD10 = ""
    ''        Else
    ''            SVAD10 = lector(55)
    ''        End If

    ''        ''-------------------------------------------

    ''        If ((lector(56) Is DBNull.Value) OrElse (lector(56) Is Nothing)) Then
    ''            FechadeRegistro = ""
    ''        Else
    ''            FechadeRegistro = lector(56)
    ''        End If

    ''        If ((lector(60) Is DBNull.Value) OrElse (lector(60) Is Nothing)) Then
    ''            FechadeRecepcionLab = ""
    ''        Else
    ''            FechadeRecepcionLab = lector(60)
    ''        End If

    ''        If ((lector(27) Is DBNull.Value) OrElse (lector(27) Is Nothing)) Then
    ''            patron1 = ""
    ''        Else
    ''            patron1 = lector(27)
    ''        End If

    ''        If ((lector(28) Is DBNull.Value) OrElse (lector(28) Is Nothing)) Then
    ''            patron2 = ""
    ''        Else
    ''            patron2 = lector(28)
    ''        End If

    ''        If ((lector(29) Is DBNull.Value) OrElse (lector(29) Is Nothing)) Then
    ''            patron3 = ""
    ''        Else
    ''            patron3 = lector(29)
    ''        End If

    ''        If ((lector(61) Is DBNull.Value) OrElse (lector(61) Is Nothing)) Then
    ''            patron4 = ""
    ''        Else
    ''            patron4 = lector(61)
    ''        End If

    ''        If ((lector(62) Is DBNull.Value) OrElse (lector(62) Is Nothing)) Then
    ''            patron5 = ""
    ''        Else
    ''            patron5 = lector(62)
    ''        End If

    ''        If ((lector(63) Is DBNull.Value) OrElse (lector(63) Is Nothing)) Then
    ''            patron6 = lector(63)
    ''        Else
    ''            patron6 = lector(63)
    ''        End If

    ''        If ((lector(64) Is DBNull.Value) OrElse (lector(64) Is Nothing)) Then
    ''            patron7 = ""
    ''        Else
    ''            patron7 = lector(64)
    ''        End If

    ''        If ((lector(65) Is DBNull.Value) OrElse (lector(65) Is Nothing)) Then
    ''            patron8 = ""
    ''        Else
    ''            patron8 = lector(65)
    ''        End If

    ''        If ((lector(66) Is DBNull.Value) OrElse (lector(66) Is Nothing)) Then
    ''            patron9 = ""
    ''        Else
    ''            patron9 = lector(66)
    ''        End If

    ''        If ((lector(67) Is DBNull.Value) OrElse (lector(67) Is Nothing)) Then
    ''            patron10 = ""
    ''        Else
    ''            patron10 = lector(67)
    ''        End If

    ''        If ((lector(68) Is DBNull.Value) OrElse (lector(68) Is Nothing)) Then
    ''            firma = ""
    ''        Else
    ''            firma = lector(68)
    ''        End If

    ''        If ((lector(69) Is DBNull.Value) OrElse (lector(69) Is Nothing)) Then
    ''            observacionStatus = ""
    ''        Else
    ''            observacionStatus = lector(69)
    ''        End If

    ''        ''---------------------------------------------------------
    ''        'Dim Aplicacion As New Excel.Application
    ''        'Dim Libro As Excel.Workbook
    ''        'Dim Hoja As Excel.Worksheet

    ''        'Aplicacion = New Excel.Application
    ''        'Libro = Aplicacion.Workbooks.Open(RutaArchivo)
    ''        'Hoja = Libro.Worksheets("Importados")

    ''        ''Aquí manipulen su archivo
    ''        'Hoja.Cells(3, 3) = "Hola"

    ''        'Libro.Close()
    ''        'Aplicacion.Quit()

    ''        'releaseObject(Aplicacion)
    ''        'releaseObject(Libro)
    ''        'releaseObject(Hoja)
    ''        ''---------------------------------------------------------
    ''    Catch ex As Exception
    ''        MsgBox("Ocurrio un error de lectura de datos, verifica la entrada de información.", MsgBoxStyle.Critical)
    ''    End Try
    ''End Sub
End Class