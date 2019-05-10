Imports System.Data.SqlClient
Imports Microsoft.Reporting.WinForms
Imports System.Configuration
Public Class ImprmirInforme
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim INF, MAG As String
        INF = txtInforme.Text
        MAG = txtMagnitud.Text
        MetodoMetasInf2019()
        comando2019 = conexion2019.CreateCommand
        Dim R As String
        R = "SELECT Folio,TIPO,Empresa, ServCatalogo1, ClavecontactoConsign,[Claves-Elaboro-Cot].[Nombre],Sv1Ajuste,Sv3Matto,
                Sv5COM02,Sv6IntervaloRe,Sv7Curva,Sv8Idioma,Sv9Calendar,SVAD10,PUNTOS, Observaciones
                FROM [INFORMES-SERVICIOS] 
                INNER JOIN [Claves-Elaboro-Cot] ON [INFORMES-SERVICIOS].CveOperador = [Claves-Elaboro-Cot].[Clave-elaboro-cot] 
            WHERE MAGNITUD = '" & txtMagnitud.Text & "' AND INFORME = '" & txtInforme.Text & "'"
        comando2019.CommandText = R
        lector2019 = comando2019.ExecuteReader
        'MsgBox(R)
        lector2019.Read()
        Dim fol, cveContCons As Integer
        Dim tipo, emp, serv, elaboroCot, ajuste, mantto, com, inter, curva, idioma, calen, vad, pun, obser As String
        If ((lector2019(0) Is DBNull.Value) OrElse (lector2019(0) Is Nothing)) Then
            fol = "-"
        Else
            fol = lector2019(0)
        End If
        If ((lector2019(1) Is DBNull.Value) OrElse (lector2019(1) Is Nothing)) Then
            tipo = "-"
        Else
            tipo = lector2019(1)
        End If
        If ((lector2019(2) Is DBNull.Value) OrElse (lector2019(2) Is Nothing)) Then
            emp = "-"
        Else
            emp = lector2019(2)
        End If
        If ((lector2019(3) Is DBNull.Value) OrElse (lector2019(3) Is Nothing)) Then
            serv = "-"
        Else
            serv = lector2019(3)
        End If
        If ((lector2019(4) Is DBNull.Value) OrElse (lector2019(4) Is Nothing)) Then
            cveContCons = "-"
        Else
            cveContCons = lector2019(4)
        End If
        If ((lector2019(5) Is DBNull.Value) OrElse (lector2019(5) Is Nothing)) Then
            elaboroCot = "-"
        Else
            elaboroCot = lector2019(5)
        End If
        If ((lector2019(6) Is DBNull.Value) OrElse (lector2019(6) Is Nothing)) Then
            ajuste = "-"
        Else
            ajuste = lector2019(6)
            Select Case ajuste
                Case 1
                    ajuste = "Sin ajuste"
                Case 2
                    ajuste = "Ajuste 50%"
                Case 3
                    ajuste = "Ajuste 100%"
                Case Else
                    ajuste = "Pintado, despintado, rotulado y ajuste de pesa paralelepípeda"
            End Select
        End If
        If ((lector2019(7) Is DBNull.Value) OrElse (lector2019(7) Is Nothing)) Then
            mantto = "-"
        Else
            mantto = lector2019(7)
            Select Case mantto
                Case 5
                    mantto = "Sin mantenimiento"
                Case 6
                    mantto = "Mantenimiento preventivo"
            End Select
        End If
        If ((lector2019(8) Is DBNull.Value) OrElse (lector2019(8) Is Nothing)) Then
            com = "-"
        Else
            com = lector2019(8)
            Select Case com
                Case 10
                    com = "Sin evaluación de la conformidad"
                Case 11
                    com = "Con evaluación de la conformidad de los resultados finales, incluyendo incertidumbres"
                Case 12
                    com = "Con evaluación de la conformidad Eléctrica"
            End Select
        End If
        If ((lector2019(9) Is DBNull.Value) OrElse (lector2019(9) Is Nothing)) Then
            inter = "-"
        Else
            inter = lector2019(9)
            Select Case inter
                Case 13
                    inter = "Sin calculo de intervalo de caibración"
                Case 14
                    inter = "Calculo de intervalo de calibración"
            End Select
        End If
        If ((lector2019(10) Is DBNull.Value) OrElse (lector2019(10) Is Nothing)) Then
            curva = "-"
        Else
            curva = lector2019(10)
            Select Case curva
                Case 15
                    curva = "Sin curva de ajuste"
                Case 16
                    curva = "Curva de ajuste con residuales y evaluación de incertidumbre"
                Case 17
                    curva = "Curva de ajuste Eléctrica"
            End Select
        End If
        If ((lector2019(11) Is DBNull.Value) OrElse (lector2019(11) Is Nothing)) Then
            idioma = "-"
        Else
            idioma = lector2019(11)
            Select Case idioma
                Case 18
                    idioma = "Normal"
                Case 19
                    idioma = "Programado"
                Case 20
                    idioma = "Urgente"
            End Select
        End If
        If ((lector2019(12) Is DBNull.Value) OrElse (lector2019(12) Is Nothing)) Then
            calen = "-"
        Else
            calen = lector2019(12)
            Select Case calen
                Case 21
                    calen = "Español"
                Case 22
                    calen = "Portada bilingüe español + inglés"
                Case 23
                    calen = "Certificado completo bilingüe español + inglés"
            End Select
        End If
        If ((lector2019(13) Is DBNull.Value) OrElse (lector2019(13) Is Nothing)) Then
            vad = "-"
        Else
            vad = lector2019(13)
        End If
        If ((lector2019(14) Is DBNull.Value) OrElse (lector2019(14) Is Nothing)) Then
            pun = "-"
        Else
            pun = lector2019(14)
        End If
        If ((lector2019(15) Is DBNull.Value) OrElse (lector2019(15) Is Nothing)) Then
            obser = "-"
        Else
            obser = lector2019(15)
        End If
        lector2019.Close()
        Dim Adaptador As New SqlDataAdapter
        Adaptador.SelectCommand = New SqlCommand
        Adaptador.SelectCommand.Connection = conexion2019
        Adaptador.SelectCommand.CommandText = "Informe"
        Adaptador.SelectCommand.CommandType = CommandType.StoredProcedure
        Dim param0 = New SqlParameter("@FOLIO", SqlDbType.VarChar)
        Dim param1 = New SqlParameter("@MAGNITUD", SqlDbType.VarChar)
        Dim param2 = New SqlParameter("@INFORME", SqlDbType.VarChar)
        Dim param3 = New SqlParameter("@TIPO", SqlDbType.VarChar)
        Dim param4 = New SqlParameter("@EMPRESA", SqlDbType.VarChar)
        Dim param5 = New SqlParameter("@CATALOGO", SqlDbType.VarChar)
        Dim param6 = New SqlParameter("@CVEMPRESA", SqlDbType.VarChar)
        Dim param7 = New SqlParameter("@OPERADOR", SqlDbType.VarChar)
        Dim param8 = New SqlParameter("@AJUSTE", SqlDbType.VarChar)
        Dim param9 = New SqlParameter("@MATTO", SqlDbType.VarChar)
        Dim param10 = New SqlParameter("@COM", SqlDbType.VarChar)
        Dim param11 = New SqlParameter("@INTERVALO", SqlDbType.VarChar)
        Dim param12 = New SqlParameter("@CURVA", SqlDbType.VarChar)
        Dim param13 = New SqlParameter("@IDIOMA", SqlDbType.VarChar)
        Dim param14 = New SqlParameter("@CALENDAR", SqlDbType.VarChar)
        Dim param15 = New SqlParameter("@VAD", SqlDbType.VarChar)
        Dim param16 = New SqlParameter("@PUNTOS", SqlDbType.VarChar)
        Dim param17 = New SqlParameter("@OBSER", SqlDbType.VarChar)
        param0.Direction = ParameterDirection.Input
        param1.Direction = ParameterDirection.Input
        param2.Direction = ParameterDirection.Input
        param3.Direction = ParameterDirection.Input
        param4.Direction = ParameterDirection.Input
        param5.Direction = ParameterDirection.Input
        param6.Direction = ParameterDirection.Input
        param7.Direction = ParameterDirection.Input
        param8.Direction = ParameterDirection.Input
        param9.Direction = ParameterDirection.Input
        param10.Direction = ParameterDirection.Input
        param11.Direction = ParameterDirection.Input
        param12.Direction = ParameterDirection.Input
        param13.Direction = ParameterDirection.Input
        param14.Direction = ParameterDirection.Input
        param15.Direction = ParameterDirection.Input
        param16.Direction = ParameterDirection.Input
        param17.Direction = ParameterDirection.Input
        param0.Value = fol
        param1.Value = MAG
        param2.Value = INF
        param3.Value = tipo
        param4.Value = emp
        param5.Value = serv
        param6.Value = cveContCons
        param7.Value = elaboroCot
        param8.Value = ajuste
        param9.Value = mantto
        param10.Value = com
        param11.Value = inter
        param12.Value = curva
        param13.Value = idioma
        param14.Value = calen
        param15.Value = vad
        param16.Value = pun
        param17.Value = obser
        Adaptador.SelectCommand.Parameters.Add(param0)
        Adaptador.SelectCommand.Parameters.Add(param1)
        Adaptador.SelectCommand.Parameters.Add(param2)
        Adaptador.SelectCommand.Parameters.Add(param3)
        Adaptador.SelectCommand.Parameters.Add(param4)
        Adaptador.SelectCommand.Parameters.Add(param5)
        Adaptador.SelectCommand.Parameters.Add(param6)
        Adaptador.SelectCommand.Parameters.Add(param7)
        Adaptador.SelectCommand.Parameters.Add(param8)
        Adaptador.SelectCommand.Parameters.Add(param9)
        Adaptador.SelectCommand.Parameters.Add(param10)
        Adaptador.SelectCommand.Parameters.Add(param11)
        Adaptador.SelectCommand.Parameters.Add(param12)
        Adaptador.SelectCommand.Parameters.Add(param13)
        Adaptador.SelectCommand.Parameters.Add(param14)
        Adaptador.SelectCommand.Parameters.Add(param15)
        Adaptador.SelectCommand.Parameters.Add(param16)
        Adaptador.SelectCommand.Parameters.Add(param17)
        Dim Data As New DataSet
        Adaptador.Fill(Data)
        Data.DataSetName = "Data1"
        Dim Datasource As New ReportDataSource("DataSet1", Data.Tables(0))
        Datasource.Name = "DataSet1"
        Datasource.Value = Data.Tables(0)
        Dim p0 As New ReportParameter("FOLIO", fol)
        Dim p1 As New ReportParameter("MAGNITUD", MAG)
        Dim p2 As New ReportParameter("INFORME", INF)
        Dim p3 As New ReportParameter("TIPO", tipo)
        Dim p4 As New ReportParameter("EMPRESA", emp)
        Dim p5 As New ReportParameter("CATALOGO", serv)
        Dim p6 As New ReportParameter("CVEMPRESA", cveContCons)
        Dim p7 As New ReportParameter("OPERADOR", elaboroCot)
        Dim p8 As New ReportParameter("AJUSTE", ajuste)
        Dim p9 As New ReportParameter("MTTO", mantto)
        Dim p10 As New ReportParameter("COM", com)
        Dim p11 As New ReportParameter("INTERVALO", inter)
        Dim p12 As New ReportParameter("CURVA", curva)
        Dim p13 As New ReportParameter("IDIOMA", idioma)
        Dim p14 As New ReportParameter("CALENDAR", calen)
        Dim p15 As New ReportParameter("VAD", vad)
        Dim p16 As New ReportParameter("PUNTOS", pun)
        Dim p17 As New ReportParameter("OBSERV", obser)
        Dim Reportes As New ReportDataSource("DataSet1", Data.Tables(0))
        FrmReportes.ReportViewer1.LocalReport.DataSources.Clear()
        FrmReportes.ReportViewer1.LocalReport.DataSources.Add(Datasource)
        FrmReportes.ReportViewer1.LocalReport.ReportPath = "C:\Users\Software TI\Documents\GitHub\ServiciosExcel\VinculacionDePestanas\Reportes\Report1.rdlc"
        FrmReportes.ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p0, p1, p2, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17})
        FrmReportes.ReportViewer1.RefreshReport()
        FrmReportes.Show()
        conexion2019.Close()
    End Sub
End Class
