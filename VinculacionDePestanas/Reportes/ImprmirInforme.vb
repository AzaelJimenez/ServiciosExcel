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
        R = "SELECT Folio,TIPO,Empresa, ServCatalogo1, ClavecontactoConsign,[Claves-Elaboro-Cot].[Clave-elaboro-cot],Sv1Ajuste,Sv3Matto,
            Sv5COM02,Sv6IntervaloRe,Sv7Curva,Sv8Idioma,Sv9Calendar,SVAD10,PUNTOS, Observaciones,EmpresaEmision, DirCalleEmision,
            DirCiudadEmision, DirEdoProvEmision,DirCPEmision, DirPaisEmision,[Contactos-Clientes-Usuarios].Email,
            [Contactos-Clientes-Usuarios].Tel,[INFORMES-SERVICIOS].Usuario,FECHARECEP,[Fecha_Reg],[FECHACALIB],[FECHA-EMISION],[FechaRecepLab],
            MARCA,MODELO,Serie,NumFuncionesCalibradas,[Patron1],[Patron2],[Patron3],[Patron4],[Patron5],[Patron6],[Patron7],[Patron8],[Patron9],
            [Patron10],[Status],[ObservacionStatus],ID,ALCANCE
            FROM [INFORMES-SERVICIOS] 
            INNER JOIN [Claves-Elaboro-Cot] ON [INFORMES-SERVICIOS].CveOperador = [Claves-Elaboro-Cot].[Clave-elaboro-cot] 
            INNER JOIN [InformacionGeneral].[dbo].[Contactos-Clientes-Usuarios] ON [INFORMES-SERVICIOS].Clavecontacto = [InformacionGeneral].[dbo].[Contactos-Clientes-Usuarios].ClaveContacto
            WHERE MAGNITUD = '" & txtMagnitud.Text & "' AND INFORME = '" & txtInforme.Text & "'"
        comando2019.CommandText = R
        lector2019 = comando2019.ExecuteReader
        'MsgBox(R)
        lector2019.Read()
        Dim fol, cveContCons As Integer
        Dim tipo, emp, serv, elaboroCot, ajuste, mantto, com, inter, curva, idioma, calen, vad, pun, obser,
         empemi, dir, cd, edo, cp, pais, email, user, marca, modelo, serie, nfun,
        pa1, pa2, pa3, pa4, pa5, pa6, pa7, pa8, pa9, pa10, status, obsersta, tel, ID, alcance, a As String
        Dim frecep, freg, frecal, femi, freceplab As Date
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
            Select Case elaboroCot
                Case 10
                    elaboroCot = "Natalia García"
                Case 8
                    elaboroCot = "Tonaxy Reyes"
                Case 17
                    elaboroCot = "Pedro Palacios"
                Case 18
                    elaboroCot = "Fernando González"
                Case 21
                    elaboroCot = "Marlene Peña"
                Case 22
                    elaboroCot = "Carlos Ramirez"
                Case 4
                    elaboroCot = "Ximena Mena"
                Case 14
                    elaboroCot = "José Larios"
                Case 6
                    elaboroCot = "Erika Berduzco"
                Case 5
                    elaboroCot = "Mayra Ramírez"
                Case Else
                    elaboroCot = "Gestor de Servicio"
            End Select
        End If
        If ((lector2019(6) Is DBNull.Value) OrElse (lector2019(6) Is Nothing)) Then
            ajuste = "-"
        Else
            ajuste = lector2019(6)
            Select Case ajuste
                Case 1
                    ajuste = "Sin ajuste"
                    a = "SI FUNCIONA"
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
                    idioma = "Español"
                Case 19
                    idioma = "Portada bilingüe español + inglés"
                Case 20
                    idioma = "Certificado completo bilingüe español + inglés"
            End Select
        End If
        If ((lector2019(12) Is DBNull.Value) OrElse (lector2019(12) Is Nothing)) Then
            calen = "-"
        Else
            calen = lector2019(12)
            Select Case calen
                Case 21
                    calen = "Normal"
                Case 22
                    calen = "Programado"
                Case 23
                    calen = "Urgente"
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
        If ((lector2019(16) Is DBNull.Value) OrElse (lector2019(16) Is Nothing)) Then
            empemi = "-"
        Else
            empemi = lector2019(16)
        End If
        If ((lector2019(17) Is DBNull.Value) OrElse (lector2019(17) Is Nothing)) Then
            dir = "-"
        Else
            dir = lector2019(17)
        End If
        If ((lector2019(18) Is DBNull.Value) OrElse (lector2019(18) Is Nothing)) Then
            cd = "-"
        Else
            cd = lector2019(18)
        End If
        If ((lector2019(19) Is DBNull.Value) OrElse (lector2019(19) Is Nothing)) Then
            edo = "-"
        Else
            edo = lector2019(19)
        End If
        If ((lector2019(20) Is DBNull.Value) OrElse (lector2019(20) Is Nothing)) Then
            cp = "-"
        Else
            cp = lector2019(20)
        End If
        If ((lector2019(21) Is DBNull.Value) OrElse (lector2019(21) Is Nothing)) Then
            pais = "-"
        Else
            pais = lector2019(21)
        End If
        If ((lector2019(22) Is DBNull.Value) OrElse (lector2019(22) Is Nothing)) Then
            email = "-"
        Else
            email = lector2019(22)
        End If
        If ((lector2019(23) Is DBNull.Value) OrElse (lector2019(23) Is Nothing)) Then
            tel = "-"
        Else
            tel = lector2019(23)
        End If
        If ((lector2019(24) Is DBNull.Value) OrElse (lector2019(24) Is Nothing)) Then
            user = "-"
        Else
            user = lector2019(24)
        End If
        If ((lector2019(25) Is DBNull.Value) OrElse (lector2019(25) Is Nothing)) Then
            frecep = "1999-01-01"
        Else
            frecep = lector2019(25)
        End If
        If ((lector2019(26) Is DBNull.Value) OrElse (lector2019(26) Is Nothing)) Then
            freg = "1999-01-01"
        Else
            freg = lector2019(26)
        End If
        If ((lector2019(27) Is DBNull.Value) OrElse (lector2019(27) Is Nothing)) Then
            frecal = "1999-01-01"
        Else
            frecal = lector2019(27)
        End If
        If ((lector2019(28) Is DBNull.Value) OrElse (lector2019(28) Is Nothing)) Then
            femi = "1999-01-01"
        Else
            femi = lector2019(28)
        End If
        If ((lector2019(29) Is DBNull.Value) OrElse (lector2019(29) Is Nothing)) Then
            freceplab = "1999-01-01"
        Else
            freceplab = lector2019(29)
        End If
        If ((lector2019(30) Is DBNull.Value) OrElse (lector2019(30) Is Nothing)) Then
            marca = " "
        Else
            marca = lector2019(30)
        End If
        If ((lector2019(31) Is DBNull.Value) OrElse (lector2019(31) Is Nothing)) Then
            modelo = " "
        Else
            modelo = lector2019(31)
        End If
        If ((lector2019(32) Is DBNull.Value) OrElse (lector2019(32) Is Nothing)) Then
            serie = " "
        Else
            serie = lector2019(32)
        End If
        If ((lector2019(33) Is DBNull.Value) OrElse (lector2019(33) Is Nothing)) Then
            nfun = " "
        Else
            nfun = lector2019(33)
        End If
        If ((lector2019(34) Is DBNull.Value) OrElse (lector2019(34) Is Nothing)) Then
            pa1 = " "
        Else
            pa1 = lector2019(34)
        End If
        If ((lector2019(35) Is DBNull.Value) OrElse (lector2019(35) Is Nothing)) Then
            pa2 = " "
        Else
            pa2 = lector2019(35)
        End If
        If ((lector2019(36) Is DBNull.Value) OrElse (lector2019(36) Is Nothing)) Then
            pa3 = " "
        Else
            pa3 = lector2019(36)
        End If
        If ((lector2019(37) Is DBNull.Value) OrElse (lector2019(37) Is Nothing)) Then
            pa4 = " "
        Else
            pa4 = lector2019(37)
        End If
        If ((lector2019(38) Is DBNull.Value) OrElse (lector2019(38) Is Nothing)) Then
            pa5 = " "
        Else
            pa5 = lector2019(38)
        End If
        If ((lector2019(39) Is DBNull.Value) OrElse (lector2019(39) Is Nothing)) Then
            pa6 = " "
        Else
            pa6 = lector2019(39)
        End If
        If ((lector2019(40) Is DBNull.Value) OrElse (lector2019(40) Is Nothing)) Then
            pa7 = " "
        Else
            pa7 = lector2019(40)
        End If
        If ((lector2019(41) Is DBNull.Value) OrElse (lector2019(41) Is Nothing)) Then
            pa8 = " "
        Else
            pa8 = lector2019(41)
        End If
        If ((lector2019(42) Is DBNull.Value) OrElse (lector2019(42) Is Nothing)) Then
            pa9 = " "
        Else
            pa9 = lector2019(42)
        End If
        If ((lector2019(43) Is DBNull.Value) OrElse (lector2019(43) Is Nothing)) Then
            pa10 = " "
        Else
            pa10 = lector2019(43)
        End If
        If ((lector2019(44) Is DBNull.Value) OrElse (lector2019(44) Is Nothing)) Then
            status = " "
        Else
            status = lector2019(44)
        End If
        If ((lector2019(45) Is DBNull.Value) OrElse (lector2019(45) Is Nothing)) Then
            obsersta = " "
        Else
            obsersta = lector2019(45)
        End If
        If ((lector2019(46) Is DBNull.Value) OrElse (lector2019(46) Is Nothing)) Then
            ID = " "
        Else
            ID = lector2019(46)
        End If
        If ((lector2019(47) Is DBNull.Value) OrElse (lector2019(47) Is Nothing)) Then
            alcance = " "
        Else
            alcance = lector2019(47)
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
        Dim param18 = New SqlParameter("@EMP", SqlDbType.VarChar)
        Dim param19 = New SqlParameter("@DIR", SqlDbType.VarChar)
        Dim param20 = New SqlParameter("@CD", SqlDbType.VarChar)
        Dim param21 = New SqlParameter("@EDO", SqlDbType.VarChar)
        Dim param22 = New SqlParameter("@CP", SqlDbType.VarChar)
        Dim param23 = New SqlParameter("@PAIS", SqlDbType.VarChar)
        Dim param24 = New SqlParameter("@EMAIL", SqlDbType.VarChar)
        Dim param25 = New SqlParameter("@TEL", SqlDbType.VarChar)
        Dim param26 = New SqlParameter("@USER", SqlDbType.VarChar)
        Dim param27 = New SqlParameter("@FRECEP", SqlDbType.DateTime)
        Dim param28 = New SqlParameter("@FREG", SqlDbType.DateTime)
        Dim param29 = New SqlParameter("@FRECAL", SqlDbType.DateTime)
        Dim param30 = New SqlParameter("@FEMI", SqlDbType.DateTime)
        Dim param31 = New SqlParameter("@FRECEPLAB", SqlDbType.DateTime)
        Dim param32 = New SqlParameter("@MARCA", SqlDbType.VarChar)
        Dim param33 = New SqlParameter("@MODELO", SqlDbType.VarChar)
        Dim param34 = New SqlParameter("@SERIE", SqlDbType.VarChar)
        Dim param35 = New SqlParameter("@NFUN", SqlDbType.VarChar)
        Dim param36 = New SqlParameter("@PA1", SqlDbType.VarChar)
        Dim param37 = New SqlParameter("@PA2", SqlDbType.VarChar)
        Dim param38 = New SqlParameter("@PA3", SqlDbType.VarChar)
        Dim param39 = New SqlParameter("@PA4", SqlDbType.VarChar)
        Dim param40 = New SqlParameter("@PA5", SqlDbType.VarChar)
        Dim param41 = New SqlParameter("@PA6", SqlDbType.VarChar)
        Dim param42 = New SqlParameter("@PA7", SqlDbType.VarChar)
        Dim param43 = New SqlParameter("@PA8", SqlDbType.VarChar)
        Dim param44 = New SqlParameter("@PA9", SqlDbType.VarChar)
        Dim param45 = New SqlParameter("@PA10", SqlDbType.VarChar)
        Dim param46 = New SqlParameter("@STATUS", SqlDbType.VarChar)
        Dim param47 = New SqlParameter("@OBSERSTA", SqlDbType.VarChar)
        Dim param48 = New SqlParameter("@ID", SqlDbType.VarChar)
        Dim param49 = New SqlParameter("@INTER", SqlDbType.VarChar)
        Dim param50 = New SqlParameter("@a", SqlDbType.VarChar)
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
        param18.Direction = ParameterDirection.Input
        param19.Direction = ParameterDirection.Input
        param20.Direction = ParameterDirection.Input
        param21.Direction = ParameterDirection.Input
        param22.Direction = ParameterDirection.Input
        param23.Direction = ParameterDirection.Input
        param24.Direction = ParameterDirection.Input
        param25.Direction = ParameterDirection.Input
        param26.Direction = ParameterDirection.Input
        param27.Direction = ParameterDirection.Input
        param28.Direction = ParameterDirection.Input
        param29.Direction = ParameterDirection.Input
        param30.Direction = ParameterDirection.Input
        param31.Direction = ParameterDirection.Input
        param32.Direction = ParameterDirection.Input
        param33.Direction = ParameterDirection.Input
        param34.Direction = ParameterDirection.Input
        param35.Direction = ParameterDirection.Input
        param36.Direction = ParameterDirection.Input
        param37.Direction = ParameterDirection.Input
        param38.Direction = ParameterDirection.Input
        param38.Direction = ParameterDirection.Input
        param40.Direction = ParameterDirection.Input
        param41.Direction = ParameterDirection.Input
        param42.Direction = ParameterDirection.Input
        param43.Direction = ParameterDirection.Input
        param44.Direction = ParameterDirection.Input
        param45.Direction = ParameterDirection.Input
        param46.Direction = ParameterDirection.Input
        param47.Direction = ParameterDirection.Input
        param48.Direction = ParameterDirection.Input
        param49.Direction = ParameterDirection.Input
        param50.Direction = ParameterDirection.Input
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
        param18.Value = empemi
        param19.Value = dir
        param20.Value = cd
        param21.Value = edo
        param22.Value = cp
        param23.Value = pais
        param24.Value = email
        param25.Value = tel
        param26.Value = user
        param27.Value = frecep
        param28.Value = freg
        param29.Value = frecal
        param30.Value = femi
        param31.Value = freceplab
        param32.Value = marca
        param33.Value = modelo
        param34.Value = serie
        param35.Value = nfun
        param36.Value = pa1
        param37.Value = pa2
        param38.Value = pa3
        param39.Value = pa4
        param40.Value = pa5
        param41.Value = pa6
        param42.Value = pa7
        param43.Value = pa8
        param44.Value = pa9
        param45.Value = pa10
        param46.Value = status
        param47.Value = obsersta
        param48.Value = ID
        param49.Value = alcance
        param50.Value = a
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
        Adaptador.SelectCommand.Parameters.Add(param18)
        Adaptador.SelectCommand.Parameters.Add(param19)
        Adaptador.SelectCommand.Parameters.Add(param20)
        Adaptador.SelectCommand.Parameters.Add(param21)
        Adaptador.SelectCommand.Parameters.Add(param22)
        Adaptador.SelectCommand.Parameters.Add(param23)
        Adaptador.SelectCommand.Parameters.Add(param24)
        Adaptador.SelectCommand.Parameters.Add(param25)
        Adaptador.SelectCommand.Parameters.Add(param26)
        Adaptador.SelectCommand.Parameters.Add(param27)
        Adaptador.SelectCommand.Parameters.Add(param28)
        Adaptador.SelectCommand.Parameters.Add(param29)
        Adaptador.SelectCommand.Parameters.Add(param30)
        Adaptador.SelectCommand.Parameters.Add(param31)
        Adaptador.SelectCommand.Parameters.Add(param32)
        Adaptador.SelectCommand.Parameters.Add(param33)
        Adaptador.SelectCommand.Parameters.Add(param34)
        Adaptador.SelectCommand.Parameters.Add(param35)
        Adaptador.SelectCommand.Parameters.Add(param36)
        Adaptador.SelectCommand.Parameters.Add(param37)
        Adaptador.SelectCommand.Parameters.Add(param38)
        Adaptador.SelectCommand.Parameters.Add(param39)
        Adaptador.SelectCommand.Parameters.Add(param40)
        Adaptador.SelectCommand.Parameters.Add(param41)
        Adaptador.SelectCommand.Parameters.Add(param42)
        Adaptador.SelectCommand.Parameters.Add(param43)
        Adaptador.SelectCommand.Parameters.Add(param44)
        Adaptador.SelectCommand.Parameters.Add(param45)
        Adaptador.SelectCommand.Parameters.Add(param46)
        Adaptador.SelectCommand.Parameters.Add(param47)
        Adaptador.SelectCommand.Parameters.Add(param48)
        Adaptador.SelectCommand.Parameters.Add(param49)
        Adaptador.SelectCommand.Parameters.Add(param50)
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
        Dim p18 As New ReportParameter("EMP", empemi)
        Dim p19 As New ReportParameter("DIR", dir)
        Dim p20 As New ReportParameter("CD", cd)
        Dim p21 As New ReportParameter("EDO", edo)
        Dim p22 As New ReportParameter("CP", cp)
        Dim p23 As New ReportParameter("PAIS", pais)
        Dim p24 As New ReportParameter("EMAIL", email)
        Dim p25 As New ReportParameter("TEL", tel)
        Dim p26 As New ReportParameter("USER", user)
        Dim p27 As New ReportParameter("FRECEP", frecep)
        Dim p28 As New ReportParameter("FREG", freg)
        Dim p29 As New ReportParameter("FRECAL", frecal)
        Dim p30 As New ReportParameter("FEMI", femi)
        Dim p31 As New ReportParameter("FRECEPLAB", freceplab)
        Dim p32 As New ReportParameter("MARCA", marca)
        Dim p33 As New ReportParameter("MODELO", modelo)
        Dim p34 As New ReportParameter("SERIE", serie)
        Dim p35 As New ReportParameter("NFUN", nfun)
        Dim p36 As New ReportParameter("PA1", pa1)
        Dim p37 As New ReportParameter("PA2", pa2)
        Dim p38 As New ReportParameter("PA3", pa3)
        Dim p39 As New ReportParameter("PA4", pa4)
        Dim p40 As New ReportParameter("PA5", pa5)
        Dim p41 As New ReportParameter("PA6", pa6)
        Dim p42 As New ReportParameter("PA7", pa7)
        Dim p43 As New ReportParameter("PA8", pa8)
        Dim p44 As New ReportParameter("PA9", pa9)
        Dim p45 As New ReportParameter("PA10", pa10)
        Dim p46 As New ReportParameter("STATUS", status)
        Dim p47 As New ReportParameter("OBSERSTA", obsersta)
        Dim p48 As New ReportParameter("ID", ID)
        Dim p49 As New ReportParameter("INTER", alcance)
        Dim p50 As New ReportParameter("a", a)
        Dim Reportes As New ReportDataSource("DataSet1", Data.Tables(0))
        FrmReportes.ReportViewer1.LocalReport.DataSources.Clear()
        FrmReportes.ReportViewer1.LocalReport.DataSources.Add(Datasource)
        FrmReportes.ReportViewer1.LocalReport.ReportPath = "C:\Users\Software TI\Documents\GitHub\ServiciosExcel\VinculacionDePestanas\Reportes\Report1.rdlc"
        FrmReportes.ReportViewer1.LocalReport.SetParameters(New ReportParameter() {p0, p1, p2, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17,
                                                            p18, p19, p20, p21, p22, p23, p24, p25, p26, p27, p28, p29, p30, p31, p32, p33, p34, p35, p36, p37, p38, p39,
                                                            p40, p41, p42, p43, p44, p45, p46, p47, p48, p49, p50})
        FrmReportes.ReportViewer1.RefreshReport()
        FrmReportes.Show()
        conexion2019.Close()
    End Sub
End Class
