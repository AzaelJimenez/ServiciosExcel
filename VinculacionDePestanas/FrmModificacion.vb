Imports System.Data.SqlClient
Public Class FrmModificacion
    Dim Operador As String
    Private Sub FrmModificacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ''Try
        ''    Dim R As String
        ''    MetodoMetasInf2019()
        ''    ' conexion2019.Open()
        ''    R = "SELECT isnull([Folio],'-') as Folio 
        ''        ,isnull([ClavecontactoConsign],'-')as claveempresa 
        ''        ,isnull([Empresa],'-')as empresa 
        ''        ,[FECHARECEP]as fechaRecep 
        ''        ,[FECHACALIB]as fechaCalib 
        ''        ,[FECHA-EMISION]as FechaEmision 
        ''        ,isnull([ServCatalogo1],'-')as Catalogo1 
        ''        ,isnull([ServCatalogo2],'-')as catalogo2 
        ''        ,isnull([ServiciosAdicionales],'-')as adicionales 
        ''        ,isnull([TIPO],'-')as Tipo 
        ''        ,isnull([ALCANCE],'-')as alcance 
        ''        ,isnull([MARCA],'-')as marca 
        ''        ,isnull([MODELO],'-')as modelo 
        ''        ,isnull([Serie],'-')as serie 
        ''        ,isnull([ID],'-')as id 
        ''        ,isnull([Patron1],'-')as patron1 
        ''        ,isnull([Patron2],'-')as patron2
        ''        ,isnull([Patron3],'-')as patron3 
        ''        ,isnull([CALIBRO],'-')as calibro 
        ''        ,isnull([Status],'-')as status 
        ''        ,[Fecha_Reg]as fechaRegistro 
        ''        ,[FechaRecepLab]as fechaReceplab 
        ''        ,isnull([Patron4],'-')as patron4 
        ''        ,isnull([Patron5],'-')as patron5 
        ''        ,isnull([Patron6],'-')as patron6 
        ''        ,isnull([Patron7],'-')as patron7 
        ''        ,isnull([Patron8],'-')as patron8 
        ''        ,isnull([Patron9],'-')as patron9 
        ''        ,isnull([Patron10],'-')as patron10 
        ''        ,isnull([Firma],'-')as Firma 
        ''      FROM [INFORMES-SERVICIOS] where [MAGNITUD]='" & txtMagnitud.Text & "' and [INFORME]='" & txtInforme.Text & "'"
        ''    Dim comando As New SqlCommand(R, conexion2019)
        ''    Dim lector As SqlDataReader
        ''    lector = comando.ExecuteReader
        ''    lector.Read()

        ''    txtFolio.Text = lector(0)
        ''    txtClaveEmpresa.Text = lector(1)
        ''    txtEmpresa.Text = lector(2)

        ''    If ((lector(3) Is DBNull.Value) OrElse (lector(3) Is Nothing)) Then
        ''        DTPRecepcion.Visible = False
        ''        txtfechaRecep.Text = "Fecha vacía"
        ''        txtfechaRecep.Visible = True
        ''    Else
        ''        DTPRecepcion.Text = lector(3)
        ''    End If

        ''    If ((lector(4) Is DBNull.Value) OrElse (lector(4) Is Nothing)) Then
        ''        DTPCalibracion.Visible = False

        ''        txtFechaCalib.Text = "Fecha vacía"
        ''        txtFechaCalib.Visible = True
        ''    Else
        ''        DTPCalibracion.Text = lector(4)
        ''    End If

        ''    If ((lector(5) Is DBNull.Value) OrElse (lector(5) Is Nothing)) Then
        ''        DTPEmision.Visible = False

        ''        txtFechaEmision.Text = "Fecha vacía"
        ''        txtFechaEmision.Visible = True
        ''    Else
        ''        DTPEmision.Text = lector(5)
        ''    End If


        ''    txtCatalogo.Text = lector(6)
        ''    txtCatalogo2.Text = lector(7)
        ''    txtActualizacion.Text = lector(8)
        ''    txtInstrumento.Text = lector(9)
        ''    txtInterval.Text = lector(10)
        ''    txtMarca.Text = lector(11)
        ''    txtModelo.Text = lector(12)
        ''    txtSerie.Text = lector(13)
        ''    txtID.Text = lector(14)
        ''    txtPatron1.Text = lector(15)
        ''    txtPatron2.Text = lector(16)
        ''    txtPatron3.Text = lector(17)
        ''    txtCalibro.Text = lector(18)
        ''    cboStatus.Text = lector(19)

        ''    If ((lector(20) Is DBNull.Value) OrElse (lector(20) Is Nothing)) Then
        ''        DTPRegistro.Visible = False
        ''        ' txtFechaRegistro.Text = ""
        ''        txtFechaRegistro.Text = "Fecha vacía"
        ''        txtFechaRegistro.Visible = True
        ''    Else
        ''        DTPRegistro.Text = lector(20)
        ''    End If



        ''    If ((lector(21) Is DBNull.Value) OrElse (lector(21) Is Nothing)) Then
        ''        DTPLaboratorio.Visible = False
        ''        ' txtEntradaLab.Text = ""
        ''        txtEntradaLab.Text = "Fecha vacía"
        ''        txtEntradaLab.Visible = True
        ''    Else
        ''        DTPLaboratorio.Text = lector(21)
        ''    End If

        ''    txtPatron4.Text = lector(22)
        ''    txtPatron5.Text = lector(23)
        ''    txtPatron6.Text = lector(24)
        ''    txtPatron7.Text = lector(25)
        ''    txtPatron8.Text = lector(26)
        ''    txtPatron9.Text = lector(27)
        ''    txtPatron10.Text = lector(28)
        ''    txtFirma.Text = lector(29)
        ''    lector.Close()
        ''    conexion2019.Close()
        ''Catch ex As Exception

        ''End Try

        Try
            conexion2019.Open()
            Dim R As String
            R = "update [METASINF-2019-3].[dbo].[INFORMES-SERVICIOS] set [FECHACALIB] = '" & DTPCalibracion.Value.ToShortDateString & "', [Fecha_Reg] = '" & DTPRegistro.Value.ToShortDateString & "', [FECHA-EMISION] = '" & DTPEmision.Value.ToShortDateString & "',
                [CALIBRO] = '" & txtCalibro.Text & "', [Patron1] = '" & txtPatron1.Text & "', [Patron2] = '" & txtPatron2.Text & "', [Patron3] = '" & txtPatron3.Text & "', [Patron4] = '" & txtPatron4.Text & "', [Patron5] = '" & txtPatron5.Text & "', 
                [Patron6] = '" & txtPatron6.Text & "', [Patron7] = '" & txtPatron7.Text & "', [Patron8] = '" & txtPatron8.Text & "', [Patron9] = '" & txtPatron9.Text & "', [Patron10] = '" & txtPatron10.Text & "', [Firma] = '" & txtFirma.Text & "',
                [Status] = '" & cboStatus.Text & "',[ObservacionStatus] = '" & txtObservaciones.Text & "', [FechaRecepLab] = '" & DTPRecepcion.Value.ToShortDateString & "' where MAGNITUD='MACF-TC' and INFORME='1565/19' "
            Dim comando As New SqlCommand(R, conexion2019)
            comando.ExecuteNonQuery()
            MsgBox("Registro modificado correctamente.", MsgBoxStyle.Information)
            conexion2019.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error del Sistema")
        End Try
    End Sub

    Private Sub txtInforme_KeyDown(sender As Object, e As KeyEventArgs) Handles txtInforme.KeyDown

        'Try
        Select Case e.KeyData
                Case Keys.Enter
                    If txtMagnitud.Text.Equals("") Or txtInforme.Text.Equals("") Then
                        MsgBox("No podemos guardar campos vacios.", MsgBoxStyle.Critical)
                    Else
                        'Try
                        Dim R As String
                            MetodoMetasInf2019()
                            ' conexion2019.Open()
                            R = "SELECT isnull([Folio],'-') as Folio 
                                    ,isnull([ClavecontactoConsign],'-')as claveempresa 
                                    ,isnull([Empresa],'-')as empresa 
                                    ,[FECHARECEP]as fechaRecep 
                                    ,[FECHACALIB]as fechaCalib 
                                    ,[FECHA-EMISION]as FechaEmision 
                                    ,isnull([ServCatalogo1],'-')as Catalogo1 
                                    ,isnull([ServCatalogo2],'-')as catalogo2 
                                    ,isnull([ServiciosAdicionales],'-')as adicionales 
                                    ,isnull([TIPO],'-')as Tipo 
                                    ,isnull([ALCANCE],'-')as alcance 
                                    ,isnull([MARCA],'-')as marca 
                                    ,isnull([MODELO],'-')as modelo 
                                    ,isnull([Serie],'-')as serie 
                                    ,isnull([ID],'-')as id 
                                    ,isnull([Patron1],'-')as patron1 
                                    ,isnull([Patron2],'-')as patron2
                                    ,isnull([Patron3],'-')as patron3 
                                    ,isnull([CALIBRO],'-')as calibro 
                                    ,isnull([Status],'-')as status 
                                    ,[Fecha_Reg]as fechaRegistro 
                                    ,[FechaRecepLab]as fechaReceplab 
                                    ,isnull([Patron4],'-')as patron4 
                                    ,isnull([Patron5],'-')as patron5 
                                    ,isnull([Patron6],'-')as patron6 
                                    ,isnull([Patron7],'-')as patron7 
                                    ,isnull([Patron8],'-')as patron8 
                                    ,isnull([Patron9],'-')as patron9 
                                    ,isnull([Patron10],'-')as patron10 
                                    ,isnull([Firma],'-')as Firma
                                    ,isnull([Usuario],'-')as Usuario
                                   
                                    ,isnull([Accesorios],'-')as [Accesorios]
                                    ,isnull([PUNTOS],'-')as [PUNTOS]
                                    ,isnull([Patron2],'-')as [Patron2]                                 
                                    ,isnull([Observaciones],'-')as [Observaciones]
                                    ,isnull([CveOperador],'-')as [CveOperador]
                                    ,isnull([EmpresaEmision],'-')as [EmpresaEmision]
                                    ,isnull([DirCalleEmision],'-')as [DirCalleEmision]
                                    ,isnull([DirColEmision],'-')as [DirColEmision]
                                    ,isnull([DirCiudadEmision],'-')as [DirCiudadEmision]
                                    ,isnull([DirEdoProvEmision],'-')as [DirEdoProvEmision]
                                    ,isnull([DirPaisEmision],'-')as [DirPaisEmision]
                                    ,isnull([DirCPEmision],'-')as [DirCPEmision]
                                    ,isnull([Sv1Ajuste],'-')as [Sv1Ajuste]
                                    ,isnull([Sv3Matto],'-')as [Sv3Matto]
                                    ,isnull([Sv5COM02],'-')as [Sv5COM02]
                                    ,isnull([Sv6IntervaloRe],'-')as [Sv6IntervaloRe]
                                    ,isnull([Sv7Curva],'-')as [Sv7Curva]
                                    ,isnull([Sv8Idioma],'-')as [Sv8Idioma]
                                    ,isnull([Sv9Calendar],'-')as [Sv9Calendar]
                                    ,isnull([ObservacionStatus],'-')as [ObservacionStatus]
                                  FROM [INFORMES-SERVICIOS] where [MAGNITUD]='" & txtMagnitud.Text & "' and [INFORME]='" & txtInforme.Text & "'"
                            Dim comando As New SqlCommand(R, conexion2019)
                            Dim lector As SqlDataReader
                            lector = comando.ExecuteReader
                            lector.Read()
                            txtFolio.Text = lector(0)
                            txtClaveEmpresa.Text = lector(1)
                            txtEmpresa.Text = lector(2)

                            If ((lector(3) Is DBNull.Value) OrElse (lector(3) Is Nothing)) Then
                                DTPRecepcion.Visible = False
                                txtfechaRecep.Text = "Fecha vacía"
                                txtfechaRecep.Visible = True
                            Else
                                DTPRecepcion.Text = lector(3)
                            End If

                            If ((lector(4) Is DBNull.Value) OrElse (lector(4) Is Nothing)) Then
                                DTPCalibracion.Visible = False

                                txtFechaCalib.Text = "Fecha vacía"
                                txtFechaCalib.Visible = True
                            Else
                                DTPCalibracion.Text = lector(4)
                            End If

                            If ((lector(5) Is DBNull.Value) OrElse (lector(5) Is Nothing)) Then
                                DTPEmision.Visible = False

                                txtFechaEmision.Text = "Fecha vacía"
                                txtFechaEmision.Visible = True
                            Else
                                DTPEmision.Text = lector(5)
                            End If
                            txtCatalogo.Text = lector(6)
                            txtCatalogo2.Text = lector(7)
                            txtActualizacion.Text = lector(8)
                            txtInstrumento.Text = lector(9)
                            txtInterval.Text = lector(10)
                            txtMarca.Text = lector(11)
                            txtModelo.Text = lector(12)
                            txtSerie.Text = lector(13)
                            txtID.Text = lector(14)
                            txtPatron1.Text = lector(15)
                            txtPatron2.Text = lector(16)
                            txtPatron3.Text = lector(17)
                            txtCalibro.Text = lector(18)
                            cboStatus.Text = lector(19)
                            If ((lector(20) Is DBNull.Value) OrElse (lector(20) Is Nothing)) Then
                                DTPRegistro.Visible = False
                                ' txtFechaRegistro.Text = ""
                                txtFechaRegistro.Text = "Fecha vacía"
                                txtFechaRegistro.Visible = True
                            Else
                                DTPRegistro.Text = lector(20)
                            End If
                            If ((lector(21) Is DBNull.Value) OrElse (lector(21) Is Nothing)) Then
                                DTPLaboratorio.Visible = False
                                ' txtEntradaLab.Text = ""
                                txtEntradaLab.Text = "Fecha vacía"
                                txtEntradaLab.Visible = True
                            Else
                                DTPLaboratorio.Text = lector(21)
                            End If
                            txtPatron4.Text = lector(22)
                            txtPatron5.Text = lector(23)
                            txtPatron6.Text = lector(24)
                            txtPatron7.Text = lector(25)
                            txtPatron8.Text = lector(26)
                            txtPatron9.Text = lector(27)
                            txtPatron10.Text = lector(28)
                            txtFirma.Text = lector(29)
                            txtUsuario.Text = lector(30)
                            ''--------------------------------
                            txtAccesorios.Text = lector(31)
                            txtPuntos.Text = lector(32)
                            txtPatron2.Text = lector(33)
                            txtObservacionesAdicionales.Text = lector(34)


                            txtOperador.Text = lector(35)
                            If lector(35).ToString.Equals("10") Then
                                Label42.Text = "Natalia García"
                            ElseIf lector(35).ToString.Equals("4") Then
                                Label42.Text = "Ximena"
                            ElseIf lector(35).ToString.Equals("8") Then
                                Label42.Text = "Tonaxy Reyes"
                            ElseIf lector(35).ToString.Equals("16") Then
                                Label42.Text = "José Armando"
                            ElseIf lector(35).ToString.Equals("17") Then
                                Label42.Text = "Pedro Palacios"
                            ElseIf lector(35).ToString.Equals("18") Then
                                Label42.Text = "Fernando González"
                            ElseIf lector(35).ToString.Equals("20") Then
                                Label42.Text = "Marcos Adrian Aguilar"
                            ElseIf lector(35).ToString.Equals("21") Then
                                Label42.Text = "Marlene Peña"
                            ElseIf lector(35).ToString.Equals("22") Then
                                Label42.Text = "Carlos Ramirez"
                            End If


                            txtNombreEmpresa.Text = lector(36)
                            txtDomicilio.Text = lector(37)
                            txtColonia.Text = lector(38)
                            txtCiudad.Text = lector(39)
                            txtEstado.Text = lector(40)
                            txtPais.Text = lector(41)
                            txtCodigoPostal.Text = lector(42)


                            'txtAjuste.Text = lector(43)
                            If lector(43).ToString.Equals("1") Then
                                txtAjuste.Text = "Sin ajuste"

                            ElseIf lector(43).ToString.Equals("2") Then
                                txtAjuste.Text = "Ajuste 50%"

                            ElseIf lector(43).ToString.Equals("3") Then
                                txtAjuste.Text = "Ajuste 100%"

                            ElseIf lector(43).ToString.Equals("4") Then
                                txtAjuste.Text = "Pintado, despintado, rotulado y ajuste de pesa paralelepípeda"

                            End If


                            'txtMantenimiento.Text = lector(44)
                            If lector(44).ToString.Equals("5") Then
                                txtMantenimiento.Text = "Sin mantenimiento"

                            ElseIf lector(44).ToString.Equals("6") Then
                                txtMantenimiento.Text = "Mantenimiento preventivo"

                            End If


                            ' txtEvaluacion.Text = lector(45)
                            If lector(45).ToString.Equals("10") Then
                                txtEvaluacion.Text = "Sin evaluación de la conformidad"

                            ElseIf lector(45).ToString.Equals("11") Then
                                txtEvaluacion.Text = "Con evaluación de la conformidad de los resultados finales, incluyendo incertidumbres"

                            ElseIf lector(45).ToString.Equals("12") Then
                                txtEvaluacion.Text = "Con evaluación de la conformidad eléctrica"

                            End If


                            'txtIntervaloDeCal.Text = lector(46)
                            If lector(46).ToString.Equals("13") Then
                                txtIntervaloDeCal.Text = "Sin calculo de intervalo de calibración"

                            ElseIf lector(46).ToString.Equals("14") Then
                                txtIntervaloDeCal.Text = "Calculo de intervalo de calibración"

                            End If




                            'txtCurvaAjuste.Text = lector(47)
                            If lector(47).ToString.Equals("15") Then
                                txtCurvaAjuste.Text = "Sin curva de ajuste"

                            ElseIf lector(47).ToString.Equals("16") Then
                                txtCurvaAjuste.Text = "Curva de ajuste con residuales y evaluación de incertidumbre"

                            ElseIf lector(47).ToString.Equals("17") Then
                                txtCurvaAjuste.Text = "Curva de ajuste eléctrica"

                            End If




                            ' txtIdioma.Text = lector(48)
                            If lector(48).ToString.Equals("18") Then
                                txtIdioma.Text = "Español"

                            ElseIf lector(48).ToString.Equals("19") Then
                                txtIdioma.Text = "Portada bilingüe español + inglés"

                            ElseIf lector(48).ToString.Equals("20") Then
                                txtIdioma.Text = "Certificado completo bilingüe español + inglés"

                            End If


                            'txtCalendarizacion.Text = lector(49)
                            If lector(49).ToString.Equals("21") Then
                                txtCalendarizacion.Text = "Normal"

                            ElseIf lector(49).ToString.Equals("22") Then
                                txtCalendarizacion.Text = "Programado"

                            ElseIf lector(49).ToString.Equals("23") Then
                                txtCalendarizacion.Text = "Urgente"

                            End If

                            txtObservaciones.Text = lector(50)

                            ''--------------------------------
                            lector.Close()
                            conexion2019.Close()
                        'Catch ex As Exception
                        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error del Sistema")
                        'End Try
                    End If
            End Select
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error del Sistema")
        'End Try
    End Sub
End Class