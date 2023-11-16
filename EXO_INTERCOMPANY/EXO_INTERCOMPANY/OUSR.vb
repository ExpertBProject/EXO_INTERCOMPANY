Imports System.Xml
Imports SAPbouiCOM

Public Class OUSR
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaCampos()
            Crear_tabla_Log()
        End If
    End Sub

    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OUSR.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_OUSR", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub

    Public Sub Crear_tabla_Log()
        Dim sSQL As String = ""
        Dim bResultado As Boolean = False
        Dim sBBDD As String = "" : Dim sBBDDDMaster As String = ""
        Try
            If objGlobal.refDi.comunes.esAdministrador Then
                sBBDD = objGlobal.refDi.compañia.CompanyDB
                sSQL = "SELECT TOP 1 ""U_EXO_BBDD"" FROM """ & sBBDD & """.""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='M' "
                sBBDDDMaster = objGlobal.refDi.SQL.sqlStringB1(sSQL)

                ' Si estamos en la master enviamos datos a los destinos
                If sBBDD = sBBDDDMaster Then
#Region "Crear Tabla LOG"
                    sSQL = objGlobal.funciones.leerEmbebido(Me.GetType(), "CREATE_TABLE_LOG.sql")
                    sSQL = sSQL.Replace("TABLE ", "TABLE """ & sBBDD & """.")
                    bResultado = objGlobal.refDi.SQL.executeNonQuery(sSQL)

                    If bResultado = True Then
                        objGlobal.SBOApp.StatusBar.SetText("Creada Tabla EXO_LOG_INTERCOMPANY", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#Region "Creamos la consulta formateada"
                        sSQL = ""
                        Dim oOUQR As SAPbobsCOM.UserQueries = Nothing
                        Dim oRs As SAPbobsCOM.Recordset = Nothing
                        Dim sIntrnalKey As String = ""

                        'Comprobamos si existe la consulta formateada dentro de la categoría General, que devuelve 
                        'los errores de integración de estructuras/rutas/herramientas KTB
                        sSQL = "SELECT t1.""IntrnalKey"" 
                        FROM """ & sBBDD & """.""OUQR"" t1
                        WHERE t1.""QCategory"" = -1 
                        And t1.""QName"" = 'LOG INTERCOMPANY'"
                        sIntrnalKey = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                        If sIntrnalKey = "" Then
                            'Creamos la consulta formateada dentro de la categoría General, que devuelve 
                            'los errores de integración de estructuras/rutas/herramientas KTB
                            oOUQR = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries), SAPbobsCOM.UserQueries)

                            sSQL = "SELECT t1.""DATEPROCESS"" ""Fecha proceso"", t1.""HORA"" ""HORA"", t1.""BBDD"" ""BBDD"", t1.""USUARIO"" ""USUARIO"", t1.""MODELO"" ""MODELO"", t1.""MESSAGE"" ""Mensaje"" 
                                    FROM """ & sBBDD & """.""EXO_LOG_INTERCOMPANY"" t1"
                            oOUQR.Query = sSQL
                            oOUQR.QueryCategory = -1 'General
                            oOUQR.QueryDescription = "LOG INTERCOMPANY"
                            oOUQR.QueryType = SAPbobsCOM.UserQueryTypeEnum.uqtWizard

                            If oOUQR.Add <> 0 Then
                                Throw New Exception(objGlobal.compañia.GetLastErrorCode & " " & objGlobal.compañia.GetLastErrorDescription)
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("Consulta formateada creada", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            End If

                            objGlobal.compañia.GetNewObjectCode(sIntrnalKey)
                            sIntrnalKey = sIntrnalKey.Split(vbTab.ToCharArray)(0)
                        End If
#End Region
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("Revisar si existe tabla EXO_LOG_INTERCOMPANY", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
#End Region
                End If

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS_INTERCOMPANY.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim bEstado As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "20700"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else

                Select Case infoEvento.FormTypeEx
                    Case "20700"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then
                                    If Intercompany_After(oForm) = False Then
                                        Return False
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    If Intercompany_After(oForm) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function Intercompany_After(ByRef oform As SAPbouiCOM.Form) As Boolean

        Dim sMensaje As String = ""
        Dim sBBDD As String = "" : Dim sBBDDDMaster As String = "" : Dim sUser As String = "" : Dim sPass As String = ""
        Dim sUSR As String = "" : Dim sUSRNAME As String = ""
        Dim sSQL As String = ""
        Dim sPassUDF As String = ""
        Dim OdtEmpresas As System.Data.DataTable = Nothing : Dim oCompanyDes As SAPbobsCOM.Company = Nothing : Dim oCompanyMaster As SAPbobsCOM.Company = Nothing
        Dim oUser As SAPbobsCOM.Users = Nothing
        Dim sExisteAutoriz As String = ""
        Dim bResultado = False
        Intercompany_After = False

        Try
            If (oform.Mode = BoFormMode.fm_ADD_MODE Or oform.Mode = BoFormMode.fm_UPDATE_MODE) Then
                sUSR = oform.DataSources.DBDataSources.Item("OUSR").GetValue("USER_CODE", 0).ToString.Trim
                sUSRNAME = oform.DataSources.DBDataSources.Item("OUSR").GetValue("U_NAME", 0).ToString.Trim
                Try
                    sPassUDF = oform.DataSources.DBDataSources.Item("OUSR").GetValue("U_EXO_PASS", 0).ToString.Trim
                Catch ex As Exception

                End Try

                sBBDD = objGlobal.refDi.compañia.CompanyDB
                sSQL = "SELECT TOP 1 ""U_EXO_BBDD"" FROM ""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='M' "
                sBBDDDMaster = objGlobal.refDi.SQL.sqlStringB1(sSQL)

                ' Si estamos en la master enviamos datos a los destinos
                If sBBDD = sBBDDDMaster Then
                    OdtEmpresas = New System.Data.DataTable
                    OdtEmpresas.Clear()
                    sSQL = "SELECT * FROM ""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='D' "
                    OdtEmpresas = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                    If OdtEmpresas.Rows.Count > 0 Then
                        oUser = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers), SAPbobsCOM.Users)
                        'Buscamos el código de usuario
                        sSQL = "SELECT ""USERID"" FROM OUSR WHERE ""USER_CODE""='" & sUSR & "'"
                        Dim sCodUSR As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                        If oUser.GetByKey(CType(sCodUSR, Integer)) = True Then
                            objGlobal.SBOApp.StatusBar.SetText("Se va a proceder a recorrer las SOCIEDADES...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
#Region "Control tabla LOG"
#Region "Borrar EXO_LOG_INTERCOMPANY"
                            sSQL = "DELETE FROM """ & sBBDD & """.""EXO_LOG_INTERCOMPANY"" "
                            bResultado = objGlobal.refDi.SQL.executeNonQuery(sSQL)
                            If bResultado = True Then
                                objGlobal.SBOApp.StatusBar.SetText("Borrado todos los datos de la tabla ""EXO_LOG_INTERCOMPANY"".", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            End If
#End Region
#Region "Crear Registro EXO_EDI_LOG"
                            Dim sFecha As String = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")
                            EXO_GLOBALES.LogTabla(objGlobal.compañia, objGlobal, sFecha, objGlobal.compañia.CompanyDB, objGlobal.compañia.UserName, "", "#####                     INICIO LOG INTERCOMPANY                 #####", "INFO")
#End Region
#End Region
                            For Each dr As DataRow In OdtEmpresas.Rows
                                Try
                                    sBBDD = dr.Item("U_EXO_BBDD").ToString : sUser = dr.Item("U_EXO_USER").ToString : sPass = dr.Item("U_EXO_PASS").ToString
                                    'If sBBDD = "SEMA_PROD" Or sBBDD = "RANTI" Or sBBDD = "SIYCF" Then
                                    '    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & sBBDD & ". No se puede sincronizar Usuario: " & sUSR & " - " & sUSRNAME, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    'Else
                                    EXO_CONEXIONES.Connect_Company(oCompanyDes, objGlobal, sUser, sPass, sBBDD)
                                    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Sincronizando Usuario: " & sUSR & " - " & sUSRNAME, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    EXO_GLOBALES.Sincroniza_User_Master(oUser, oCompanyDes, objGlobal)
                                    'End If
                                Catch ex As Exception
                                    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    EXO_GLOBALES.LogTabla(objGlobal.compañia, objGlobal, sFecha, sBBDD, sUSR, "", ex.Message, "ERROR")
                                Finally
                                    Try
                                        objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                        EXO_CONEXIONES.Disconnect_Company(oCompanyDes)
                                    Catch ex As Exception

                                    End Try
                                End Try
                            Next

                            Try
                                'sBBDD = "EMPRESA_CONSOLIDACION" : sUser = "manager" : sPass = "Sol@ri@123"

                                'EXO_CONEXIONES.Connect_Company(oCompanyDes, objGlobal, sUser, sPass, sBBDD)
                                'objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Sincronizando Usuario: " & sUSR & " - " & sUSRNAME, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                'EXO_GLOBALES.Sincroniza_User_Master(oUser, oCompanyDes, objGlobal)

                            Catch ex As Exception
                                objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & sBBDD & ". Error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Finally
                                Try
#Region "Crear Registro EXO_EDI_LOG"
                                    sFecha = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")
                                    EXO_GLOBALES.LogTabla(objGlobal.compañia, objGlobal, sFecha, objGlobal.compañia.CompanyDB, objGlobal.compañia.UserName, "", "#####                       FIN LOG INTERCOMPANY                   #####", "INFO")
#End Region
                                    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    EXO_CONEXIONES.Disconnect_Company(oCompanyDes)
                                Catch ex As Exception

                                End Try
                            End Try


                            EXO_GLOBALES.Sincroniza_User_Autoriz(objGlobal, sUSR, sPassUDF)

                        Else
                            objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el Usuario: " & sUSR & " - " & sUSRNAME & ". No se puede sincronizar.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        End If
                    End If
                End If
            End If

            Intercompany_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oUser, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyDes, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyMaster, Object))
        End Try
    End Function
End Class
