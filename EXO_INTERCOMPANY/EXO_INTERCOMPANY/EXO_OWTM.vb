Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OWTM
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

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
                    Case "50102"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else

                Select Case infoEvento.FormTypeEx
                    Case "50102"
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
        Dim sModelo As String = "" : Dim sModeloName As String = ""
        Dim sSQL As String = ""
        Dim OdtEmpresas As System.Data.DataTable = Nothing : Dim oCompanyDes As SAPbobsCOM.Company = Nothing
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
        Dim oApprovalTemplateService As SAPbobsCOM.ApprovalTemplatesService = Nothing
        Dim oApprovalTemplate As SAPbobsCOM.ApprovalTemplate = Nothing
        Dim oApprovalTemplateParams As SAPbobsCOM.ApprovalTemplateParams = Nothing
        Dim bResultado As Boolean = False

        Intercompany_After = False

        Try
            If (oform.Mode = BoFormMode.fm_ADD_MODE Or oform.Mode = BoFormMode.fm_UPDATE_MODE) Then
                sModelo = oform.DataSources.DBDataSources.Item("OWTM").GetValue("WtmCode", 0).ToString.Trim
                sModeloName = oform.DataSources.DBDataSources.Item("OWTM").GetValue("Name", 0).ToString.Trim
                sBBDD = objGlobal.refDi.compañia.CompanyDB
                sSQL = "SELECT TOP 1 ""U_EXO_BBDD"" FROM ""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='M' "
                sBBDDDMaster = objGlobal.refDi.SQL.sqlStringB1(sSQL)

                ' Si estamos en la master enviamos datos a los destinos
                If sBBDD = sBBDDDMaster Then
                    OdtEmpresas = New System.Data.DataTable
                    OdtEmpresas.Clear()
                    sSQL = "SELECT * FROM ""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='D' ORDER BY ""LineId"" "
                    OdtEmpresas = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                    If OdtEmpresas.Rows.Count > 0 Then
                        oCmpSrv = objGlobal.compañia.GetCompanyService()
                        oApprovalTemplateService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalTemplatesService), SAPbobsCOM.ApprovalTemplatesService)
                        oApprovalTemplate = CType(oApprovalTemplateService.GetDataInterface(SAPbobsCOM.ApprovalTemplatesServiceDataInterfaces.atsdiApprovalTemplate), SAPbobsCOM.ApprovalTemplate)
                        oApprovalTemplateParams = CType(oApprovalTemplateService.GetDataInterface(SAPbobsCOM.ApprovalTemplatesServiceDataInterfaces.atsdiApprovalTemplateParams), SAPbobsCOM.ApprovalTemplateParams)
                        oApprovalTemplateParams.Code = CType(sModelo, Integer)
                        oApprovalTemplate = oApprovalTemplateService.GetApprovalTemplate(oApprovalTemplateParams)

                        If oApprovalTemplate.Name <> "" Then
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
                                    If sBBDD = "SEMA_PROD" Or sBBDD = "RANTI" Or sBBDD = "SIYCF" Then
                                        objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & sBBDD & ". No se puede sincronizar Modelo de autorización: " & sModelo & " - " & sModeloName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    Else
                                        EXO_CONEXIONES.Connect_Company(oCompanyDes, objGlobal, sUser, sPass, sBBDD)
                                        objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Sincronizando Modelo de autorización: " & sModelo & " - " & sModeloName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                        EXO_GLOBALES.Sincroniza_Modelo_Autorización_Master(oApprovalTemplate, oCompanyDes, objGlobal)
                                        objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                        EXO_CONEXIONES.Disconnect_Company(oCompanyDes)
                                    End If
                                Catch ex As Exception
                                    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Finally

                                End Try
                            Next
#Region "Crear Registro EXO_EDI_LOG"
                            sFecha = Now.Year.ToString("0000") & Now.Month.ToString("00") & Now.Day.ToString("00")
                            EXO_GLOBALES.LogTabla(objGlobal.compañia, objGlobal, sFecha, objGlobal.compañia.CompanyDB, objGlobal.compañia.UserName, "", "#####                       FIN LOG INTERCOMPANY                   #####", "INFO")
#End Region
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el modelo de autorización: " & sModelo & " - " & sModeloName & ". No se puede sincronizar.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
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
#Region "Liberar"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oApprovalTemplate, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyDes, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCmpSrv, Object))
#End Region

        End Try
    End Function
End Class
