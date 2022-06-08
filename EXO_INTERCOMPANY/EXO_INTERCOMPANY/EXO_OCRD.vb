Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OCRD
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables Globales"
    Dim _sEstado_Formulario As String = ""
#End Region
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
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If infoEvento.ActionSuccess Then
                                        If EventHandler_ItemPressed_After(infoEvento) = False Then
                                            Return False
                                        End If
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "134"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.TypeEx = "134" And pVal.ItemUID = "1" Then
                Select Case _sEstado_Formulario
                    Case "3" : objGlobal.SBOApp.ActivateMenuItem("1289") 'Ultimo dato
                    Case "2" : objGlobal.SBOApp.ActivateMenuItem("1304") 'Actualizar
                    Case "0" : objGlobal.SBOApp.ActivateMenuItem("1289") 'Ultimo dato
                End Select
                _sEstado_Formulario = ""
            End If


            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function Intercompany_After(ByRef oform As SAPbouiCOM.Form) As Boolean
        Dim sMensaje As String = ""
        Dim sBBDD As String = "" : Dim sBBDDDMaster As String = "" : Dim sUser As String = "" : Dim sPass As String = ""
        Dim sCardCode As String = "" : Dim sCardName As String = "" : Dim sCardType As String = "" : Dim sSerie As String = ""
        Dim sSQL As String = ""
        Dim OdtEmpresas As System.Data.DataTable = Nothing : Dim oCompanyDes As SAPbobsCOM.Company = Nothing : Dim oCompanyMaster As SAPbobsCOM.Company = Nothing
        Dim oOCRD As SAPbobsCOM.BusinessPartners = Nothing
        Dim bHaSincronizado = False
        Intercompany_After = False

        Try
            If (oform.Mode = BoFormMode.fm_ADD_MODE Or oform.Mode = BoFormMode.fm_UPDATE_MODE) Then
                sCardCode = oform.DataSources.DBDataSources.Item("OCRD").GetValue("CardCode", 0).ToString.Trim
                sCardName = oform.DataSources.DBDataSources.Item("OCRD").GetValue("CardName", 0).ToString.Trim
                sCardType = oform.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).ToString.Trim
                sSerie = CType(oform.Items.Item("1320002080").Specific, SAPbouiCOM.ComboBox).Selected.Description
                sBBDD = objGlobal.refDi.compañia.CompanyDB
                sSQL = "SELECT TOP 1 ""U_EXO_BBDD"" FROM ""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='M' ORDER BY ""LineId""  "
                sBBDDDMaster = objGlobal.refDi.SQL.sqlStringB1(sSQL)

                ' Si estamos en la master enviamos datos a los destinos
                If sBBDD = sBBDDDMaster And ((sCardType = "S" And oform.Mode = BoFormMode.fm_UPDATE_MODE) Or (sCardType = "C" And sSerie = "CI")) Then
                    'If sBBDD = sBBDDDMaster And ((sCardType = "C" And sSerie = "CI")) Then
                    OdtEmpresas = New System.Data.DataTable
                    OdtEmpresas.Clear()
                    sSQL = "SELECT * FROM ""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='D' ORDER BY ""LineId"" "
                    OdtEmpresas = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                    If OdtEmpresas.Rows.Count > 0 Then
                        oOCRD = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
                        If oOCRD.GetByKey(sCardCode) = True Then
                            objGlobal.SBOApp.StatusBar.SetText("Se va a proceder a recorrer las SOCIEDADES...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            For Each dr As DataRow In OdtEmpresas.Rows
                                Try
                                    sBBDD = dr.Item("U_EXO_BBDD").ToString : sUser = dr.Item("U_EXO_USER").ToString : sPass = dr.Item("U_EXO_PASS").ToString
                                    EXO_CONEXIONES.Connect_Company(oCompanyDes, objGlobal, sUser, sPass, sBBDD)
                                    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Sincronizando Proveedor: " & sCardCode & " - " & sCardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    EXO_GLOBALES.Sincroniza_proveedor_Master(oOCRD, oCompanyDes, objGlobal, sSerie)
                                Catch ex As Exception
                                    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Finally
                                    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    EXO_CONEXIONES.Disconnect_Company(oCompanyDes)
                                End Try
                            Next
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el IC: " & sCardCode & " - " & sCardName & ". No se puede sincronizar.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        End If
                    End If
                ElseIf sBBDD <> sBBDDDMaster And sCardType = "S" Then
                    OdtEmpresas = New System.Data.DataTable
                    OdtEmpresas.Clear()
                    sSQL = "SELECT * FROM ""@EXO_IPANELL"" WHERE ""Code""='INTERCOMPANY' and ""U_EXO_TIPO""='M' "
                    OdtEmpresas = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                    If OdtEmpresas.Rows.Count > 0 Then
                        oOCRD = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
                        If oOCRD.GetByKey(sCardCode) = True Then
                            objGlobal.SBOApp.StatusBar.SetText("Se va a buscar datos en la empresa Consolidación...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                            For Each dr As DataRow In OdtEmpresas.Rows
                                Try
                                    sBBDD = dr.Item("U_EXO_BBDD").ToString : sUser = dr.Item("U_EXO_USER").ToString : sPass = dr.Item("U_EXO_PASS").ToString
                                    EXO_CONEXIONES.Connect_Company(oCompanyMaster, objGlobal, sUser, sPass, sBBDD)
                                    objGlobal.SBOApp.StatusBar.SetText("Sincronizando Proveedor: " & sCardCode & " - " & sCardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    bHaSincronizado = EXO_GLOBALES.Sincroniza_proveedor(oOCRD, oCompanyMaster, objGlobal)
                                    If bHaSincronizado = False Then
                                        objGlobal.SBOApp.StatusBar.SetText("Proveedor no creado en la empresa Consolidación. Por favor, dadlo de alta en la empresa Consolidación. ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                        oOCRD.Remove()
                                        objGlobal.SBOApp.MessageBox("Proveedor no creado en la empresa Consolidación. Por favor, dadlo de alta en la empresa Consolidación. ")
                                    End If
                                Catch ex As Exception
                                    objGlobal.SBOApp.StatusBar.SetText("Error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Finally
                                    objGlobal.SBOApp.StatusBar.SetText("Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                    EXO_CONEXIONES.Disconnect_Company(oCompanyMaster)
                                    If bHaSincronizado = True Then
                                        Select Case oform.Mode
                                            Case oform.Mode.fm_ADD_MODE : _sEstado_Formulario = "3"
                                            Case oform.Mode.fm_UPDATE_MODE : _sEstado_Formulario = "2"
                                        End Select
                                    Else
                                        _sEstado_Formulario = "0"
                                    End If
                                End Try
                            Next
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el IC: " & sCardCode & " - " & sCardName & ". No se puede sincronizar.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        End If
                    End If
                End If
            End If

            Intercompany_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oform.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oform.Freeze(False)
            Throw ex
        Finally
            oform.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oform)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRD, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyDes, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyMaster, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim bEstado As String = ""
        Dim sLicTradNum As String = "" : Dim sCardType As String = "" : Dim sSerie As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "134"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                sLicTradNum = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("LicTradNum", 0).ToString.Trim
                                sCardType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).ToString.Trim
                                sSerie = CType(oForm.Items.Item("1320002080").Specific, SAPbouiCOM.ComboBox).Selected.Description
                                If EXO_GLOBALES.Comprueba_Proveedor_en_Master(objGlobal, sLicTradNum, sCardType, sSerie) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                sLicTradNum = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("LicTradNum", 0).ToString.Trim
                                sCardType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).ToString.Trim
                                If EXO_GLOBALES.Comprueba_Proveedor_en_Master(objGlobal, sLicTradNum, sCardType, sSerie) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select
                End Select
            Else

                Select Case infoEvento.FormTypeEx
                    Case "134"
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
End Class
