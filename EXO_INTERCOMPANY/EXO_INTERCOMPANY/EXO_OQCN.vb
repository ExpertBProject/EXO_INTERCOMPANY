Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_OQCN
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
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "521"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "521"
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
                        Case "521"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "521"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
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
    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oCheckBox As SAPbouiCOM.Button = Nothing
        Dim oItem As SAPbouiCOM.Item = Nothing
        Dim sSQL As String = ""
        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            oForm.Visible = False
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Boton Inter"
            oItem = oForm.Items.Add("btnInterQ", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5
            oItem.Width = oForm.Items.Item("2").Width * 2
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.LinkTo = "2"
            oCheckBox = CType(oItem.Specific, Button)
            oCheckBox.Caption = "Traspaso a Emp. Inter"
#End Region

            oForm.Visible = True

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw exCOM
        Catch ex As Exception
            If oForm IsNot Nothing Then oForm.Visible = True

            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sMensaje As String = ""
        Dim sBBDD As String = "" : Dim sBBDDDMaster As String = "" : Dim sUser As String = "" : Dim sPass As String = ""
        Dim sCatCode As String = "" : Dim sCatName As String = ""
        Dim sSQL As String = ""
        Dim OdtEmpresas As System.Data.DataTable = Nothing : Dim oCompanyDes As SAPbobsCOM.Company = Nothing
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
        Dim sboUserQueryCatOrigen As SAPbobsCOM.QueryCategories = Nothing : Dim sboUserQueryCatDes As SAPbobsCOM.QueryCategories = Nothing

        EventHandler_ItemPressed_After = False

        Try

            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "btnInterQ" Then
                sCatName = oForm.DataSources.UserDataSources.Item(0).ValueEx
                If sCatName = "" Then
                    objGlobal.SBOApp.StatusBar.SetText("Seleccione una categoría para traspasar", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    Return False
                Else
                    sSQL = "SELECT ""CategoryId"" FROM """ & objGlobal.compañia.CompanyDB & """.""OQCN"" Where ""CatName""='" & sCatName & "' "
                    sCatCode = objGlobal.refDi.SQL.sqlStringB1(sSQL)
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
                            sboUserQueryCatOrigen = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories), SAPbobsCOM.QueryCategories)

                            If sboUserQueryCatOrigen.GetByKey(CType(sCatCode, Integer)) = True Then
                                objGlobal.SBOApp.StatusBar.SetText("Se va a proceder a recorrer las SOCIEDADES...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                For Each dr As DataRow In OdtEmpresas.Rows
                                    Try
                                        'sBBDD = dr.Item("U_EXO_BBDD").ToString : sUser = dr.Item("U_EXO_USER").ToString : sPass = dr.Item("U_EXO_PASS").ToString
                                        'If sBBDD = "SEMA_PROD" Or sBBDD = "RANTI" Or sBBDD = "SIYCF" Then
                                        '    objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & sBBDD & ". No se puede sincronizar las querys de la categoría " & sCatName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                        'Else
                                        EXO_CONEXIONES.Connect_Company(oCompanyDes, objGlobal, sUser, sPass, sBBDD)
                                        objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Sincronizando Querys de la Categoría: " & sCatName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                        EXO_GLOBALES.CrearQuerysDesde_CAT_Master(sCatCode, sCatName, oCompanyDes, objGlobal)
                                        objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Fin Sincronización.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                                        EXO_CONEXIONES.Disconnect_Company(oCompanyDes)
                                        'End If
                                    Catch ex As Exception
                                        objGlobal.SBOApp.StatusBar.SetText("Sociedad: " & oCompanyDes.CompanyName & ". Error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    Finally

                                    End Try
                                Next
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("Error inesperado. No se encuentra la categoría " & sCatName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    End If
                End If

            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
#Region "Liberar"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(sboUserQueryCatOrigen, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompanyDes, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCmpSrv, Object))
#End Region

        End Try
    End Function
End Class
