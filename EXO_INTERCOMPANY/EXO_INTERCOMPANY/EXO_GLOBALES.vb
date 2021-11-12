Imports SAPbouiCOM

Public Class EXO_GLOBALES
    Public Enum FuenteInformacion
        Visual = 1
        Otros = 2
    End Enum
#Region "Funciones formateos datos"
    Public Shared Function DblTextToNumber(ByRef oCompany As SAPbobsCOM.Company, ByVal sValor As String) As Double
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim cValor As Double = 0
        Dim sValorAux As String = "0"
        Dim sSeparadorMillarB1 As String = "."
        Dim sSeparadorDecimalB1 As String = ","
        Dim sSeparadorDecimalSO As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

        DblTextToNumber = 0

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT COALESCE(""DecSep"", ',') ""DecSep"", COALESCE(""ThousSep"", '.') ""ThousSep"" " &
                   "FROM ""OADM"" " &
                   "WHERE ""Code"" = 1"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sSeparadorMillarB1 = oRs.Fields.Item("ThousSep").Value.ToString
                sSeparadorDecimalB1 = oRs.Fields.Item("DecSep").Value.ToString
            End If

            sValorAux = sValor

            If sSeparadorDecimalSO = "," Then
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "." Then sValorAux = "0" & sValorAux

                    If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ""))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "").Replace(".", ","))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            Else
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "," Then sValorAux = "0" & sValorAux

                    If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", "").Replace(",", "."))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", ""))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            End If

            DblTextToNumber = cValor

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Shared Function DblNumberToText(ByRef oCompany As SAPbobsCOM.Company, ByVal cValor As Double, ByVal oDestino As FuenteInformacion) As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim sNumberDouble As String = "0"
        Dim sSeparadorMillarB1 As String = "."
        Dim sSeparadorDecimalB1 As String = ","
        Dim sSeparadorDecimalSO As String = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator

        DblNumberToText = "0"

        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT COALESCE(""DecSep"", ',') ""DecSep"", COALESCE(""ThousSep"", '.') ""ThousSep"" " &
                   "FROM ""OADM"" " &
                   "WHERE ""Code"" = 1"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sSeparadorMillarB1 = oRs.Fields.Item("ThousSep").Value.ToString
                sSeparadorDecimalB1 = oRs.Fields.Item("DecSep").Value.ToString
            End If

            If cValor.ToString <> "" Then
                If sSeparadorMillarB1 = "." AndAlso sSeparadorDecimalB1 = "," Then 'Decimales ES
                    sNumberDouble = cValor.ToString
                Else 'Decimales USA
                    sNumberDouble = cValor.ToString.Replace(",", ".")
                End If
            End If

            If oDestino = FuenteInformacion.Visual Then
                If sSeparadorDecimalSO = "," Then
                    DblNumberToText = sNumberDouble
                Else
                    DblNumberToText = sNumberDouble.Replace(".", ",")
                End If
            Else
                DblNumberToText = sNumberDouble.Replace(",", ".")
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region

    Public Shared Function Sincroniza_proveedor_Master(ByRef oOCRD As SAPbobsCOM.BusinessPartners, ByRef oCompanyDes As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
#Region "Variables"
        Dim oOCRD_Destino As SAPbobsCOM.BusinessPartners = Nothing
        Dim sLicTradNum As String = "" : Dim sCardCode As String = "" : Dim sCardType As String = ""
        Dim sSQL As String = "" : Dim sSQL2 As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sExiste_IC As Boolean = False
        Dim oOCRG As SAPbobsCOM.BusinessPartnerGroups = Nothing : Dim sGrupo As String = "" : Dim oRsGrupos_Des As SAPbobsCOM.Recordset = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOSHP As SAPbobsCOM.ShippingTypes = Nothing : Dim sClase_Expe As String = "" : Dim oRsClase_Expe As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsClase_Expe_Des As SAPbobsCOM.Recordset = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOIDC As SAPbobsCOM.FactoringIndicators = Nothing : Dim sIndicator As String = "" : Dim oRsIndicator As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOOND As SAPbobsCOM.Industries = Nothing : Dim sRamo As String = "" : Dim oRsRamo As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsRamo_Des As SAPbobsCOM.Recordset = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOSLP As SAPbobsCOM.SalesPersons = Nothing : Dim sEmpleado As String = "" : Dim oRsEmpleado As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsEmpleado_Des As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sResponsable As String = "" : Dim oRsResponsable_Des As SAPbobsCOM.Recordset = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sAgente As String = "" : Dim iContactos As Integer = 0 : Dim iDirecciones As Integer = 0
        Dim oOCTG As SAPbobsCOM.PaymentTermsTypes = Nothing : Dim sCondPago As String = "" : Dim oRsCondPago As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sExiste_Cond_pago As Boolean = False : Dim oRsCondPago_Des As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOCDC As SAPbobsCOM.CashDiscount = Nothing
        Dim sDtoPP As String = "" : Dim oRsdtoPP_Des As SAPbobsCOM.Recordset = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sGroupNum As String = "" : Dim sInstNum As String = ""
        Dim oOPYM As SAPbobsCOM.WizardPaymentMethods = Nothing : Dim oRsOPYM As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset) : Dim sExiste_OPYM As Boolean = False
        Dim sPrioridad As String = "" : Dim oRsPrioridad As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOBPP_Destino As SAPbobsCOM.BPPriorities = Nothing
        Dim oRsCamposUsuario As SAPbobsCOM.Recordset = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
#End Region

        Sincroniza_proveedor_Master = False
        Try
            'Primero buscamos si existe el IC con el NIF
            sLicTradNum = oOCRD.FederalTaxID
            Select Case oOCRD.CardType
                Case SAPbobsCOM.BoCardTypes.cSupplier : sCardType = "S"
                Case SAPbobsCOM.BoCardTypes.cLid : sCardType = "L"
                Case SAPbobsCOM.BoCardTypes.cCustomer : sCardType = "C"
            End Select
            sGrupo = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""GroupName"" FROM OCRG WHERE ""GroupCode""='" & oOCRD.GroupCode & "' and ""GroupType""='" & sCardType & "' ")
            sClase_Expe = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""TrnspName"" FROM OSHP WHERE ""TrnspCode""='" & oOCRD.ShippingType & "' ")
            sIndicator = oOCRD.Indicator
            sRamo = CType(oOCRD.Industry, String)
            sEmpleado = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""SlpName"" FROM OSLP WHERE ""SlpCode""='" & oOCRD.SalesPersonCode & "' ")
            sResponsable = oOCRD.AgentCode
            sCondPago = oObjGlobal.refDi.SQL.sqlStringB1("SELECT ""PymntGroup"" FROM OCTG WHERE ""GroupNum""='" & oOCRD.PayTermsGrpCode & "' ")
            sGroupNum = CType(oOCRD.PayTermsGrpCode, String)
            oRs = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = "SELECT ""CardCode"" FROM OCRD WHERE ""LicTradNum""='" & sLicTradNum & "' and ""CardType""='" & sCardType & "' "
            oRs.DoQuery(sSQL)
            oOCRD_Destino = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
            If oRs.RecordCount > 0 Then
                sCardCode = oRs.Fields.Item("CardCode").Value.ToString
                If oOCRD_Destino.GetByKey(sCardCode) = True Then
                    oObjGlobal.SBOApp.StatusBar.SetText("Se procede a actualizar el interlocutor " & oOCRD.CardName & " con CIF/NIF " & sLicTradNum, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                    sExiste_IC = True
                Else
                    sExiste_IC = False
                End If
            Else
                sExiste_IC = False
                'oObjGlobal.SBOApp.StatusBar.SetText("No se encuentra con CIF/NIF " & sLicTradNum & " el interlocutor " & oOCRD.CardName & ". Se procede a crearlo.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                oOCRD_Destino.Series = oOCRD.Series
                oOCRD_Destino.CardCode = oOCRD.CardCode
            End If

            oOCRD_Destino.CardName = oOCRD.CardName
            oOCRD_Destino.CardType = oOCRD.CardType
            oOCRD_Destino.CardForeignName = oOCRD.CardForeignName
            oOCRD_Destino.Currency = oOCRD.Currency
#Region "Grupos"
            If sCardType <> "" And sGrupo <> "" And sGrupo <> "0" Then
                oOCRG = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups), SAPbobsCOM.BusinessPartnerGroups)
                'Vemos si existe el grupo
                sSQL = "SELECT * FROM OCRG WHERE ""GroupName""='" & sGrupo & "' and ""GroupType""='" & sCardType & "' "
                oRsGrupos_Des.DoQuery(sSQL)
                If oRsGrupos_Des.RecordCount = 0 Then
                    Select Case sCardType
                        Case "S" : oOCRG.Type = SAPbobsCOM.BoBusinessPartnerGroupTypes.bbpgt_VendorGroup
                        Case "C", "L" : oOCRG.Type = SAPbobsCOM.BoBusinessPartnerGroupTypes.bbpgt_CustomerGroup
                    End Select
                    oOCRG.Name = sGrupo
                    'Añadir
                    If oOCRG.Add() <> 0 Then
                        oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Grupo para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If
                    oCompanyDes.GetNewObjectCode(sGrupo)
                Else
                    sGrupo = oRsGrupos_Des.Fields.Item("GroupCode").Value.ToString
                End If
                oOCRD_Destino.GroupCode = CType(sGrupo, Integer)
            End If
#End Region
            oOCRD_Destino.FederalTaxID = oOCRD.FederalTaxID
            'Pestaña General
            oOCRD_Destino.Phone1 = oOCRD.Phone1
            oOCRD_Destino.Phone2 = oOCRD.Phone2
            oOCRD_Destino.Cellular = oOCRD.Cellular
            oOCRD_Destino.Fax = oOCRD.Fax
            oOCRD_Destino.EmailAddress = oOCRD.EmailAddress
            oOCRD_Destino.MailAddress = oOCRD.MailAddress
            oOCRD_Destino.MailCity = oOCRD.MailCity
            oOCRD_Destino.MailCounty = oOCRD.MailCounty
            oOCRD_Destino.MailZipCode = oOCRD.MailZipCode
            oOCRD_Destino.ETaxWebSite = oOCRD.ETaxWebSite
            oOCRD_Destino.Website = oOCRD.Website
#Region "Clase de Expedición"
            If sClase_Expe <> "" Then
                oOSHP = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes), SAPbobsCOM.ShippingTypes)
                sSQL = "SELECT * FROM OSHP WHERE ""TrnspName""='" & sClase_Expe.Trim & "' "
                oRsClase_Expe.DoQuery(sSQL)
                If oRsClase_Expe.RecordCount > 0 Then
                    sSQL = "SELECT * FROM OSHP WHERE ""TrnspName""='" & sClase_Expe.Trim & "' "
                    oRsClase_Expe_Des.DoQuery(sSQL)
                    If oRsClase_Expe_Des.RecordCount > 0 Then
                        If oOSHP.GetByKey(CType(oRsClase_Expe_Des.Fields.Item("TrnspCode").Value.ToString, Integer)) = True Then
                            oOSHP.Website = oRsClase_Expe.Fields.Item("WebSite").Value.ToString
                            If oOSHP.Update() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Clase de Expedición para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oCompanyDes.GetNewObjectCode(sClase_Expe)
                            End If
                        End If
                    Else
                        oOSHP.Name = sClase_Expe
                        oOSHP.Website = oRsClase_Expe.Fields.Item("WebSite").Value.ToString
                        If oOSHP.Add() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Clase de Expedición para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        Else
                            oCompanyDes.GetNewObjectCode(sClase_Expe)
                        End If
                    End If
                Else
                    oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa activa la clase de expedición " & sClase_Expe, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
                oOCRD_Destino.ShippingType = CType(sClase_Expe, Integer)
            End If
#End Region
            oOCRD_Destino.Password = oOCRD.Password
#Region "Indicador de Factoring"
            If sIndicator <> "" Then
                oOIDC = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFactoringIndicators), SAPbobsCOM.FactoringIndicators)
                sSQL = "SELECT * FROM OIDC WHERE ""Code""='" & sIndicator & "' "
                oRsIndicator.DoQuery(sSQL)
                If oRsIndicator.RecordCount > 0 Then
                    If oOIDC.GetByKey(sIndicator) = True Then
                        oOIDC.IndicatorName = oRsIndicator.Fields.Item("Name").Value.ToString
                        If oOIDC.Update() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Indicador de Factoring para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        End If
                    Else
                        oOIDC.IndicatorCode = sIndicator
                        oOIDC.IndicatorName = oRsIndicator.Fields.Item("Name").Value.ToString
                        If oOIDC.Add() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Indicador de Factoring para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        End If
                    End If
                Else
                    oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa activa el indicador de Factoring " & sIndicator, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
                oOCRD_Destino.Indicator = oOCRD.Indicator
            End If
#End Region
#Region "Ramos"
            If sRamo <> "" And sRamo <> "0" Then
                oOOND = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIndustries), SAPbobsCOM.Industries)
                sSQL = "SELECT * FROM OOND WHERE ""IndCode""='" & sRamo & "' "
                oRsRamo.DoQuery(sSQL)
                If oRsRamo.RecordCount > 0 Then
                    sSQL = "SELECT * FROM OOND WHERE ""IndName""='" & oRsRamo.Fields.Item("IndName").Value.ToString & "' "
                    oRsRamo_Des.DoQuery(sSQL)
                    If oRsRamo_Des.RecordCount > 0 Then
                        oOOND.GetByKey(CType(oRsRamo_Des.Fields.Item("IndCode").Value.ToString, Integer))
                        oOOND.IndustryName = oRsRamo.Fields.Item("IndName").Value.ToString
                        oOOND.IndustryDescription = oRsRamo.Fields.Item("IndDesc").Value.ToString
                        If oOOND.Update() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Ramo para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        Else
                            oCompanyDes.GetNewObjectCode(sRamo)
                        End If
                    Else
                        oOOND.IndustryName = oRsRamo.Fields.Item("IndName").Value.ToString
                        oOOND.IndustryDescription = oRsRamo.Fields.Item("IndDesc").Value.ToString
                        If oOOND.Add() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Ramo para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        Else
                            oCompanyDes.GetNewObjectCode(sRamo)
                        End If
                    End If
                Else
                    oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa activa el ramo " & sRamo, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
                oOCRD_Destino.Industry = CType(sRamo, Integer)
            End If
#End Region
            oOCRD_Destino.CompanyPrivate = oOCRD.CompanyPrivate
            oOCRD_Destino.AliasName = oOCRD.AliasName
            oOCRD_Destino.Valid = oOCRD.Valid
            oOCRD_Destino.ValidFrom = oOCRD.ValidFrom
            oOCRD_Destino.ValidRemarks = oOCRD.ValidRemarks
            oOCRD_Destino.ValidTo = oOCRD.ValidTo

            oOCRD_Destino.AdditionalID = oOCRD.AdditionalID
            oOCRD_Destino.UnifiedFederalTaxID = oOCRD.UnifiedFederalTaxID
            oOCRD_Destino.VATRegistrationNumber = oOCRD.VATRegistrationNumber
            oOCRD_Destino.ResidenNumber = oOCRD.ResidenNumber
            oOCRD_Destino.Notes = oOCRD.Notes
#Region "Medios de comunicación"
            ' No veo como Pasarlo
#End Region
#Region "Empleado de ventas"
            If sEmpleado <> "" Then
                oOSLP = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons), SAPbobsCOM.SalesPersons)
                sSQL = "SELECT * FROM OSLP WHERE ""SlpName""='" & sEmpleado & "' "
                oRsEmpleado.DoQuery(sSQL)
                If oRsEmpleado.RecordCount > 0 Then
                    sSQL = "SELECT * FROM OSLP WHERE ""SlpName""='" & sEmpleado & "' "
                    oRsEmpleado_Des.DoQuery(sSQL)
                    If oRsEmpleado_Des.RecordCount > 0 Then
                        If oOSLP.GetByKey(CType(oRsEmpleado_Des.Fields.Item("SlpCode").Value.ToString, Integer)) = True Then
                            oOSLP.SalesEmployeeName = sEmpleado
                            Select Case oRsEmpleado.Fields.Item("Active").Value.ToString
                                Case "Y" : oOSLP.Active = SAPbobsCOM.BoYesNoEnum.tYES
                                Case Else : oOSLP.Active = SAPbobsCOM.BoYesNoEnum.tNO
                            End Select

                            oOSLP.CommissionForSalesEmployee = EXO_GLOBALES.DblTextToNumber(oCompanyDes, oRsEmpleado.Fields.Item("Commission").Value.ToString)
                            oOSLP.CommissionGroup = CType(oRsEmpleado.Fields.Item("GroupCode").Value.ToString, Integer)
                            If oOSLP.Update() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Empleado de venta para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oCompanyDes.GetNewObjectCode(sEmpleado)
                            End If
                        Else
                            oOSLP.SalesEmployeeName = sEmpleado
                            Select Case oRsEmpleado.Fields.Item("Active").Value.ToString
                                Case "Y" : oOSLP.Active = SAPbobsCOM.BoYesNoEnum.tYES
                                Case Else : oOSLP.Active = SAPbobsCOM.BoYesNoEnum.tNO
                            End Select

                            oOSLP.CommissionForSalesEmployee = EXO_GLOBALES.DblTextToNumber(oCompanyDes, oRsEmpleado.Fields.Item("Commission").Value.ToString)
                            oOSLP.CommissionGroup = CType(oRsEmpleado.Fields.Item("GroupCode").Value.ToString, Integer)
                            If oOSLP.Add() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Empleado de venta para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                                                                oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oCompanyDes.GetNewObjectCode(sEmpleado)
                            End If
                        End If
                    End If
                Else
                    oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa activa el empleado." & sClase_Expe, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
                oOCRD_Destino.SalesPersonCode = CType(sEmpleado, Integer)
            End If
#End Region
#Region "Responsables"
            If sResponsable <> "" Then
                sSQL = "SELECT * FROM OAGP WHERE ""AgentCode"" = '" & sResponsable & "'"
                oRsResponsable_Des.DoQuery(sSQL)
                If oRsResponsable_Des.RecordCount > 0 Then
                    sSQL = "INSERT INTO """ & oCompanyDes.CompanyDB & """.""OAGP"" " &
                               "SELECT ""AgentCode"", ""AgentName"", ""Memo"", ""Locked"", ""DataSource"", ""UserSign"" " &
                               "FROM """ & oObjGlobal.compañia.CompanyDB & """.""OAGP"" t0  " &
                               "WHERE t0.""AgentCode"" = '" & sResponsable & "' "
                Else
                    sSQL = "UPDATE t1 SET ""AgentName"" = t0.""AgentName"", " &
                              """Memo"" = t0.""Memo"", " &
                              """Locked"" = t0.""Locked"", " &
                              """DataSource"" = t0.""DataSource"", " &
                              """UserSign"" = t0.""UserSign"" " &
                              "FROM """ & oObjGlobal.compañia.CompanyDB & """.""OAGP"" t0  INNER JOIN " &
                              """" & oCompanyDes.CompanyDB & """.""OAGP"" t1  ON t0.""AgentCode"" = t1.""AgentCode"" " &
                              "WHERE t0.""AgentCode"" = '" & sResponsable & "' "
                End If
                If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                    oOCRD_Destino.SalesPersonCode = CType(sResponsable, Integer)
                Else
                    oObjGlobal.SBOApp.StatusBar.SetText("Error Responsable para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
            End If
#End Region
#Region "Canal"
            oOCRD_Destino.ChannelBP = oOCRD.ChannelBP
#End Region
            oOCRD_Destino.GlobalLocationNumber = oOCRD.GlobalLocationNumber

            'Pestaña Personas de contacto
#Region "Personas de contacto"
            For i = 0 To oOCRD_Destino.ContactEmployees.Count - 1
                oOCRD_Destino.ContactEmployees.SetCurrentLine(0)
                oOCRD_Destino.ContactEmployees.Delete()
            Next
            iContactos = oOCRD.ContactEmployees.Count
            For i = 0 To iContactos - 1
                oOCRD.ContactEmployees.SetCurrentLine(i)
                oOCRD_Destino.ContactEmployees.Name = oOCRD.ContactEmployees.Name
                oOCRD_Destino.ContactEmployees.Active = oOCRD.ContactEmployees.Active
                oOCRD_Destino.ContactEmployees.FirstName = oOCRD.ContactEmployees.FirstName
                oOCRD_Destino.ContactEmployees.LastName = oOCRD.ContactEmployees.LastName
                oOCRD_Destino.ContactEmployees.MiddleName = oOCRD.ContactEmployees.MiddleName
                oOCRD_Destino.ContactEmployees.Phone1 = oOCRD.ContactEmployees.Phone1
                oOCRD_Destino.ContactEmployees.MobilePhone = oOCRD.ContactEmployees.MobilePhone
                oOCRD_Destino.ContactEmployees.E_Mail = oOCRD.ContactEmployees.E_Mail
                oOCRD_Destino.ContactEmployees.Position = oOCRD.ContactEmployees.Position
                oOCRD_Destino.ContactEmployees.Add()
            Next
            oOCRD_Destino.ContactPerson = oOCRD.ContactPerson
#End Region
            'Pestaña Direcciones
#Region "Direcciones"
            'Eliminamos direcciones
            sSQL = "DELETE FROM """ & oCompanyDes.CompanyDB & """.""CRD1"" Where ""CardCode""='" & sCardCode & "' "
            oObjGlobal.refDi.SQL.executeNonQuery(sSQL)
            For i = 0 To oOCRD_Destino.Addresses.Count - 1
                oOCRD_Destino.Addresses.SetCurrentLine(i)
                oOCRD_Destino.Addresses.Delete()
            Next
            iDirecciones = oOCRD.Addresses.Count
            For i = 0 To iDirecciones - 1
                oOCRD.Addresses.SetCurrentLine(i)
                If oOCRD.Addresses.AddressName <> "" Then
                    oOCRD_Destino.Addresses.AddressType = oOCRD.Addresses.AddressType
                    oOCRD_Destino.Addresses.AddressName = oOCRD.Addresses.AddressName
                    oOCRD_Destino.Addresses.AddressName2 = oOCRD.Addresses.AddressName2
                    oOCRD_Destino.Addresses.AddressName3 = oOCRD.Addresses.AddressName3
                    oOCRD_Destino.Addresses.BuildingFloorRoom = oOCRD.Addresses.BuildingFloorRoom
                    oOCRD_Destino.Addresses.Block = oOCRD.Addresses.Block
                    oOCRD_Destino.Addresses.City = oOCRD.Addresses.City
                    oOCRD_Destino.Addresses.Country = oOCRD.Addresses.Country
                    oOCRD_Destino.Addresses.County = oOCRD.Addresses.County
                    oOCRD_Destino.Addresses.FederalTaxID = oOCRD.Addresses.FederalTaxID
                    oOCRD_Destino.Addresses.GlobalLocationNumber = oOCRD.Addresses.GlobalLocationNumber
                    oOCRD_Destino.Addresses.GSTIN = oOCRD.Addresses.GSTIN
                    oOCRD_Destino.Addresses.GstType = oOCRD.Addresses.GstType
                    oOCRD_Destino.Addresses.MYFType = oOCRD.Addresses.MYFType
                    oOCRD_Destino.Addresses.Nationality = oOCRD.Addresses.Nationality
                    oOCRD_Destino.Addresses.State = oOCRD.Addresses.State
                    oOCRD_Destino.Addresses.Street = oOCRD.Addresses.Street
                    oOCRD_Destino.Addresses.StreetNo = oOCRD.Addresses.StreetNo
                    oOCRD_Destino.Addresses.TaasEnabled = oOCRD.Addresses.TaasEnabled
                    oOCRD_Destino.Addresses.TaxCode = oOCRD.Addresses.TaxCode
                    oOCRD_Destino.Addresses.TaxOffice = oOCRD.Addresses.TaxOffice
                    oOCRD_Destino.Addresses.TypeOfAddress = oOCRD.Addresses.TypeOfAddress
                    oOCRD_Destino.Addresses.ZipCode = oOCRD.Addresses.ZipCode
                    oOCRD_Destino.Addresses.Add()
                End If
            Next
            oOCRD_Destino.ShipToBuildingFloorRoom = oOCRD.ShipToBuildingFloorRoom
            oOCRD_Destino.ShipToDefault = oOCRD.ShipToDefault
            oOCRD_Destino.BilltoDefault = oOCRD.BilltoDefault
#End Region
            'Pestaña condiciones de pago
#Region "Condiciones de pago"
            If sCondPago <> "" Then
                oOCTG = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes), SAPbobsCOM.PaymentTermsTypes)
                sSQL = "SELECT * FROM OCTG WHERE ""PymntGroup""='" & sCondPago & "' "
                oRsCondPago.DoQuery(sSQL)
                If oRsCondPago.RecordCount > 0 Then
                    sSQL = "SELECT * FROM OCTG WHERE ""PymntGroup""='" & sCondPago & "' "
                    oRsCondPago_Des.DoQuery(sSQL)
                    If oRsCondPago_Des.RecordCount > 0 Then
                        oOCTG.GetByKey(CType(oRsCondPago_Des.Fields.Item("GroupNum").Value.ToString, Integer))
                        sExiste_Cond_pago = True

                    Else
                        sExiste_Cond_pago = False
                    End If
                    oOCTG.PaymentTermsGroupName = oRsCondPago.Fields.Item("PymntGroup").Value.ToString
                    Select Case oRsCondPago.Fields.Item("BsLineDate").Value.ToString
                        Case "T" : oOCTG.BaselineDate = SAPbobsCOM.BoBaselineDate.bld_PostingDate
                        Case "S" : oOCTG.BaselineDate = SAPbobsCOM.BoBaselineDate.bld_SystemDate
                        Case "P" : oOCTG.BaselineDate = SAPbobsCOM.BoBaselineDate.bld_TaxDate
                        Case Else : oOCTG.BaselineDate = SAPbobsCOM.BoBaselineDate.bld_ClosingDate
                    End Select
                    oOCTG.CreditLimit = EXO_GLOBALES.DblTextToNumber(oCompanyDes, oRsCondPago.Fields.Item("CredLimit").Value.ToString)
#Region "Dto. PP"
                    sDtoPP = oRsCondPago.Fields.Item("DiscCode").Value.ToString
                    If sDtoPP.Trim <> "" Then
                        sSQL = "SELECT * FROM OCDC WHERE ""Code"" = '" & sDtoPP & "'"
                        oRsdtoPP_Des.DoQuery(sSQL)
                        If oRsdtoPP_Des.RecordCount = 0 Then
                            sSQL = "INSERT INTO """ & oCompanyDes.CompanyDB & """.""OCDC"" " &
                                       "SELECT ""Code"", ""TableDesc"", ""ByDate"", ""Freight"", ""Tax"", ""VatCrctn"",""BaseDate"" " &
                                       "FROM """ & oObjGlobal.compañia.CompanyDB & """.""OCDC"" t0  " &
                                       "WHERE t0.""Code"" = '" & sDtoPP & "'; "

                            sSQL2 = " INSERT INTO """ & oCompanyDes.CompanyDB & """.""CDC1"" " &
                                       "SELECT ""CdcCode"", ""LineId"", ""NumOfDays"", ""Discount"", ""Day"", ""Month""  " &
                                       "FROM """ & oObjGlobal.compañia.CompanyDB & """.""CDC1"" t0  " &
                                       "WHERE t0.""CdcCode"" = '" & sDtoPP & "'; "
                        Else
                            sSQL = "UPDATE t1 SET ""Code"" = t0.""Code"", " &
                                      """TableDesc"" = t0.""TableDesc"", " &
                                      """ByDate"" = t0.""ByDate"", " &
                                      """Freight"" = t0.""Freight"", " &
                                      """Tax"" = t0.""Tax"", " &
                                      """VatCrctn"" = t0.""VatCrctn"", " &
                                       """BaseDate"" = t0.""BaseDate"" " &
                                      "FROM """ & oObjGlobal.compañia.CompanyDB & """.""OCDC"" t0  INNER JOIN " &
                                      """" & oCompanyDes.CompanyDB & """.""OCDC"" t1  ON t0.""Code"" = t1.""Code"" " &
                                      "WHERE t0.""Code"" = '" & sDtoPP & "'; "

                            sSQL2 = " UPDATE t1 SET ""CdcCode"" = t0.""CdcCode"", " &
                                      """LineId"" = t0.""LineId"", " &
                                      """NumOfDays"" = t0.""NumOfDays"", " &
                                      """Discount"" = t0.""Discount"", " &
                                      """Day"" = t0.""Day"", " &
                                      """Month"" = t0.""Month"" " &
                                      "FROM """ & oObjGlobal.compañia.CompanyDB & """.""CDC1"" t0  INNER JOIN " &
                                      """" & oCompanyDes.CompanyDB & """.""CDC1"" t1  ON t0.""CdcCode"" = t1.""CdcCode"" " &
                                      "WHERE t0.""CdcCode"" = '" & sDtoPP & "'; "
                        End If
                        If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                            If oObjGlobal.refDi.SQL.executeNonQuery(sSQL2) = True Then
                                oOCTG.DiscountCode = sDtoPP
                            Else
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Línea Dto. PP para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            End If
                        Else
                            oObjGlobal.SBOApp.StatusBar.SetText("Error cabecera Dto. PP para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        End If
                    End If
#End Region
                    oOCTG.PriceListNo = CType(oRsCondPago.Fields.Item("ListNum").Value.ToString, Integer)
                    oOCTG.CreditLimit = EXO_GLOBALES.DblTextToNumber(oCompanyDes, oRsCondPago.Fields.Item("CredLimit").Value.ToString)
                    oOCTG.LoadLimit = EXO_GLOBALES.DblTextToNumber(oCompanyDes, oRsCondPago.Fields.Item("ObligLimit").Value.ToString)
                    Select Case oRsCondPago.Fields.Item("OpenRcpt").Value.ToString
                        Case "N" : oOCTG.OpenReceipt = SAPbobsCOM.BoOpenIncPayment.oip_No
                        Case "3" : oOCTG.OpenReceipt = SAPbobsCOM.BoOpenIncPayment.oip_Cash
                        Case "1" : oOCTG.OpenReceipt = SAPbobsCOM.BoOpenIncPayment.oip_Checks
                        Case "4" : oOCTG.OpenReceipt = SAPbobsCOM.BoOpenIncPayment.oip_Credit
                        Case "2" : oOCTG.OpenReceipt = SAPbobsCOM.BoOpenIncPayment.oip_BankTransfer
                        Case "5" : oOCTG.OpenReceipt = SAPbobsCOM.BoOpenIncPayment.oip_Cash
                    End Select

                    If sExiste_Cond_pago = True Then
                        If oOCTG.Update() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Actulizando Condiciones de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                                                                oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        Else
                            oCompanyDes.GetNewObjectCode(sGroupNum)
                            ' No se puede añadir, por lo que lo insertamos
                            sInstNum = oRsCondPago.Fields.Item("InstNum").Value.ToString
                            If sInstNum.Trim <> "" Then
                                sSQL = "UPDATE """ & oCompanyDes.CompanyDB & """.""OCTG"" SET ""InstNum"" = " & sInstNum & " " &
                                      "WHERE ""GroupNum"" = " & sGroupNum & "; "
                                sSQL2 = " INSERT INTO """ & oCompanyDes.CompanyDB & """.""CTG1"" " &
                                            "SELECT " & sGroupNum & ", ""IntsNo"", ""InstMonth"", ""InstDays"", ""InstPrcnt"" " &
                                            "FROM """ & oObjGlobal.compañia.CompanyDB & """.""CTG1"" t0  " &
                                            "WHERE t0.""CTGCode"" = " & sGroupNum & "; "
                                If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                    If oObjGlobal.refDi.SQL.executeNonQuery(sSQL2) <> True Then
                                        sSQL2 = "UPDATE D SET  ""InstMonth""=O.""InstMonth"", ""InstDays""= O.""InstDays"", ""InstPrcnt""=O.""InstPrcnt"" "
                                        sSQL2 &= " FROM """ & oObjGlobal.compañia.CompanyDB & """.""CTG1"" O "
                                        sSQL2 &= " INNER JOIN """ & oCompanyDes.CompanyDB & """.""CTG1"" D ON O.""CTGCode""=D.""CTGCode"" And O.""IntsNo""=D.""IntsNo"" "
                                        sSQL2 &= " WHERE O.""CTGCode"" = " & sGroupNum & "; "
                                        If oObjGlobal.refDi.SQL.executeNonQuery(sSQL2) <> True Then
                                            oObjGlobal.SBOApp.StatusBar.SetText("Error Días de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                            Exit Function
                                        End If
                                    End If
                                Else
                                    oObjGlobal.SBOApp.StatusBar.SetText("Error SQL: " & sSQL & " para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If oOCTG.Add() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Condiciones de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                                                                oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        Else
                            oCompanyDes.GetNewObjectCode(sGroupNum)
                            ' No se puede añadir, por lo que lo insertamos
                            sSQL = "UPDATE """ & oCompanyDes.CompanyDB & """.""OCTG"" SET ""InstNum"" = " & sInstNum & " "
                            sSQL &= "WHERE GroupNum = " & sGroupNum & "; "
                            sSQL &= "INSERT INTO """ & oCompanyDes.CompanyDB & """.""CTG1"" "
                            sSQL &= "SELECT " & sGroupNum & ", ""IntsNo"", ""InstMonth"", ""InstDays"", ""InstPrcnt"" "
                            sSQL &= "FROM """ & oObjGlobal.compañia.CompanyDB & """.""CTG1"" t0  "
                            sSQL &= " WHERE t0.""CTGCode"" = " & sGroupNum & "; "
                            If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) <> True Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Días de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            End If
                        End If
                    End If
                    oOCRD_Destino.PayTermsGrpCode = CType(sGroupNum, Integer)
                Else
                    oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa activa la cond. de pago." & sCondPago, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                    Exit Function
                End If
            End If
#End Region
#Region "Días de pago"
            For i = 0 To oOCRD_Destino.BPPaymentDates.Count - 1
                oOCRD_Destino.BPPaymentDates.SetCurrentLine(0)
                oOCRD_Destino.BPPaymentDates.Delete()
            Next
            For i = 0 To oOCRD.BPPaymentDates.Count - 1
                oOCRD.BPPaymentDates.SetCurrentLine(i)
                oOCRD_Destino.BPPaymentDates.PaymentDate = oOCRD.BPPaymentDates.PaymentDate
                oOCRD_Destino.BPPaymentDates.Add()
            Next
#End Region

            oOCRD_Destino.IntrestRatePercent = oOCRD.IntrestRatePercent
            oOCRD_Destino.DiscountPercent = oOCRD.DiscountPercent
            oOCRD_Destino.CreditLimit = oOCRD.CreditLimit
            oOCRD_Destino.DeductionPercent = oOCRD.DeductionPercent
#Region "Plazos de reclamación"
            'Falta ODUT
            oOCRD_Destino.DunningTerm = oOCRD.DunningTerm
#End Region

            oOCRD_Destino.DiscountRelations = oOCRD.DiscountRelations
            oOCRD_Destino.EffectivePrice = oOCRD.EffectivePrice
            oOCRD_Destino.EffectiveDiscount = oOCRD.EffectiveDiscount

            oOCRD_Destino.PartialDelivery = oOCRD.PartialDelivery
            oOCRD_Destino.BackOrder = oOCRD.BackOrder
            oOCRD_Destino.NoDiscounts = oOCRD.NoDiscounts
            oOCRD_Destino.EndorsableChecksFromBP = oOCRD.EndorsableChecksFromBP
            oOCRD_Destino.AcceptsEndorsedChecks = oOCRD.AcceptsEndorsedChecks
#Region "clase de tarjeta Credit"
            'Falta la creación
            oOCRD_Destino.CreditCardCode = oOCRD.CreditCardCode
            oOCRD_Destino.CreditCardExpiration = oOCRD.CreditCardExpiration
            oOCRD_Destino.CreditCardNum = oOCRD.CreditCardNum
            oOCRD_Destino.CreditLimit = oOCRD.CreditLimit
            oOCRD_Destino.AvarageLate = oOCRD.AvarageLate
#End Region
#Region "Prioridad"
            sPrioridad = CType(oOCRD.Priority, String)
            If sPrioridad <> "-1" Then
                oOBPP_Destino = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBPPriorities), SAPbobsCOM.BPPriorities)
                sSQL = "SELECT * FROM OBPP WHERE ""PrioCode""='" & sPrioridad & "' "
                oRsPrioridad.DoQuery(sSQL)
                If oRsPrioridad.RecordCount > 0 Then
                    If oOBPP_Destino.GetByKey(CType(sPrioridad, Integer)) = True Then
                        oOBPP_Destino.PriorityDescription = oRsPrioridad.Fields.Item("PrioDesc").Value.ToString
                        If oOBPP_Destino.Update() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Prioridad para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        Else
                            oCompanyDes.GetNewObjectCode(sPrioridad)
                        End If
                    Else
                        oOBPP_Destino.Priority = CType(sPrioridad, Integer)
                        oOBPP_Destino.PriorityDescription = oRsPrioridad.Fields.Item("PrioDesc").Value.ToString
                        If oOBPP_Destino.Add() <> 0 Then
                            oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Prioridad para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        Else
                            oCompanyDes.GetNewObjectCode(sPrioridad)
                        End If
                    End If
                    oOCRD_Destino.Priority = CType(sPrioridad, Integer)
                End If
            End If
#End Region
            oOCRD_Destino.IBAN = oOCRD.IBAN
#Region "Vacaciones"
            'Falta definir
#End Region
            'Pestaña Ejecución de pago
#Region "Ejecución de pago"
            oOCRD_Destino.HouseBankCountry = oOCRD.HouseBankCountry
            oOCRD_Destino.HouseBank = oOCRD.HouseBank
            oOCRD_Destino.HouseBankBranch = oOCRD.HouseBankBranch
            oOCRD_Destino.HouseBankAccount = oOCRD.HouseBankAccount

            oOCRD_Destino.DME = oOCRD.DME
            oOCRD_Destino.InstructionKey = oOCRD.InstructionKey
            oOCRD_Destino.ReferenceDetails = oOCRD.ReferenceDetails

            oOCRD_Destino.PaymentBlock = oOCRD.PaymentBlock
            oOCRD_Destino.SinglePayment = oOCRD.SinglePayment

            oOCRD_Destino.BankChargesAllocationCode = oOCRD.BankChargesAllocationCode


            oOCRD_Destino.BPPaymentMethods.Delete()
#Region "Vía de pago"
            For i = 0 To oOCRD.BPPaymentMethods.Count - 1
                oOCRD.BPPaymentMethods.SetCurrentLine(i)
                'Comprobamos que exista la vía
                Dim sViaPago As String = oOCRD.BPPaymentMethods.PaymentMethodCode
                sSQL = "Select * FROM OPYM WHERE ""PayMethCod"" = '" & sViaPago & "' "
                oRsOPYM.DoQuery(sSQL)
                oOPYM = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWizardPaymentMethods), SAPbobsCOM.WizardPaymentMethods)
                If sViaPago.Trim <> "" Then
                    If oRsOPYM.RecordCount <> 0 Then
                        If oOPYM.GetByKey(sViaPago) = True Then
                            sExiste_OPYM = True
                        Else
                            sExiste_OPYM = False
                            oOPYM.PaymentMethodCode = sViaPago
                        End If
                        oOPYM.Description = oRsOPYM.Fields.Item("Descript").Value.ToString
                        Select Case oRsOPYM.Fields.Item("Descript").Value.ToString
                            Case "Y" : oOPYM.Active = SAPbobsCOM.BoYesNoEnum.tYES
                            Case Else : oOPYM.Active = SAPbobsCOM.BoYesNoEnum.tNO
                        End Select
                        Select Case oRsOPYM.Fields.Item("Type").Value.ToString
                            Case "I" : oOPYM.Type = SAPbobsCOM.BoPaymentTypeEnum.boptIncoming
                            Case Else : oOPYM.Type = SAPbobsCOM.BoPaymentTypeEnum.boptOutgoing
                        End Select

                        'oOPYM.= oRsOPYM.Fields.Item("BankTransf").Value.ToString
                        oOPYM.BankCountry = oRsOPYM.Fields.Item("BankCountr").Value.ToString
                        oOPYM.DefaultBank = oRsOPYM.Fields.Item("BnkDflt").Value.ToString
                        oOPYM.DefaultAccount = oRsOPYM.Fields.Item("DflAccount").Value.ToString
                        'Porcentaje gastos no lo veo
                        Select Case oRsOPYM.Fields.Item("GrpByDate").Value.ToString
                            Case "Y" : oOPYM.GroupByDate = SAPbobsCOM.BoYesNoEnum.tYES
                            Case Else : oOPYM.GroupByDate = SAPbobsCOM.BoYesNoEnum.tNO
                        End Select

                        If sExiste_OPYM = True Then
                            If oOPYM.Update() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error actualizando Vía de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oCompanyDes.GetNewObjectCode(sViaPago)
                            End If
                        Else
                            If oOPYM.Add() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Vía de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oCompanyDes.GetNewObjectCode(sViaPago)
                            End If
                        End If
                    End If
                End If
                oOCRD_Destino.BPPaymentMethods.PaymentMethodCode = oOCRD.BPPaymentMethods.PaymentMethodCode
                oOCRD_Destino.BPPaymentMethods.Add()
            Next
            oOCRD_Destino.PeymentMethodCode = oOCRD.PeymentMethodCode
#End Region
#End Region
            'Pestaña finanzas
#Region "Pestaña finanzas"
            oOCRD_Destino.FatherCard = oOCRD.FatherCard
            oOCRD_Destino.FatherType = oOCRD.FatherType

            oOCRD_Destino.DownPaymentInterimAccount = oOCRD.DownPaymentInterimAccount
            oOCRD_Destino.DownPaymentClearAct = oOCRD.DownPaymentClearAct
            'Falta una cuenta y las del botón

            'Falta connbp

            oOCRD_Destino.PlanningGroup = oOCRD.PlanningGroup
            oOCRD_Destino.Affiliate = oOCRD.Affiliate

            'Pestaña de impuesto
            oOCRD_Destino.Equalization = oOCRD.Equalization
            oOCRD_Destino.VatIDNum = oOCRD.VatIDNum
            oOCRD_Destino.ECommerceMerchantID = oOCRD.ECommerceMerchantID
            oOCRD_Destino.AccrualCriteria = oOCRD.AccrualCriteria
            oOCRD_Destino.CertificateNumber = oOCRD.CertificateNumber
            oOCRD_Destino.ExpirationDate = oOCRD.ExpirationDate

            oOCRD_Destino.OperationCode347 = oOCRD.OperationCode347
            oOCRD_Destino.InsuranceOperation347 = oOCRD.InsuranceOperation347
#End Region
#Region "Pestaña propiedades"
            For i = 1 To 64
                oOCRD_Destino.Properties(i) = oOCRD.Properties(i)
            Next
#End Region
            oOCRD_Destino.FreeText = oOCRD.FreeText
#Region "Documentos electrónicos"
            oOCRD_Destino.EDocGenerationType = oOCRD.EDocGenerationType
            oOCRD_Destino.FCERelevant = oOCRD.FCERelevant
            oOCRD_Destino.FCEValidateBaseDelivery = oOCRD.FCEValidateBaseDelivery
#End Region
#Region "Campos de usuario"
            sSQL = "select ""AliasID"" FROM """ & oCompanyDes.CompanyDB & """.""CUFD"" WHERE ""TableID"" = 'OCRD';"
            oRsCamposUsuario.DoQuery(sSQL)
            For i = 0 To oRsCamposUsuario.RecordCount - 1
                Try
                    Dim sCampo As String = "U_" & oRsCamposUsuario.Fields.Item("AliasID").Value.ToString
                    oOCRD_Destino.UserFields.Fields.Item(sCampo).Value = oOCRD.UserFields.Fields.Item(sCampo).Value
                Catch ex As Exception

                End Try
                oRsCamposUsuario.MoveNext()
            Next
#End Region

            If sExiste_IC = False Then
                'If oOCRD_Destino.Add() <> 0 Then
                '    oObjGlobal.SBOApp.StatusBar.SetText("Error Creando IC - " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                '                                                oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                'Else
                '    oObjGlobal.SBOApp.StatusBar.SetText("IC Creado- " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                'End If
            Else
                If oOCRD_Destino.Update() <> 0 Then
                    oObjGlobal.SBOApp.StatusBar.SetText("Error actualizando IC - " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                                                                oCompanyDes.GetLastErrorCode & " / " & oCompanyDes.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                Else
                    oObjGlobal.SBOApp.StatusBar.SetText("IC Actualizado - " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                End If
            End If
            Sincroniza_proveedor_Master = True
        Catch ex As Exception
            Throw ex
        Finally
#Region "Liberar"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRD_Destino, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsGrupos_Des, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsClase_Expe, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsClase_Expe_Des, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOSHP, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsIndicator, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsRamo, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsRamo_Des, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsEmpleado, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsResponsable_Des, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCondPago, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCondPago_Des, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsdtoPP_Des, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOBPP_Destino, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCDC, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRD, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRG, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCTG, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOIDC, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOOND, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOPYM, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOSHP, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOSLP, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsOPYM, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCamposUsuario, Object))
#End Region
        End Try
    End Function
    Public Shared Function Sincroniza_proveedor(ByRef oOCRD As SAPbobsCOM.BusinessPartners, ByRef oCompanyMaster As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
#Region "Variables"
        Dim oOCRD_Master As SAPbobsCOM.BusinessPartners = Nothing
        Dim sLicTradNum As String = "" : Dim sCardCode As String = "" : Dim sCardType As String = ""
        Dim sSQL As String = "" : Dim sSQL2 As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sGrupo As String = "" : Dim oRsGrupos_Des As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOCRG As SAPbobsCOM.BusinessPartnerGroups = Nothing : Dim oOCRG_Master As SAPbobsCOM.BusinessPartnerGroups = Nothing
        Dim sClase_Expe As String = "" : Dim oOSHP As SAPbobsCOM.ShippingTypes = Nothing : Dim oOSHP_Master As SAPbobsCOM.ShippingTypes = Nothing
        Dim oRsClase_Expe_Des As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sIndicator As String = "" : Dim oOIDC As SAPbobsCOM.FactoringIndicators = Nothing : Dim oOIDC_Master As SAPbobsCOM.FactoringIndicators = Nothing
        Dim sRamo As String = "" : Dim oOOND As SAPbobsCOM.Industries = Nothing : Dim oOOND_Master As SAPbobsCOM.Industries = Nothing
        Dim oRsRamo As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sEmpleado As String = "" : Dim oOSLP As SAPbobsCOM.SalesPersons = Nothing : Dim oOSLP_Master As SAPbobsCOM.SalesPersons = Nothing
        Dim oRsEmpleado As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sResponsable As String = ""
        Dim oRsResponsable As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim iContactos As Integer = 0 : Dim iDirecciones As Integer = 0
        Dim sCondPago As String = "" : Dim sExiste_Cond_pago As Boolean = False
        Dim oOCTG As SAPbobsCOM.PaymentTermsTypes = Nothing : Dim oOCTG_Master As SAPbobsCOM.PaymentTermsTypes = Nothing
        Dim oRsCondPago As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sDtoPP As String = "" : Dim oRsdtoPP As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sGroupNum As String = "" : Dim sInstNum As String = ""
        Dim sPrioridad As String = "" : Dim oRsPrioridad As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOBPP As SAPbobsCOM.BPPriorities = Nothing : Dim oOBPP_Master As SAPbobsCOM.BPPriorities = Nothing
        Dim oRsOPYM As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oOPYM As SAPbobsCOM.WizardPaymentMethods = Nothing : Dim oOPYM_Master As SAPbobsCOM.WizardPaymentMethods = Nothing : Dim sExiste_OPYM As Boolean = False
        Dim oRsCamposUsuario As SAPbobsCOM.Recordset = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
#End Region

        Sincroniza_proveedor = False
        Try
            'Primero buscamos si existe el IC con el NIF
            sLicTradNum = oOCRD.FederalTaxID
            Select Case oOCRD.CardType
                Case SAPbobsCOM.BoCardTypes.cSupplier : sCardType = "S"
                Case SAPbobsCOM.BoCardTypes.cLid : sCardType = "L"
                Case SAPbobsCOM.BoCardTypes.cCustomer : sCardType = "C"
            End Select
            '            sGroupNum = oOCRD.PayTermsGrpCode
            oRs = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = "SELECT ""CardCode"" FROM OCRD WHERE ""LicTradNum""='" & sLicTradNum & "' and ""CardType""='" & sCardType & "' "
            oRs.DoQuery(sSQL)
            oOCRD_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
            If oRs.RecordCount > 0 Then
                sCardCode = oRs.Fields.Item("CardCode").Value.ToString
                If oOCRD_Master.GetByKey(sCardCode) = True Then
                    oObjGlobal.SBOApp.StatusBar.SetText("Se procede a actualizar el interlocutor " & oOCRD.CardName & " con CIF/NIF " & sLicTradNum, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning)
                End If

                oOCRD.CardName = oOCRD_Master.CardName
                oOCRD.CardForeignName = oOCRD_Master.CardForeignName
                oOCRD.Currency = oOCRD_Master.Currency
#Region "Grupos"
                oOCRG_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups), SAPbobsCOM.BusinessPartnerGroups)
                If oOCRG_Master.GetByKey(oOCRD_Master.GroupCode) = True Then
                    sGrupo = oOCRG_Master.Name
                    If sCardType <> "" And sGrupo <> "" And sGrupo <> "0" Then
                        oOCRG = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups), SAPbobsCOM.BusinessPartnerGroups)
                        'Vemos si existe el grupo
                        sSQL = "SELECT * FROM OCRG WHERE ""GroupName""='" & sGrupo & "' and ""GroupType""='" & sCardType & "' "
                        oRsGrupos_Des.DoQuery(sSQL)
                        If oRsGrupos_Des.RecordCount = 0 Then
                            Select Case sCardType
                                Case "S" : oOCRG.Type = SAPbobsCOM.BoBusinessPartnerGroupTypes.bbpgt_VendorGroup
                                Case "C", "L" : oOCRG.Type = SAPbobsCOM.BoBusinessPartnerGroupTypes.bbpgt_CustomerGroup
                            End Select
                            oOCRG.Name = sGrupo
                            'Añadir
                            If oOCRG.Add() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Grupo para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            End If
                            oObjGlobal.compañia.GetNewObjectCode(sGrupo)
                        Else
                            sGrupo = oRsGrupos_Des.Fields.Item("GroupCode").Value.ToString
                        End If
                        oOCRD.GroupCode = CType(sGrupo, Integer)
                    End If
                End If
#End Region
                oOCRD.FederalTaxID = oOCRD_Master.FederalTaxID
                'Pestaña General
                oOCRD.Phone1 = oOCRD_Master.Phone1
                oOCRD.Phone2 = oOCRD_Master.Phone2
                oOCRD.Cellular = oOCRD_Master.Cellular
                oOCRD.Fax = oOCRD_Master.Fax
                oOCRD.EmailAddress = oOCRD_Master.EmailAddress
                oOCRD.MailAddress = oOCRD_Master.MailAddress
                oOCRD.MailCity = oOCRD_Master.MailCity
                oOCRD.MailCounty = oOCRD_Master.MailCounty
                oOCRD.MailZipCode = oOCRD_Master.MailZipCode
                oOCRD.ETaxWebSite = oOCRD_Master.ETaxWebSite
                oOCRD.Website = oOCRD_Master.Website
#Region "Clase de Expedición"
                oOSHP_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes), SAPbobsCOM.ShippingTypes)
                If oOSHP_Master.GetByKey(oOCRD_Master.ShippingType) = True Then
                    sClase_Expe = oOSHP_Master.Name

                    If sClase_Expe <> "" Then
                        oOSHP = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes), SAPbobsCOM.ShippingTypes)

                        sSQL = "SELECT * FROM OSHP WHERE ""TrnspName""='" & sClase_Expe.Trim & "' "
                        oRsClase_Expe_Des.DoQuery(sSQL)
                        If oRsClase_Expe_Des.RecordCount > 0 Then
                            If oOSHP.GetByKey(CType(oRsClase_Expe_Des.Fields.Item("TrnspCode").Value.ToString, Integer)) = True Then
                                oOSHP.Website = oOSHP_Master.Website
                                If oOSHP.Update() <> 0 Then
                                    oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Clase de Expedición para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    Exit Function
                                Else
                                    oObjGlobal.compañia.GetNewObjectCode(sClase_Expe)
                                End If
                            End If
                        Else
                            oOSHP.Name = sClase_Expe
                            oOSHP.Website = oOSHP_Master.Website
                            If oOSHP.Add() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Clase de Expedición para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oObjGlobal.compañia.GetNewObjectCode(sClase_Expe)
                            End If
                        End If
                        oOCRD.ShippingType = CType(sClase_Expe, Integer)
                    End If
                End If
#End Region
                oOCRD.Password = oOCRD_Master.Password
#Region "Indicador de Factoring"
                sIndicator = oOCRD_Master.Indicator
                If sIndicator <> "" Then
                    oOIDC_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFactoringIndicators), SAPbobsCOM.FactoringIndicators)
                    If oOIDC_Master.GetByKey(sIndicator) = True Then
                        oOIDC = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFactoringIndicators), SAPbobsCOM.FactoringIndicators)
                        If oOIDC.GetByKey(sIndicator) = True Then
                            oOIDC.IndicatorName = oOIDC_Master.IndicatorName
                            If oOIDC.Update() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Indicador de Factoring para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            End If
                        Else
                            oOIDC.IndicatorCode = sIndicator
                            oOIDC.IndicatorName = oOIDC_Master.IndicatorName
                            If oOIDC.Add() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Indicador de Factoring para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            End If
                        End If
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa Master el indicador de Factoring " & sIndicator, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If
                    oOCRD.Indicator = oOCRD_Master.Indicator
                End If
#End Region
#Region "Ramos"
                sRamo = CType(oOCRD_Master.Industry, String)
                If sRamo <> "" And sRamo <> "0" Then
                    oOOND_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIndustries), SAPbobsCOM.Industries)
                    If oOOND_Master.GetByKey(CType(sRamo, Integer)) = True Then
                        sSQL = "SELECT * FROM OOND WHERE ""IndName""='" & oOOND_Master.IndustryName & "' "
                        oRsRamo.DoQuery(sSQL)
                        oOOND = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIndustries), SAPbobsCOM.Industries)
                        If oRsRamo.RecordCount > 0 Then
                            oOOND.GetByKey(CType(oRsRamo.Fields.Item("IndCode").Value.ToString, Integer))
                            oOOND.IndustryName = oOOND_Master.IndustryName
                            oOOND.IndustryDescription = oOOND_Master.IndustryDescription
                            If oOOND.Update() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Ramo para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oObjGlobal.compañia.GetNewObjectCode(sRamo)
                            End If
                        Else
                            oOOND.IndustryName = oOOND_Master.IndustryName
                            oOOND.IndustryDescription = oOOND_Master.IndustryDescription
                            If oOOND.Add() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Ramo para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oObjGlobal.compañia.GetNewObjectCode(sRamo)
                            End If
                        End If
                        oOCRD.Industry = CType(sRamo, Integer)
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa Master el Ramo " & sRamo, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If
                End If
#End Region
                oOCRD.CompanyPrivate = oOCRD_Master.CompanyPrivate
                oOCRD.AliasName = oOCRD_Master.AliasName
                oOCRD.Valid = oOCRD_Master.Valid
                oOCRD.ValidFrom = oOCRD_Master.ValidFrom
                oOCRD.ValidRemarks = oOCRD_Master.ValidRemarks
                oOCRD.ValidTo = oOCRD_Master.ValidTo

                oOCRD.AdditionalID = oOCRD_Master.AdditionalID
                oOCRD.UnifiedFederalTaxID = oOCRD_Master.UnifiedFederalTaxID
                oOCRD.VATRegistrationNumber = oOCRD_Master.VATRegistrationNumber
                oOCRD.ResidenNumber = oOCRD_Master.ResidenNumber
                oOCRD.Notes = oOCRD_Master.Notes
#Region "Medios de comunicación"
                ' No veo como Pasarlo
#End Region
#Region "Empleado de ventas"
                sEmpleado = CType(oOCRD_Master.SalesPersonCode, String)
                If sEmpleado <> "" Then
                    oOSLP_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons), SAPbobsCOM.SalesPersons)
                    If oOSLP_Master.GetByKey(CType(sEmpleado, Integer)) = True Then
                        sSQL = "SELECT * FROM OSLP WHERE ""SlpName""='" & oOSLP_Master.SalesEmployeeName & "' "
                        oRsEmpleado.DoQuery(sSQL)
                        oOSLP = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons), SAPbobsCOM.SalesPersons)
                        If oRsEmpleado.RecordCount > 0 Then
                            oOSLP.GetByKey(CType(oRsEmpleado.Fields.Item("SlpCode").Value.ToString, Integer))
                            oOSLP.SalesEmployeeName = oOSLP_Master.SalesEmployeeName
                            oOSLP.Active = oOSLP_Master.Active
                            oOSLP.CommissionForSalesEmployee = oOSLP_Master.CommissionForSalesEmployee
                            oOSLP.CommissionGroup = oOSLP_Master.CommissionGroup
                            If oOSLP.Update() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Empleado de venta para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oObjGlobal.compañia.GetNewObjectCode(sEmpleado)
                            End If
                        Else
                            oOSLP.SalesEmployeeName = oOSLP_Master.SalesEmployeeName
                            oOSLP.Active = oOSLP_Master.Active
                            oOSLP.CommissionForSalesEmployee = oOSLP_Master.CommissionForSalesEmployee
                            oOSLP.CommissionGroup = oOSLP_Master.CommissionGroup
                            If oOSLP.Add() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Empleado de venta para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                                                                oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oObjGlobal.compañia.GetNewObjectCode(sEmpleado)
                            End If
                        End If
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa Master el empleado de ventas " & sEmpleado, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If
                    oOCRD.SalesPersonCode = CType(sEmpleado, Integer)
                End If
#End Region
#Region "Responsables"
                sResponsable = oOCRD_Master.AgentCode
                If sResponsable <> "" Then
                    sSQL = "SELECT * FROM """ & oObjGlobal.compañia.CompanyDB & """.""OAGP"" WHERE ""AgentCode"" = '" & sResponsable & "'"
                    oRsResponsable.DoQuery(sSQL)
                    If oRsResponsable.RecordCount > 0 Then
                        sSQL = "UPDATE t1 SET ""AgentName"" = t0.""AgentName"", " &
                                  """Memo"" = t0.""Memo"", " &
                                  """Locked"" = t0.""Locked"", " &
                                  """DataSource"" = t0.""DataSource"", " &
                                  """UserSign"" = t0.""UserSign"" " &
                                  "FROM """ & oCompanyMaster.CompanyDB & """.""OAGP"" t0  INNER JOIN " &
                                  """" & oObjGlobal.compañia.CompanyDB & """.""OAGP"" t1  ON t0.""AgentCode"" = t1.""AgentCode"" " &
                                  "WHERE t0.""AgentCode"" = '" & sResponsable & "' "

                    Else
                        sSQL = "INSERT INTO """ & oObjGlobal.compañia.CompanyDB & """.""OAGP"" " &
                                   "SELECT ""AgentCode"", ""AgentName"", ""Memo"", ""Locked"", ""DataSource"", ""UserSign"" " &
                                   "FROM """ & oCompanyMaster.CompanyDB & """.""OAGP"" t0  " &
                                   "WHERE t0.""AgentCode"" = '" & sResponsable & "' "
                    End If
                    If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                        oOCRD.SalesPersonCode = CType(sResponsable, Integer)
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("Error Responsable para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If
                End If
#End Region
#Region "Canal"
                oOCRD.ChannelBP = oOCRD_Master.ChannelBP
#End Region
                oOCRD.GlobalLocationNumber = oOCRD_Master.GlobalLocationNumber

                'Pestaña Personas de contacto
#Region "Personas de contacto"
                For i = 0 To oOCRD.ContactEmployees.Count - 1
                    oOCRD.ContactEmployees.SetCurrentLine(0)
                    oOCRD.ContactEmployees.Delete()
                Next

                iContactos = oOCRD_Master.ContactEmployees.Count
                For i = 0 To iContactos - 1
                    oOCRD_Master.ContactEmployees.SetCurrentLine(i)
                    oOCRD.ContactEmployees.Name = oOCRD_Master.ContactEmployees.Name
                    oOCRD.ContactEmployees.Active = oOCRD_Master.ContactEmployees.Active
                    oOCRD.ContactEmployees.FirstName = oOCRD_Master.ContactEmployees.FirstName
                    oOCRD.ContactEmployees.LastName = oOCRD_Master.ContactEmployees.LastName
                    oOCRD.ContactEmployees.MiddleName = oOCRD_Master.ContactEmployees.MiddleName
                    oOCRD.ContactEmployees.Phone1 = oOCRD_Master.ContactEmployees.Phone1
                    oOCRD.ContactEmployees.MobilePhone = oOCRD_Master.ContactEmployees.MobilePhone
                    oOCRD.ContactEmployees.E_Mail = oOCRD_Master.ContactEmployees.E_Mail
                    oOCRD.ContactEmployees.Position = oOCRD_Master.ContactEmployees.Position
                    oOCRD.ContactEmployees.Add()
                Next
                oOCRD.ContactPerson = oOCRD_Master.ContactPerson
#End Region
                'Pestaña Direcciones
#Region "Direcciones"
                'Eliminamos direcciones
                sSQL = "DELETE FROM """ & oObjGlobal.compañia.CompanyDB & """.""CRD1"" Where ""CardCode""='" & sCardCode & "' "
                oObjGlobal.refDi.SQL.executeNonQuery(sSQL)
                For i = 0 To oOCRD.Addresses.Count - 1
                    oOCRD.Addresses.SetCurrentLine(0)
                    oOCRD.Addresses.Delete()
                Next

                iDirecciones = oOCRD_Master.Addresses.Count
                For i = 0 To iDirecciones - 1
                    oOCRD_Master.Addresses.SetCurrentLine(i)
                    If oOCRD_Master.Addresses.AddressName <> "" Then
                        oOCRD.Addresses.AddressType = oOCRD_Master.Addresses.AddressType
                        oOCRD.Addresses.AddressName = oOCRD_Master.Addresses.AddressName
                        oOCRD.Addresses.AddressName2 = oOCRD_Master.Addresses.AddressName2
                        oOCRD.Addresses.AddressName3 = oOCRD_Master.Addresses.AddressName3
                        oOCRD.Addresses.BuildingFloorRoom = oOCRD_Master.Addresses.BuildingFloorRoom
                        oOCRD.Addresses.Block = oOCRD_Master.Addresses.Block
                        oOCRD.Addresses.City = oOCRD_Master.Addresses.City
                        oOCRD.Addresses.Country = oOCRD_Master.Addresses.Country
                        oOCRD.Addresses.County = oOCRD_Master.Addresses.County
                        oOCRD.Addresses.FederalTaxID = oOCRD_Master.Addresses.FederalTaxID
                        oOCRD.Addresses.GlobalLocationNumber = oOCRD_Master.Addresses.GlobalLocationNumber
                        oOCRD.Addresses.GSTIN = oOCRD_Master.Addresses.GSTIN
                        oOCRD.Addresses.GstType = oOCRD_Master.Addresses.GstType
                        oOCRD.Addresses.MYFType = oOCRD_Master.Addresses.MYFType
                        oOCRD.Addresses.Nationality = oOCRD_Master.Addresses.Nationality
                        oOCRD.Addresses.State = oOCRD_Master.Addresses.State
                        oOCRD.Addresses.Street = oOCRD_Master.Addresses.Street
                        oOCRD.Addresses.StreetNo = oOCRD_Master.Addresses.StreetNo
                        oOCRD.Addresses.TaasEnabled = oOCRD_Master.Addresses.TaasEnabled
                        oOCRD.Addresses.TaxCode = oOCRD_Master.Addresses.TaxCode
                        oOCRD.Addresses.TaxOffice = oOCRD_Master.Addresses.TaxOffice
                        oOCRD.Addresses.TypeOfAddress = oOCRD_Master.Addresses.TypeOfAddress
                        oOCRD.Addresses.ZipCode = oOCRD_Master.Addresses.ZipCode
                        oOCRD.Addresses.Add()
                    End If
                Next
                oOCRD.ShipToBuildingFloorRoom = oOCRD_Master.ShipToBuildingFloorRoom
                oOCRD.ShipToDefault = oOCRD_Master.ShipToDefault
                oOCRD.BilltoDefault = oOCRD_Master.BilltoDefault
#End Region
                'Pestaña condiciones de pago
#Region "Condiciones de pago"
                sCondPago = CType(oOCRD_Master.PayTermsGrpCode, String)
                If sCondPago <> "" Then
                    oOCTG_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes), SAPbobsCOM.PaymentTermsTypes)
                    If oOCTG_Master.GetByKey(CType(sCondPago, Integer)) = True Then
                        oOCTG = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes), SAPbobsCOM.PaymentTermsTypes)
                        sSQL = "SELECT * FROM OCTG WHERE ""PymntGroup""='" & oOCTG_Master.PaymentTermsGroupName & "' "
                        oRsCondPago.DoQuery(sSQL)
                        If oRsCondPago.RecordCount > 0 Then
                            oOCTG.GetByKey(CType(oRsCondPago.Fields.Item("GroupNum").Value.ToString, Integer))
                            sExiste_Cond_pago = True
                        Else
                            sExiste_Cond_pago = False
                        End If
                        oOCTG.PaymentTermsGroupName = oOCTG_Master.PaymentTermsGroupName
                        oOCTG.BaselineDate = oOCTG_Master.BaselineDate
                        oOCTG.CreditLimit = oOCTG_Master.CreditLimit
#Region "Dto. PP"
                        sDtoPP = oOCTG_Master.DiscountCode
                        If sDtoPP.Trim <> "" Then
                            sSQL = "SELECT * FROM OCDC WHERE ""Code"" = '" & sDtoPP & "'"
                            oRsdtoPP.DoQuery(sSQL)
                            If oRsdtoPP.RecordCount = 0 Then
                                sSQL = "INSERT INTO """ & oObjGlobal.compañia.CompanyDB & """.""OCDC"" " &
                                           "SELECT ""Code"", ""TableDesc"", ""ByDate"", ""Freight"", ""Tax"", ""VatCrctn"",""BaseDate"" " &
                                           "FROM """ & oCompanyMaster.CompanyDB & """.""OCDC"" t0  " &
                                           "WHERE t0.""Code"" = '" & sDtoPP & "'; "

                                sSQL2 = " INSERT INTO """ & oObjGlobal.compañia.CompanyDB & """.""CDC1"" " &
                                           "SELECT ""CdcCode"", ""LineId"", ""NumOfDays"", ""Discount"", ""Day"", ""Month""  " &
                                           "FROM """ & oObjGlobal.compañia.CompanyDB & """.""CDC1"" t0  " &
                                           "WHERE t0.""CdcCode"" = '" & sDtoPP & "'; "
                            Else
                                sSQL = "UPDATE t1 SET ""Code"" = t0.""Code"", " &
                                          """TableDesc"" = t0.""TableDesc"", " &
                                          """ByDate"" = t0.""ByDate"", " &
                                          """Freight"" = t0.""Freight"", " &
                                          """Tax"" = t0.""Tax"", " &
                                          """VatCrctn"" = t0.""VatCrctn"", " &
                                           """BaseDate"" = t0.""BaseDate"" " &
                                          "FROM """ & oCompanyMaster.CompanyDB & """.""OCDC"" t0  INNER JOIN " &
                                          """" & oObjGlobal.compañia.CompanyDB & """.""OCDC"" t1  ON t0.""Code"" = t1.""Code"" " &
                                          "WHERE t0.""Code"" = '" & sDtoPP & "'; "

                                sSQL2 = " UPDATE t1 SET ""CdcCode"" = t0.""CdcCode"", " &
                                          """LineId"" = t0.""LineId"", " &
                                          """NumOfDays"" = t0.""NumOfDays"", " &
                                          """Discount"" = t0.""Discount"", " &
                                          """Day"" = t0.""Day"", " &
                                          """Month"" = t0.""Month"" " &
                                          "FROM """ & oCompanyMaster.CompanyDB & """.""CDC1"" t0  INNER JOIN " &
                                          """" & oObjGlobal.compañia.CompanyDB & """.""CDC1"" t1  ON t0.""CdcCode"" = t1.""CdcCode"" " &
                                          "WHERE t0.""CdcCode"" = '" & sDtoPP & "'; "
                            End If
                            If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                If oObjGlobal.refDi.SQL.executeNonQuery(sSQL2) = True Then
                                    oOCTG.DiscountCode = sDtoPP
                                Else
                                    oObjGlobal.SBOApp.StatusBar.SetText("Error Línea Dto. PP para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    Exit Function
                                End If
                            Else
                                oObjGlobal.SBOApp.StatusBar.SetText("Error cabecera Dto. PP para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            End If
                        End If
#End Region
                        oOCTG.PriceListNo = oOCTG_Master.PriceListNo
                        oOCTG.CreditLimit = oOCTG_Master.CreditLimit
                        oOCTG.LoadLimit = oOCTG_Master.LoadLimit
                        oOCTG.OpenReceipt = oOCTG_Master.OpenReceipt

                        If sExiste_Cond_pago = True Then
                            If oOCTG.Update() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Actulizando Condiciones de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                                                                    oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oObjGlobal.compañia.GetNewObjectCode(sGroupNum)
                                ' No se puede añadir, por lo que lo insertamos
                                sInstNum = oRsCondPago.Fields.Item("InstNum").Value.ToString
                                If sInstNum.Trim <> "" Then
                                    sSQL = "UPDATE """ & oObjGlobal.compañia.CompanyDB & """.""OCTG"" SET ""InstNum"" = " & sInstNum & " " &
                                          "WHERE ""GroupNum"" = " & sGroupNum & "; "
                                    sSQL2 = " INSERT INTO """ & oObjGlobal.compañia.CompanyDB & """.""CTG1"" " &
                                                "SELECT " & sGroupNum & ", ""IntsNo"", ""InstMonth"", ""InstDays"", ""InstPrcnt"" " &
                                                "FROM """ & oCompanyMaster.CompanyDB & """.""CTG1"" t0  " &
                                                "WHERE t0.""CTGCode"" = " & sGroupNum & "; "
                                    If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                        If oObjGlobal.refDi.SQL.executeNonQuery(sSQL2) <> True Then
                                            sSQL2 = "UPDATE D SET  ""InstMonth""=O.""InstMonth"", ""InstDays""= O.""InstDays"", ""InstPrcnt""=O.""InstPrcnt"" "
                                            sSQL2 &= " FROM """ & oCompanyMaster.CompanyDB & """.""CTG1"" O "
                                            sSQL2 &= " INNER JOIN """ & oObjGlobal.compañia.CompanyDB & """.""CTG1"" D ON O.""CTGCode""=D.""CTGCode"" And O.""IntsNo""=D.""IntsNo"" "
                                            sSQL2 &= " WHERE O.""CTGCode"" = " & sGroupNum & "; "
                                            If oObjGlobal.refDi.SQL.executeNonQuery(sSQL2) <> True Then
                                                oObjGlobal.SBOApp.StatusBar.SetText("Error Días de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                                Exit Function
                                            End If
                                        End If
                                    Else
                                        oObjGlobal.SBOApp.StatusBar.SetText("Error SQL: " & sSQL & " para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                        Exit Function
                                    End If
                                End If
                            End If
                        Else
                            If oOCTG.Add() <> 0 Then
                                oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Condiciones de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                                                                    oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                Exit Function
                            Else
                                oObjGlobal.compañia.GetNewObjectCode(sGroupNum)
                                ' No se puede añadir, por lo que lo insertamos
                                sSQL = "UPDATE """ & oObjGlobal.compañia.CompanyDB & """.""OCTG"" SET ""InstNum"" = " & sInstNum & " "
                                sSQL &= "WHERE GroupNum = " & sGroupNum & "; "
                                sSQL &= "INSERT INTO """ & oObjGlobal.compañia.CompanyDB & """.""CTG1"" "
                                sSQL &= "SELECT " & sGroupNum & ", ""IntsNo"", ""InstMonth"", ""InstDays"", ""InstPrcnt"" "
                                sSQL &= "FROM """ & oCompanyMaster.CompanyDB & """.""CTG1"" t0  "
                                sSQL &= " WHERE t0.""CTGCode"" = " & sGroupNum & "; "
                                If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) <> True Then
                                    oObjGlobal.SBOApp.StatusBar.SetText("Error Días de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    Exit Function
                                End If
                            End If
                        End If
                        oOCRD.PayTermsGrpCode = CType(sGroupNum, Integer)
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa Master la condición de pago." & sCondPago, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If
                End If
#End Region
#Region "Días de pago"
                For i = 0 To oOCRD.BPPaymentDates.Count - 1
                    oOCRD.BPPaymentDates.SetCurrentLine(0)
                    oOCRD.BPPaymentDates.Delete()
                Next
                For i = 0 To oOCRD_Master.BPPaymentDates.Count - 1
                    oOCRD_Master.BPPaymentDates.SetCurrentLine(i)
                    oOCRD.BPPaymentDates.PaymentDate = oOCRD_Master.BPPaymentDates.PaymentDate
                    oOCRD.BPPaymentDates.Add()
                Next
#End Region

                oOCRD.IntrestRatePercent = oOCRD_Master.IntrestRatePercent
                oOCRD.DiscountPercent = oOCRD_Master.DiscountPercent
                oOCRD.CreditLimit = oOCRD_Master.CreditLimit
                oOCRD.DeductionPercent = oOCRD_Master.DeductionPercent
#Region "Plazos de reclamación"
                'Falta ODUT
                oOCRD.DunningTerm = oOCRD_Master.DunningTerm
#End Region

                oOCRD.DiscountRelations = oOCRD_Master.DiscountRelations
                oOCRD.EffectivePrice = oOCRD_Master.EffectivePrice
                oOCRD.EffectiveDiscount = oOCRD_Master.EffectiveDiscount

                oOCRD.PartialDelivery = oOCRD_Master.PartialDelivery
                oOCRD.BackOrder = oOCRD_Master.BackOrder
                oOCRD.NoDiscounts = oOCRD_Master.NoDiscounts
                oOCRD.EndorsableChecksFromBP = oOCRD_Master.EndorsableChecksFromBP
                oOCRD.AcceptsEndorsedChecks = oOCRD_Master.AcceptsEndorsedChecks
#Region "clase de tarjeta Credit"
                'Falta la creación
                oOCRD.CreditCardCode = oOCRD_Master.CreditCardCode
                oOCRD.CreditCardExpiration = oOCRD_Master.CreditCardExpiration
                oOCRD.CreditCardNum = oOCRD_Master.CreditCardNum
                oOCRD.CreditLimit = oOCRD_Master.CreditLimit
                oOCRD.AvarageLate = oOCRD_Master.AvarageLate
#End Region
#Region "Prioridad"
                sPrioridad = CType(oOCRD_Master.Priority, String)
                If sPrioridad <> "-1" Then
                    oOBPP_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBPPriorities), SAPbobsCOM.BPPriorities)
                    If oOBPP_Master.GetByKey(CType(sPrioridad, Integer)) = True Then
                        sSQL = "SELECT * FROM OBPP WHERE ""PrioCode""='" & sPrioridad & "' "
                        oRsPrioridad.DoQuery(sSQL)
                        oOBPP = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBPPriorities), SAPbobsCOM.BPPriorities)
                        If oRsPrioridad.RecordCount > 0 Then
                            If oOBPP.GetByKey(CType(sPrioridad, Integer)) = True Then
                                oOBPP.PriorityDescription = oOBPP_Master.PriorityDescription
                                If oOBPP.Update() <> 0 Then
                                    oObjGlobal.SBOApp.StatusBar.SetText("Error Actualizando Prioridad para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    Exit Function
                                Else
                                    oObjGlobal.compañia.GetNewObjectCode(sPrioridad)
                                End If
                            Else
                                oOBPP.Priority = CType(sPrioridad, Integer)
                                oOBPP.PriorityDescription = oOBPP_Master.PriorityDescription
                                If oOBPP.Add() <> 0 Then
                                    oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Prioridad para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                    Exit Function
                                Else
                                    oObjGlobal.compañia.GetNewObjectCode(sPrioridad)
                                End If
                            End If
                            oOCRD.Priority = CType(sPrioridad, Integer)
                        End If
                    Else
                        oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa Master la prioridad " & sPrioridad, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                        Exit Function
                    End If
                End If
#End Region
                oOCRD.IBAN = oOCRD_Master.IBAN
#Region "Vacaciones"
                'Falta definir
#End Region
                'Pestaña Ejecución de pago
#Region "Ejecución de pago"
                oOCRD.HouseBankCountry = oOCRD_Master.HouseBankCountry
                oOCRD.HouseBank = oOCRD_Master.HouseBank
                oOCRD.HouseBankBranch = oOCRD_Master.HouseBankBranch
                oOCRD.HouseBankAccount = oOCRD_Master.HouseBankAccount

                oOCRD.DME = oOCRD_Master.DME
                oOCRD.InstructionKey = oOCRD_Master.InstructionKey
                oOCRD.ReferenceDetails = oOCRD_Master.ReferenceDetails

                oOCRD.PaymentBlock = oOCRD_Master.PaymentBlock
                oOCRD.SinglePayment = oOCRD_Master.SinglePayment

                oOCRD.BankChargesAllocationCode = oOCRD_Master.BankChargesAllocationCode


                oOCRD.BPPaymentMethods.Delete()
#Region "Vía de pago"
                For i = 0 To oOCRD.BPPaymentMethods.Count - 1
                    oOCRD.BPPaymentMethods.SetCurrentLine(0)
                    oOCRD.BPPaymentMethods.Delete()
                Next
                For i = 0 To oOCRD_Master.BPPaymentMethods.Count - 1
                    oOCRD_Master.BPPaymentMethods.SetCurrentLine(i)
                    'Comprobamos que exista la vía
                    Dim sViaPago As String = oOCRD_Master.BPPaymentMethods.PaymentMethodCode
                    oOPYM_Master = CType(oCompanyMaster.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWizardPaymentMethods), SAPbobsCOM.WizardPaymentMethods)
                    If sViaPago.Trim <> "" Then
                        If oOPYM_Master.GetByKey(sViaPago) = True Then
                            sSQL = "Select * FROM OPYM WHERE ""PayMethCod"" = '" & sViaPago & "' "
                            oRsOPYM.DoQuery(sSQL)
                            oOPYM = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWizardPaymentMethods), SAPbobsCOM.WizardPaymentMethods)
                            If oRsOPYM.RecordCount > 0 Then
                                If oOPYM.GetByKey(sViaPago) = True Then
                                    sExiste_OPYM = True
                                Else
                                    sExiste_OPYM = False
                                    oOPYM.PaymentMethodCode = sViaPago
                                End If
                                oOPYM.Description = oOPYM_Master.Description
                                oOPYM.Active = oOPYM_Master.Active
                                oOPYM.Type = oOPYM_Master.Type

                                'oOPYM.= oRsOPYM.Fields.Item("BankTransf").Value.ToString
                                oOPYM.BankCountry = oOPYM_Master.BankCountry
                                oOPYM.DefaultBank = oOPYM_Master.DefaultBank
                                oOPYM.DefaultAccount = oOPYM_Master.DefaultAccount
                                'Porcentaje gastos no lo veo
                                oOPYM.GroupByDate = oOPYM_Master.GroupByDate

                                If sExiste_OPYM = True Then
                                    If oOPYM.Update() <> 0 Then
                                        oObjGlobal.SBOApp.StatusBar.SetText("Error actualizando Vía de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                        Exit Function
                                    Else
                                        oObjGlobal.compañia.GetNewObjectCode(sViaPago)
                                    End If
                                Else
                                    If oOPYM.Add() <> 0 Then
                                        oObjGlobal.SBOApp.StatusBar.SetText("Error Creando Vía de pago para el IC " & sLicTradNum & " - " & oOCRD.CardName & " - " & oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                                        Exit Function
                                    Else
                                        oObjGlobal.compañia.GetNewObjectCode(sViaPago)
                                    End If
                                End If
                            End If
                        Else
                            oObjGlobal.SBOApp.StatusBar.SetText("Error grave. No se encuentra en la empresa Master la vía de pago " & sViaPago, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                            Exit Function
                        End If
                    End If
#End Region
                    oOCRD.BPPaymentMethods.PaymentMethodCode = oOCRD_Master.BPPaymentMethods.PaymentMethodCode
                    oOCRD.BPPaymentMethods.Add()
                Next
                oOCRD.PeymentMethodCode = oOCRD_Master.PeymentMethodCode
#End Region
                'Pestaña finanzas
#Region "Pestaña finanzas"
                oOCRD.FatherCard = oOCRD_Master.FatherCard
                oOCRD.FatherType = oOCRD_Master.FatherType

                oOCRD.DownPaymentInterimAccount = oOCRD_Master.DownPaymentInterimAccount
                oOCRD.DownPaymentClearAct = oOCRD_Master.DownPaymentClearAct
                'Falta una cuenta y las del botón

                'Falta connbp

                oOCRD.PlanningGroup = oOCRD_Master.PlanningGroup
                oOCRD.Affiliate = oOCRD_Master.Affiliate

                'Pestaña de impuesto
                oOCRD.Equalization = oOCRD_Master.Equalization
                oOCRD.VatIDNum = oOCRD_Master.VatIDNum
                oOCRD.ECommerceMerchantID = oOCRD_Master.ECommerceMerchantID
                oOCRD.AccrualCriteria = oOCRD_Master.AccrualCriteria
                oOCRD.CertificateNumber = oOCRD_Master.CertificateNumber
                oOCRD.ExpirationDate = oOCRD_Master.ExpirationDate

                oOCRD.OperationCode347 = oOCRD_Master.OperationCode347
                oOCRD.InsuranceOperation347 = oOCRD_Master.InsuranceOperation347
#End Region
#Region "Pestaña propiedades"
                For i = 1 To 64
                    oOCRD.Properties(i) = oOCRD_Master.Properties(i)
                Next
#End Region
                oOCRD.FreeText = oOCRD_Master.FreeText
#Region "Documentos electrónicos"
                oOCRD.EDocGenerationType = oOCRD.EDocGenerationType
                oOCRD.FCERelevant = oOCRD.FCERelevant
                oOCRD.FCEValidateBaseDelivery = oOCRD.FCEValidateBaseDelivery
#End Region
#Region "Campos de usuario"
                sSQL = "select ""AliasID"" FROM """ & oObjGlobal.compañia.CompanyDB & """.""CUFD"" WHERE ""TableID"" = 'OCRD';"
                oRsCamposUsuario.DoQuery(sSQL)
                For i = 0 To oRsCamposUsuario.RecordCount - 1
                    Try
                        Dim sCampo As String = "U_" & oRsCamposUsuario.Fields.Item("AliasID").Value.ToString
                        oOCRD.UserFields.Fields.Item(sCampo).Value = oOCRD_Master.UserFields.Fields.Item(sCampo).Value
                    Catch ex As Exception

                    End Try
                    oRsCamposUsuario.MoveNext()
                Next
#End Region

                If oOCRD.Update() <> 0 Then
                    oObjGlobal.SBOApp.StatusBar.SetText("Error actualizando IC - " & sLicTradNum & " - " & oOCRD.CardName & " - " &
                                                                oObjGlobal.compañia.GetLastErrorCode & " / " & oObjGlobal.compañia.GetLastErrorDescription, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                Else
                    oObjGlobal.SBOApp.StatusBar.SetText("IC Actualizado - " & sLicTradNum & " - " & oOCRD.CardName, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                End If
            End If
            Sincroniza_proveedor = True
        Catch ex As Exception
            Throw ex
        Finally
#Region "Liberar"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRD, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRD_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsGrupos_Des, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRG, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRG_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOSHP, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOSHP_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsClase_Expe_Des, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOIDC, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOIDC_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOOND, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOOND_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsRamo, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOSLP, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOSLP_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsEmpleado, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsResponsable, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCTG, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCTG_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCondPago, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsPrioridad, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOBPP, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOBPP_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsOPYM, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOPYM, Object)) : EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOPYM_Master, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCamposUsuario, Object))
#End Region
        End Try
    End Function

    Public Shared Function Sincroniza_Series(ByRef oCompanyDes As SAPbobsCOM.Company, ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
#Region "Variables"
        Dim sSQL As String = "" : Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing : Dim oSeriesService As SAPbobsCOM.SeriesService = Nothing
        Dim oSeries As SAPbobsCOM.Series = Nothing
        Dim oSeriesParams As SAPbobsCOM.SeriesParams = Nothing
        Dim oDocSeriesParam As SAPbobsCOM.DocumentSeriesParams = Nothing
        Dim sObjectCode_Nombre As String = "" : Dim sSerieDflt As String = ""
#End Region
        Sincroniza_Series = False
        Try
            oRs = CType(oCompanyDes.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            'Primero Insertamos los "Indicator"(Indicador de periodo en la tabla OPID)
            sSQL = "INSERT INTO """ & oCompanyDes.CompanyDB & """.""OPID"" "
            sSQL &= " SELECT ""O"".""Indicator"" "
            sSQL &= " FROM """ & oObjGlobal.compañia.CompanyDB & """.""OPID"" ""O"" "
            sSQL &= " LEFT JOIN """ & oCompanyDes.CompanyDB & """.""OPID"" ""D"" ON ""O"".""Indicator""=""D"".""Indicator"" "
            sSQL &= " WHERE ifnull(""D"".""Indicator"",'')='' "

            If oObjGlobal.refDi.SQL.executeNonQuery(sSQL) <> True Then
                oObjGlobal.SBOApp.StatusBar.SetText("Error sincronizando Indicadores de periodo.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            End If
            sSQL = "SELECT ""C"".""DfltSeries"",  ""M"".* FROM """ & oObjGlobal.compañia.CompanyDB & """.""NNM1"" ""M"" "
            sSQL &= "INNER JOIN """ & oObjGlobal.compañia.CompanyDB & """.""ONNM"" ""C"" ON ""C"".""ObjectCode""=""M"".""ObjectCode"" "
            sSQL &= " Left JOIN """ & oCompanyDes.CompanyDB & """.""NNM1"" ""D""  "
            sSQL &= " ON ""M"".""ObjectCode""=""D"".""ObjectCode"" And  ""M"".""DocSubType""=""D"".""DocSubType"" And ""M"".""SeriesName""=""D"".""SeriesName"" And  ""M"".""SeriesType""=""D"".""SeriesType"" "
            sSQL &= " WHERE ifnull(""D"".""Series"",'0')=0 "
            sSQL &= " Order by ""M"".""ObjectCode"",""M"".""SeriesName"" "
            oRs.DoQuery(sSQL)
            For i = 0 To oRs.RecordCount - 1
                sObjectCode_Nombre = Nombre_ObjectType(oRs.Fields.Item("ObjectCode").Value.ToString)
                oCmpSrv = oCompanyDes.GetCompanyService()
                oSeriesService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService), SAPbobsCOM.SeriesService)
                oSeries = CType(oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeries), SAPbobsCOM.Series)
                oSeries.Name = oRs.Fields.Item("SeriesName").Value.ToString
                oSeries.Document = oRs.Fields.Item("ObjectCode").Value.ToString
                oSeries.PeriodIndicator = oRs.Fields.Item("Indicator").Value.ToString
                oSeries.GroupCode = CType(oRs.Fields.Item("GroupCode").Value.ToString, SAPbobsCOM.BoSeriesGroupEnum)
                If IsNumeric(oRs.Fields.Item("InitialNum").Value.ToString) Then
                    oSeries.InitialNumber = CType(oRs.Fields.Item("InitialNum").Value.ToString, Integer)
                End If
                If IsNumeric(oRs.Fields.Item("LastNum").Value.ToString) And CType(oRs.Fields.Item("LastNum").Value.ToString, Integer) > 0 Then
                    oSeries.LastNumber = CType(oRs.Fields.Item("LastNum").Value.ToString, Integer)
                End If
                oSeries.Prefix = oRs.Fields.Item("BeginStr").Value.ToString
                oSeries.Suffix = oRs.Fields.Item("EndStr").Value.ToString
                oSeries.Remarks = oRs.Fields.Item("Remark").Value.ToString
                Try
                    'Graba la serie
                    oSeriesParams = oSeriesService.AddSeries(oSeries)
                    oObjGlobal.SBOApp.StatusBar.SetText("Sincronizado ObjectCode: " & sObjectCode_Nombre & " - Series Name: " &
                                                        oRs.Fields.Item("SeriesName").Value.ToString & " - Inicio: " & oRs.Fields.Item("InitialNum").Value.ToString &
                                                        " - Fin : " & oRs.Fields.Item("LastNum").Value.ToString, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                    ''set document type(e.g. Deliveries=15)
                    'oSeriesParams.Document = oRs.Fields.Item("ObjectCode").Value.ToString

                    ''set the series code
                    'oSeriesParams.Series = oSeriesParams.Series

                    ''attach Series to document
                    'Call oSeriesService.AttachSeriesToDocument(oDocSeriesParam)

                    'Para poner la serie por defecto
                    If oRs.Fields.Item("DfltSeries").Value.ToString = oRs.Fields.Item("Series").Value.ToString Then
                        oDocSeriesParam = CType(oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiDocumentSeriesParams), SAPbobsCOM.DocumentSeriesParams)
                        oDocSeriesParam.Document = oRs.Fields.Item("ObjectCode").Value.ToString
                        'oDocSeriesParam.Series = oSeriesParams.Series Esto es el que se ha creado
                        sSQL = "SELECT ""SeriesName"" FROM ""NNM1"" WHERE ""Series""=" & oRs.Fields.Item("DfltSeries").Value.ToString & " and ""ObjectCode""='" & oRs.Fields.Item("ObjectCode").Value.ToString & "' "
                        sSerieDflt = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                        sSQL = "SELECT ""Series"" FROM """ & oCompanyDes.CompanyDB & """.""NNM1"" WHERE ""SeriesName""='" & sSerieDflt & "' and ""ObjectCode""='" & oRs.Fields.Item("ObjectCode").Value.ToString & "' "
                        sSerieDflt = oObjGlobal.refDi.SQL.sqlStringB1(sSQL)
                        oDocSeriesParam.Series = CType(sSerieDflt, Integer)
                        Call oSeriesService.SetDefaultSeriesForCurrentUser(oDocSeriesParam)

                        oObjGlobal.SBOApp.StatusBar.SetText("ObjectCode: " & sObjectCode_Nombre & " - Serie Por Defecto: " & oRs.Fields.Item("DfltSeries").Value.ToString, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                    End If

                Catch ex As Exception
                    oObjGlobal.SBOApp.StatusBar.SetText("ObjectCode: " & sObjectCode_Nombre & " - Series Name: " &
                                                        oRs.Fields.Item("SeriesName").Value.ToString & " - Inicio: " & oRs.Fields.Item("InitialNum").Value.ToString &
                                                        " - Fin : " & oRs.Fields.Item("LastNum").Value.ToString & ". Error: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
                End Try
                oRs.MoveNext()
            Next
            Sincroniza_Series = True
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Shared Function Nombre_ObjectType(ByVal sObjectType As String) As String
        Dim sObjectCode As String = ""
        Nombre_ObjectType = ""

        Try
            Select Case sObjectType
                Case "1" : sObjectCode = "Cuentas de mayor"
                Case "2" : sObjectCode = "Interlocutor comercial"
                Case "3" : sObjectCode = "Bancos"
                Case "4" : sObjectCode = "Artículos"
                Case "5" : sObjectCode = "Definición de Impuesto"
                Case "6" : sObjectCode = "Lista de precios"
                Case "7" : sObjectCode = "Precios Especiales"
                Case "8" : sObjectCode = "Propiedades de Artículos"
                Case "9" : sObjectCode = "Diferencias de cambio"
                Case "10" : sObjectCode = "Grupos de Interlocutores"
                Case "11" : sObjectCode = "Contactos"
                Case "12" : sObjectCode = "Usuarios"
                Case "13" : sObjectCode = "Facturas de venta"
                Case "14" : sObjectCode = "Abonos de venta"
                Case "15" : sObjectCode = "Entregas de venta"
                Case "16" : sObjectCode = "Devoluciones"
                Case "17" : sObjectCode = "Pedidos de venta"
                Case "18" : sObjectCode = "Facturas de compras"
                Case "19" : sObjectCode = "Abono de compras"
                Case "20" : sObjectCode = "Pedido de entrada de mercancías"
                Case "21" : sObjectCode = "Devolución de mercancías"
                Case "22" : sObjectCode = "Pedido de compras"
                Case "23" : sObjectCode = "Oferta de ventas"
                Case "24" : sObjectCode = "Cobros"
                Case "25" : sObjectCode = "Depósitos"
                Case "26" : sObjectCode = "Historial reconciliación"
                Case "27" : sObjectCode = "Registro de cheques"
                Case "28" : sObjectCode = "Entrada de documento preliminar"
                Case "29" : sObjectCode = "Lista de comprobantes de diario"
                Case "30" : sObjectCode = "Diario"
                Case "31" : sObjectCode = "Artículos: Almacén"
                Case "32" : sObjectCode = "Preferencias de impresión"
                Case "33" : sObjectCode = "Actividades"
                Case "34" : sObjectCode = "Contabilizaciones periódicas"
                Case "35" : sObjectCode = "Series de documentos"
                Case "36" : sObjectCode = "Tarjetas de crédito"
                Case "37" : sObjectCode = "Códigos de moneda"
                Case "38" : sObjectCode = "Códigos de CPI"
                Case "39" : sObjectCode = "Empresa"
                Case "40" : sObjectCode = "Condiciones de pago"
                Case "41" : sObjectCode = "Preferencias"
                Case "42" : sObjectCode = "Extracto bancario externo recibido"
                Case "43" : sObjectCode = "Fabricantes"
                Case "44" : sObjectCode = "Propiedades tarjeta"
                Case "45" : sObjectCode = "Códigos de registro en el diario"
                Case "46" : sObjectCode = "Pagos"
                Case "47" : sObjectCode = "Números de serie"
                Case "48" : sObjectCode = "Gastos de carga"
                Case "49" : sObjectCode = "Clases de entrega"
                Case "50" : sObjectCode = "Unidades de longitud"
                Case "51" : sObjectCode = "Unidades de peso"
                Case "52" : sObjectCode = "Grupos de artículos"
                Case "53" : sObjectCode = "Empleado del departamento de ventas"
                Case "54" : sObjectCode = "Informe - Criterios de selección"
                Case "55" : sObjectCode = "Modelos de transacción"
                Case "56" : sObjectCode = "Grupos de aduanas"
                Case "57" : sObjectCode = "Cheques para pago"
                Case "58" : sObjectCode = "Diario de almacén"
                Case "59" : sObjectCode = "Entrada de mercancías"
                Case "60" : sObjectCode = "Salida de mercancías"
                Case "61" : sObjectCode = "Centro de coste"
                Case "62" : sObjectCode = "Norma de reparto"
                Case "63" : sObjectCode = "Códigos de proyecto"
                Case "64" : sObjectCode = "Almacenes"
                Case "65" : sObjectCode = "Grupo de comisiones"
                Case "66" : sObjectCode = "Árbol de productos"
                Case "67" : sObjectCode = "Traslado"
                Case "68" : sObjectCode = "Instrucciones fabricación"
                Case "69" : sObjectCode = "Precios de entrega"
                Case "70" : sObjectCode = "Vías de pago"
                Case "71" : sObjectCode = "Pago con tarjeta de crédito"
                Case "72" : sObjectCode = "Gestión de tarjetas de crédito"
                Case "73" : sObjectCode = "Número de catálogo del cliente/proveedor"
                Case "74" : sObjectCode = "Pagos de crédito"
                Case "75" : sObjectCode = "Tasas IPC y ME"
                Case "76" : sObjectCode = "Depósito con fecha posterior al día de la creación"
                Case "77" : sObjectCode = "Presupuesto"
                Case "78" : sObjectCode = "Método subreparto costes presup."
                Case "79" : sObjectCode = "Cadenas de comercio al por menor"
                Case "80" : sObjectCode = "Plantilla de alertas"
                Case "81" : sObjectCode = "Alertas"
                Case "82" : sObjectCode = "Alertas recibidas"
                Case "83" : sObjectCode = "Mensajes enviados"
                Case "84" : sObjectCode = "Objetos de actividad"
                Case "85" : sObjectCode = "Precios especiales para grupos"
                Case "86" : sObjectCode = "Inicio de la aplicación"
                Case "87" : sObjectCode = "Lista de distribución"
                Case "88" : sObjectCode = "Tipos de envío"
                Case "89" : sObjectCode = "OSAL    Outgoing"
                Case "90" : sObjectCode = "OTRA    Transition"
                Case "91" : sObjectCode = "Escenario de presupuesto"
                Case "92" : sObjectCode = "Precios de interés"
                Case "93" : sObjectCode = "Opciones de usuario"
                Case "94" : sObjectCode = "Números de serie para artículos"
                Case "95" : sObjectCode = "Modelos informe financiero"
                Case "96" : sObjectCode = "Categorías de informes financieros"
                Case "97" : sObjectCode = "Oportunidad"
                Case "98" : sObjectCode = "Interés"
                Case "99" : sObjectCode = "Nivel del tipo de interés"
                Case "100" : sObjectCode = "Fuente de información"
                Case "101" : sObjectCode = "Nivel de oportunidad"
                Case "102" : sObjectCode = "Causas del defecto"
                Case "103" : sObjectCode = "Clases actividad"
                Case "104" : sObjectCode = "Lugar de reuniones"
                Case "105" : sObjectCode = "Llamadas de servicio"
                Case "106" : sObjectCode = "Número de lote para artículo"
                Case "107" : sObjectCode = "Artículos alternativos 2"
                Case "108" : sObjectCode = "Partners"
                Case "109" : sObjectCode = "Competidores"
                Case "110" : sObjectCode = "OUVV    Validaciones de usuario"
                Case "111" : sObjectCode = "Período contable"
                Case "112" : sObjectCode = "Documentos preliminares"
                Case "113" : sObjectCode = "Lotes y números de serie"
                Case "114" : sObjectCode = "OUDC    Pantalla de usuario Cat."
                Case "115" : sObjectCode = "Acreedor - Pelecard"
                Case "116" : sObjectCode = "Jerarquía de la deducción de la retención de impuestos"
                Case "117" : sObjectCode = "Grupos de deducción de retención de impuestos"
                Case "118" : sObjectCode = "Sucursales"
                Case "119" : sObjectCode = "Departamentos"
                Case "120" : sObjectCode = "Nivel de confirmación"
                Case "121" : sObjectCode = "Modelos de autorización"
                Case "122" : sObjectCode = "Documentos de confirmación"
                Case "123" : sObjectCode = "Cheques para documentos preliminares de pago"
                Case "124" : sObjectCode = "CINF    Información de la compañía"
                Case "125" : sObjectCode = "OEXD    Definir porte"
                Case "126" : sObjectCode = "Autoridades de impuestos sobre ventas"
                Case "127" : sObjectCode = "Clase de autoridades de impuestos sobre ventas"
                Case "128" : sObjectCode = "Indicadores de IVA"
                Case "129" : sObjectCode = "Países"
                Case "130" : sObjectCode = "Estados"
                Case "131" : sObjectCode = "Address Formats"
                Case "132" : sObjectCode = "Factura de corrección de clientes"
                Case "133" : sObjectCode = "Categorías de consultas"
                Case "134" : sObjectCode = "OQCN    Query Catagories"
                Case "135" : sObjectCode = "Operación triangular"
                Case "136" : sObjectCode = "Migración de datos"
                Case "137" : sObjectCode = "OCSTN   Workstation ID"
                Case "138" : sObjectCode = "Indicador"
                Case "139" : sObjectCode = "Transporte de mercancías"
                Case "140" : sObjectCode = "Propuesta de pago"
                Case "141" : sObjectCode = "Asistente consulta"
                Case "142" : sObjectCode = "Segmentación de cuenta"
                Case "143" : sObjectCode = "Categorías de segmentación de cuentas"
                Case "144" : sObjectCode = "Emplazamiento"
                Case "145" : sObjectCode = "Formularios 1099"
                Case "146" : sObjectCode = "Ciclo"
                Case "147" : sObjectCode = "Vías de pago para asistente de pagos"
                Case "148" : sObjectCode = "1099 Balance apertura"
                Case "149" : sObjectCode = "Tipo de interés de reclamación"
                Case "150" : sObjectCode = "Prioridades IC"
                Case "151" : sObjectCode = "Reclamaciones"
                Case "152" : sObjectCode = "Campos de usuario: descripción"
                Case "153" : sObjectCode = "Tablas usuario"
                Case "154" : sObjectCode = "Elementos de mi menú"
                Case "155" : sObjectCode = "Ejecución pago"
                Case "156" : sObjectCode = "Lista de picking"
                Case "157" : sObjectCode = "Asistente de pago"
                Case "158" : sObjectCode = "Tabla de resultados de pagos"
                Case "159" : sObjectCode = "OPYB    Payment Block"
                Case "160" : sObjectCode = "Consultas"
                Case "161" : sObjectCode = "Ind.banco central"
                Case "162" : sObjectCode = "Revaloración de inventario"
                Case "163" : sObjectCode = "Factura de corrección de proveedores"
                Case "164" : sObjectCode = "Anulación de factura de corrección de proveedores"
                Case "165" : sObjectCode = "Factura de corrección de clientes"
                Case "166" : sObjectCode = "Anulación de factura de corrección de clientes"
                Case "167" : sObjectCode = "Status de llamada de servicio"
                Case "168" : sObjectCode = "Tipos de llamada de servicio"
                Case "169" : sObjectCode = "Tipos de problema de llamada de servicio"
                Case "170" : sObjectCode = "Modelo de contrato"
                Case "171" : sObjectCode = "Empleados"
                Case "172" : sObjectCode = "Tipos de empleado"
                Case "173" : sObjectCode = "Status de empleado"
                Case "174" : sObjectCode = "Motivo de rescisión"
                Case "175" : sObjectCode = "Clases de formación"
                Case "176" : sObjectCode = "Tarjeta del equipo del cliente"
                Case "177" : sObjectCode = "Nombre de agente"
                Case "178" : sObjectCode = "Retención de impuestos"
                Case "179" : sObjectCode = "Reports 347, 349 e IR ya visualizados"
                Case "180" : sObjectCode = "Informe fiscal"
                Case "181" : sObjectCode = "Efecto para pagos"
                Case "182" : sObjectCode = "OBOT    Bill Of Exchang Transaction"
                Case "183" : sObjectCode = "Formato de fichero"
                Case "184" : sObjectCode = "Indicador de período"
                Case "185" : sObjectCode = "Créditos de dudoso cobro"
                Case "186" : sObjectCode = "Tabla de festivos"
                Case "187" : sObjectCode = "Interlocutor comercial: Cuenta bancaria"
                Case "188" : sObjectCode = "Status de solución de llamada de servicio"
                Case "189" : sObjectCode = "Soluciones de llamada de servicio"
                Case "190" : sObjectCode = "Contratos de servicio"
                Case "191" : sObjectCode = "Llamadas de servicio"
                Case "192" : sObjectCode = "Orígenes de llamada de servicio"
                Case "193" : sObjectCode = "OUKD    Descripción de la clave de usuario"
                Case "194" : sObjectCode = "Cola"
                Case "195" : sObjectCode = "Asistente de inflación"
                Case "196" : sObjectCode = "Condiciones de reclamación"
                Case "197" : sObjectCode = "Asistente de reclamaciones"
                Case "198" : sObjectCode = "Previsión de ventas"
                Case "199" : sObjectCode = "Escenarios de planificación de necesidades"
                Case "200" : sObjectCode = "Territorios"
                Case "201" : sObjectCode = "Ramos"
                Case "202" : sObjectCode = "Orden de fabricación"
                Case "203" : sObjectCode = "Anticipo de clientes"
                Case "204" : sObjectCode = "Anticipo de proveedores"
                Case "205" : sObjectCode = "Clases de paquete"
                Case "206" : sObjectCode = "Objeto definido por el usuario"
                Case "207" : sObjectCode = "Propiedad de datos - Objetos"
                Case "208" : sObjectCode = "Propiedad datos - Excepciones"
                Case "210" : sObjectCode = "Posición del empleado"
                Case "211" : sObjectCode = "Equipos de empleados"
                Case "212" : sObjectCode = "Relaciones"
                Case "213" : sObjectCode = "Fecha de recomendación"
                Case "214" : sObjectCode = "Árbol de autorización de usuario"
                Case "215" : sObjectCode = "Texto predefinido"
                Case "216" : sObjectCode = "Definición de casilla"
                Case "217" : sObjectCode = "Status de operación"
                Case "218" : sObjectCode = "OCHF    312"
                Case "219" : sObjectCode = "OCSHS   Valores definidos por el usuario"
                Case "220" : sObjectCode = "Tipos de períodos"
                Case "221" : sObjectCode = "Anexos"
                Case "222" : sObjectCode = "Trama filtro"
                Case "223" : sObjectCode = "Tabla idioma usuario"
                Case "224" : sObjectCode = "Traducción multilingüe"
                Case "225" : sObjectCode = "OAPA3           225"
                Case "226" : sObjectCode = "OAPA4           226"
                Case "227" : sObjectCode = "OAPA5           227"
                Case "229" : sObjectCode = "SDIS   Interfaz dinámica (cadenas)"
                Case "230" : sObjectCode = "Reconciliaciones grabadas"
                Case "231" : sObjectCode = "Cuentas banco propio"
                Case "232" : sObjectCode = "RDOC    Documento"
                Case "233" : sObjectCode = "Creación documentos grupos parámetros"
                Case "234" : sObjectCode = "OMHD    #740"
                Case "238" : sObjectCode = "Categoría de cuenta"
                Case "239" : sObjectCode = "Códigos de imputación de gastos bancarios"
                Case "241" : sObjectCode = "Operaciones de flujo de caja - Apuntes"
                Case "242" : sObjectCode = "Posición de documento flujo de caja"
                Case "247" : sObjectCode = "Lugar comercial"
                Case "250" : sObjectCode = "Calendario de era local"
                Case "251" : sObjectCode = "Dimensión contabilidad costes"
                Case "254" : sObjectCode = "Tabla de códigos de servicio"
                Case "256" : sObjectCode = "Grupo de materiales"
                Case "257" : sObjectCode = "Código NCM"
                Case "258" : sObjectCode = "CFOP para nota fiscal"
                Case "259" : sObjectCode = "Código CST para Nota Fiscal"
                Case "260" : sObjectCode = "Utilización de nota fiscal"
                Case "261" : sObjectCode = "Procedimiento de fecha de cierre"
                Case "263" : sObjectCode = "Numeración de nota fiscal"
                Case "265" : sObjectCode = "Regiones"
                Case "266" : sObjectCode = "Determinación de indicador de IVA"
                Case "267" : sObjectCode = "Clase de documento de efecto"
                Case "268" : sObjectCode = "Portafolio de efectos"
                Case "269" : sObjectCode = "Instrucción de efecto"
                Case "271" : sObjectCode = "Parámetros de impuesto"
                Case "275" : sObjectCode = "Combinación de clases de impuestos"
                Case "276" : sObjectCode = "Tabla maestra de fórmulas de impuestos"
                Case "278" : sObjectCode = "Código CNAE"
                Case "280" : sObjectCode = "Factura de impuestos de ventas"
                Case "281" : sObjectCode = "Factura de impuestos de compras"
                Case "283" : sObjectCode = "Número de declaración de aduana de portes"
                Case "290" : sObjectCode = "Recursos"
                Case "291" : sObjectCode = "Propiedades de recurso"
                Case "292" : sObjectCode = "Grupos de recursos"
                Case "321" : sObjectCode = "Reconciliación interna"
                Case "541" : sObjectCode = "Datos maestros de TPV"
                Case "1179" : sObjectCode = "Documentos preliminares"
                Case "10000105" : sObjectCode = "Opciones de servicio de mensajes"
                Case "10000044" : sObjectCode = "Datos maestros números de lote"
                Case "10000045" : sObjectCode = "Datos maestros números de serie"
                Case "10000062" : sObjectCode = "IVL Vs OINM Keys"
                Case "10000071" : sObjectCode = "Contabilización de stocks"
                Case "10000073" : sObjectCode = "Ejercicio maestro"
                Case "10000074" : sObjectCode = "Secciones"
                Case "10000075" : sObjectCode = "Serie de certificados"
                Case "10000077" : sObjectCode = "Clase de sujeto pasivo"
                Case "10000196" : sObjectCode = "Lista de clases de documento"
                Case "10000197" : sObjectCode = "Grupo de unidades de medida"
                Case "10000199" : sObjectCode = "Datos maestros de la unidad de medida"
                Case "10000203" : sObjectCode = "Configuración del campo de ubicación"
                Case "10000204" : sObjectCode = "Atributo de ubicación"
                Case "10000205" : sObjectCode = "Subnivel de almacén"
                Case "10000206" : sObjectCode = "Ubicación"
                Case "140000041" : sObjectCode = "Código DNF"
                Case "231000000" : sObjectCode = "Grupo de autorización"
                Case "234000004" : sObjectCode = "Grupo de correo electrónico"
                Case "243000001" : sObjectCode = "Código de pago del gobierno"
                Case "310000001" : sObjectCode = "Saldo de apertura de inventario"
                Case "310000008" : sObjectCode = "Atributos de lote en la ubicación"
                Case "410000005" : sObjectCode = "Formato de lista legal"
                Case "480000001" : sObjectCode = "Objeto: Transferencia de empleado de RR. HH."
                Case "540000005" : sObjectCode = "Determinación de indicador de IVA"
                Case "540000006" : sObjectCode = "Solicitud de pedido"
                Case "540000040" : sObjectCode = "Modelo de transacción periódica"
                Case "540000042" : sObjectCode = "Tipo de centro de coste"
                Case "540000048" : sObjectCode = "Clase de periodificación"
                Case "540000056" : sObjectCode = "Modelo nota fiscal"
                Case "540000067" : sObjectCode = "Indexador de combustible de Brasil"
                Case "540000068" : sObjectCode = "Indexador de bebidas de Brasil"
                Case "1210000000" : sObjectCode = "Tabla principal de cockpit"
                Case "1250000001" : sObjectCode = "Solicitud de traslado"
                Case "1250000025" : sObjectCode = "Acuerdo global"
                Case "1320000000" : sObjectCode = "Paquete de indicadores de rendimiento clave"
                Case "1320000002" : sObjectCode = "Grupo de destino"
                Case "1320000012" : sObjectCode = "Campaña"
                Case "1320000028" : sObjectCode = "Códigos de operaciones de Retorno"
                Case "1320000039" : sObjectCode = "Código fuente de producto"
                Case "1470000000" : sObjectCode = "Clases de amortización de activos fijos"
                Case "1470000002" : sObjectCode = "Determinación de cuentas de activos fijos"
                Case "1470000003" : sObjectCode = "Áreas de amortización de activo fijo"
                Case "1470000004" : sObjectCode = "Pools tipos amortización"
                Case "1470000032" : sObjectCode = "Clases activos fijos"
                Case "1470000046" : sObjectCode = "Grupos de activos"
                Case "1470000048" : sObjectCode = "Criterios de determinación de cuenta de mayor - Inventario"
                Case "1470000049" : sObjectCode = "Capitalización"
                Case "1470000057" : sObjectCode = "Reglas avanzadas de cuenta de mayor"
                Case "1470000060" : sObjectCode = "Abono"
                Case "1470000062" : sObjectCode = "Datos maestros de código de barras"
                Case "1470000065" : sObjectCode = "Recuento de inventario"
                Case "1470000077" : sObjectCode = "Grupos de descuento"
                Case "1470000092" : sObjectCode = "Determinación de recuento de ciclo"
                Case "1470000113" : sObjectCode = "Solicitud de compra"
                Case "1620000000" : sObjectCode = "Workflow - Detalles de la tarea"
                Case Else : sObjectCode = sObjectType


            End Select
            Nombre_ObjectType = sObjectCode
        Catch ex As Exception
            Throw ex
        End Try

    End Function

End Class
