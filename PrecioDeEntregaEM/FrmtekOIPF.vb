Public Class FrmtekOIPF

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private csFormUID As String
    Private stDocNum As String
    Friend Monto As Double


    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Me.stDocNum = stDocNum
    End Sub

    'Private Property stRuta As String

    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function openForm(ByVal psDirectory As String, ByVal DocNum As String)

        Try

            csFormUID = "TekDeliveryPrices"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")

            End If

            '--- Referencia de Forma
            setForm(csFormUID)

            InsertDelyPri(DocNum)

            AgregarPrecios()

            DropTemp()

        Catch ex As Exception
            If (ex.Message <> "") Then
                cSBOApplication.MessageBox("FrmTratamientoPedidos. No se pudo iniciar la forma. " & ex.Message)
            End If
            Me.close()
        End Try
    End Function


    '//----- CIERRA LA VENTANA
    Public Function close() As Integer
        close = 0
        coForm.Close()
    End Function


    '//----- ABRE LA FORMA DENTRO DE LA APLICACION
    Public Function setForm(ByVal psFormUID As String) As Integer
        Try
            setForm = 0
            '//ESTABLECE LA REFERENCIA A LA FORMA
            coForm = cSBOApplication.Forms.Item(psFormUID)
            '//OBTIENE LA REFERENCIA A LOS USER DATA SOURCES
            setForm = getUserDataSources()
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar a la forma. " & ex.Message)
            setForm = -1
        End Try
    End Function


    '//----- OBTIENE LA REFERENCIA A LOS USERDATASOURCES
    Private Function getUserDataSources() As Integer
        'Dim llIndice As Integer
        Try
            coForm.Freeze(True)
            getUserDataSources = 0
            '//SI YA EXISTEN LOS DATASOURCES, SOLO LOS ASOCIA
            If (coForm.DataSources.UserDataSources.Count() > 0) Then
            Else '//EN CASO DE QUE NO EXISTAN, LOS CREA
                getUserDataSources = bindUserDataSources()
            End If
            coForm.Freeze(False)
        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al referenciar los UserDataSources" & ex.Message)
            getUserDataSources = -1
        End Try
    End Function


    '//----- ASOCIA LOS USERDATA A ITEMS
    Private Function bindUserDataSources() As Integer
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid

        Try
            bindUserDataSources = 0

            oGrid = coForm.Items.Item("2").Specific
            oDataTable = coForm.DataSources.DataTables.Add("DelyPri")
            oGrid.DataTable = oDataTable

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function


    Public Sub InsertDelyPri(ByVal DocNum As String)
        Dim DocEntry, ObjType, DocNumPE, DocDate, DocTotal As String
        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset

        Try

            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "call ShowButtonPriceDelivery('" & DocNum & "')"
            oRecSet.DoQuery(stQuery)

            If oRecSet.RecordCount = 1 Then

                DocEntry = oRecSet.Fields.Item("DocEntry").Value
                ObjType = oRecSet.Fields.Item("ObjType").Value
                DocNumPE = oRecSet.Fields.Item("DocNum").Value
                DocDate = oRecSet.Fields.Item("DocDate").Value
                DocTotal = oRecSet.Fields.Item("DocTotal").Value

                PushTemp(DocEntry, DocNumPE, DocDate, DocTotal)

                For i = 0 To 1

                    stQuery = "call DeliveryPrices('" & DocEntry & "','" & ObjType & "')"
                    oRecSet.DoQuery(stQuery)

                    If oRecSet.RecordCount = 1 Then

                        DocEntry = oRecSet.Fields.Item("DocEntry").Value
                        ObjType = oRecSet.Fields.Item("ObjType").Value
                        DocNumPE = oRecSet.Fields.Item("DocNum").Value
                        DocDate = oRecSet.Fields.Item("DocDate").Value
                        DocTotal = oRecSet.Fields.Item("DocTotal").Value

                        PushTemp(DocEntry, DocNumPE, DocDate, DocTotal)

                        i = 0

                    Else

                        i = 1

                    End If

                Next

            End If

        Catch ex As Exception
            cSBOApplication.MessageBox("DocumentoSBO. agregar elementos a la forma. " & ex.Message)
        Finally
        End Try
    End Sub


    Public Function PushTemp(ByVal DocEntry As String, ByVal DocNumPE As String, ByVal DocDate As String, ByVal DocTotal As Integer)

        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset

        oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            stQuery = "Insert Into TEMP_DeliveryPrice (DocEntry,DocNum,DocDate,DocTotal) values('" & DocEntry & "','" & DocNumPE & "','" & DocDate & "'," & DocTotal & ")"
            oRecSet.DoQuery(stQuery)

        Catch ex As System.Exception
            Throw New System.Exception(ex.Message)
        End Try

    End Function


    Public Function DropTemp()

        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset

        oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            ''Elimina tabla temporal
            stQuery = "Delete From TEMP_DeliveryPrice"
            oRecSet.DoQuery(stQuery)

        Catch ex As System.Exception
            Throw New System.Exception(ex.Message)
        End Try

    End Function


    Public Function AgregarPrecios()
        Dim oGrid As SAPbouiCOM.Grid
        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim oColumn As SAPbouiCOM.GridColumn

        Try

            oGrid = coForm.Items.Item("2").Specific
            oGrid.DataTable.Clear()

            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "Select * From TEMP_DeliveryPrice"
            oGrid.DataTable.ExecuteQuery(stQuery)

            oColumn = oGrid.Columns.Item("DOCENTRY")
            oColumn.LinkedObjectType = 69

            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False
            oGrid.Columns.Item(3).Editable = False

            Return 0

        Catch ex As Exception

            MsgBox("FrmtekDel. fallo la carga previa de la forma AgregarLineas: " & ex.Message)

        Finally

            oGrid = Nothing

        End Try

    End Function


    Public Function CrearOIPF(ByVal DocNum As String)
        Dim oOPDN As FrmtekOPDN
        Dim oGrid As SAPbouiCOM.Grid
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim stQueryH, stQueryH1, stQueryH2, stQueryH3, stQueryH4 As String
        Dim oRecSetH, oRecSetH1, oRecSetH2, oRecSetH3, oRecSetH4 As SAPbobsCOM.Recordset
        Dim oOIPF As SAPbobsCOM.LandedCostsService
        Dim oLandedCost As SAPbobsCOM.LandedCost
        Dim oLandedCost_ItemLine As SAPbobsCOM.LandedCost_ItemLine
        Dim oLandedCost_CostLine As SAPbobsCOM.LandedCost_CostLine
        Dim oLandedCostParams As SAPbobsCOM.LandedCostParams
        Dim ObjType, AclName, CardCode, Fecha, WhsCode, AlcCode, Currency As String
        Dim CostSum, Quantity, Price As Double
        Dim DocEntry, LineNum As Integer

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH1 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH2 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH3 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH4 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oOIPF = cSBOCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService)

        oLandedCost = oOIPF.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCost)

        Try

            coForm = cSBOApplication.Forms.Item("TekPriDelEM")
            oGrid = coForm.Items.Item("2").Specific
            oDataTable = oGrid.DataTable

            stQueryH = "call ShowButtonPriceDelivery('" & DocNum & "')"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount = 1 Then

                DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                ObjType = oRecSetH.Fields.Item("ObjType").Value

                For i = 0 To 1

                    stQueryH = "call DeliveryPrices('" & DocEntry & "','" & ObjType & "')"
                    oRecSetH.DoQuery(stQueryH)

                    If oRecSetH.RecordCount = 1 Then

                        DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                        ObjType = oRecSetH.Fields.Item("ObjType").Value

                        i = 0

                    Else

                        i = 1

                    End If

                Next

            Else

                stQueryH = "Select ""DocEntry"",""ObjType"" From OPDN where ""DocNum""=" & DocNum
                oRecSetH.DoQuery(stQueryH)

                If oRecSetH.RecordCount = 1 Then

                    DocEntry = oRecSetH.Fields.Item("DocEntry").Value
                    ObjType = oRecSetH.Fields.Item("ObjType").Value

                End If

            End If

            If ObjType = 69 Then

                stQueryH1 = "Select ""DocEntry"",""CardCode"" From OIPF where ""DocEntry""=" & DocEntry
                oRecSetH1.DoQuery(stQueryH1)

            Else

                stQueryH1 = "Select ""DocEntry"",""CardCode"" From OPDN where ""DocNum""=" & DocNum
                oRecSetH1.DoQuery(stQueryH1)

            End If

            If oRecSetH1.RecordCount = 1 Then

                CardCode = oRecSetH1.Fields.Item("CardCode").Value
                Fecha = Now.Date.Year & "/" & Now.Date.Month & "/" & Now.Date.Day

                oLandedCost.VendorCode = CardCode
                oLandedCost.PostingDate = Fecha
                oLandedCost.DueDate = Fecha
                oLandedCost.Remarks = "Basado en Pedido de entrada de mercancías " & DocNum

                Dim oLandedCostEntry As Long = 0
                'Dim GRPOEntry As Integer = 15

                If ObjType = 69 Then

                    stQueryH2 = "Select ""LineNum"",""Quantity"",""WhsCode"",""PriceFOB"" as ""Price"",""Currency"" From IPF1 where ""DocEntry""=" & oRecSetH1.Fields.Item("DocEntry").Value
                    oRecSetH2.DoQuery(stQueryH2)

                Else

                    stQueryH2 = "Select ""LineNum"",""Quantity"",""WhsCode"",""Price"",""Currency"" From PDN1 where ""DocEntry""=" & oRecSetH1.Fields.Item("DocEntry").Value
                    oRecSetH2.DoQuery(stQueryH2)

                End If

                If oRecSetH2.RecordCount > 0 Then

                    oRecSetH2.MoveFirst()

                    For i = 0 To oRecSetH2.RecordCount - 1

                        oLandedCost_ItemLine = oLandedCost.LandedCost_ItemLines.Add()
                        If ObjType = 69 Then
                            oLandedCost_ItemLine.BaseDocumentType = SAPbobsCOM.LandedCostBaseDocumentTypeEnum.asLandedCosts
                        Else
                            oLandedCost_ItemLine.BaseDocumentType = SAPbobsCOM.LandedCostBaseDocumentTypeEnum.asGoodsReceiptPO
                        End If

                        LineNum = oRecSetH2.Fields.Item("LineNum").Value
                        Quantity = oRecSetH2.Fields.Item("Quantity").Value
                        WhsCode = oRecSetH2.Fields.Item("WhsCode").Value
                        Price = oRecSetH2.Fields.Item("Price").Value
                        Currency = oRecSetH2.Fields.Item("Currency").Value

                        oLandedCost_ItemLine.BaseEntry = DocEntry
                        oLandedCost_ItemLine.BaseLine = LineNum
                        oLandedCost_ItemLine.Quantity = Quantity
                        oLandedCost_ItemLine.Warehouse = WhsCode

                        oRecSetH2.MoveNext()

                    Next

                End If

                For i = 0 To oDataTable.Rows.Count - 2

                    CostSum = oDataTable.GetValue("Importe", i)
                    AclName = oDataTable.GetValue("Precio de Entrega", i)

                    If CostSum <> 0 Then

                        stQueryH3 = "Select ""AlcCode"" from ""OALC"" where ""AlcName""='" & AclName & "'"
                        oRecSetH3.DoQuery(stQueryH3)

                        If oRecSetH3.RecordCount > 0 Then

                            AlcCode = oRecSetH3.Fields.Item("AlcCode").Value

                            oLandedCost_CostLine = oLandedCost.LandedCost_CostLines.Add()
                            oLandedCost_CostLine.LandedCostCode = AlcCode
                            oLandedCost_CostLine.amount = CostSum

                        End If

                    End If

                Next

                oLandedCostParams = oOIPF.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCostParams)

                'Add a landed cost 
                oLandedCostParams = oOIPF.AddLandedCost(oLandedCost)
                oLandedCostEntry = oLandedCostParams.LandedCostNumber

                stQueryH4 = "Select ""DocNum"" from OIPF where ""DocEntry""=" & oLandedCostEntry
                oRecSetH4.DoQuery(stQueryH4)

                If oRecSetH4.RecordCount = 1 Then

                    cSBOApplication.MessageBox("Se creo exitosamente el Precio de Entrega: " & oRecSetH4.Fields.Item("DocNum").Value)

                    oOPDN = New FrmtekOPDN
                    oOPDN.AgregarPrecios(CardCode, "TekPriDelEM")
                    oOPDN.HideOrShowFormItems(DocNum, "TekPriDelEM")

                End If

            End If


        Catch ex As Exception
            cSBOApplication.MessageBox("cerrarOrdenesVentas: " & ex.Message)
            Return -1
        End Try
    End Function


End Class
