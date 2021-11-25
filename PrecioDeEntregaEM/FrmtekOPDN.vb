Public Class FrmtekOPDN

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
    Public Function openForm(ByVal psDirectory As String, ByVal CardCode As String, ByVal DocNum As String)

        Try

            csFormUID = "TekPriDelEM"
            '//CARGA LA FORMA
            If (loadFormXML(cSBOApplication, csFormUID, psDirectory + "\Forms\" + csFormUID + ".srf") <> 0) Then

                Err.Raise(-1, 1, "")

            End If

            '--- Referencia de Forma
            setForm(csFormUID)

            AgregarPrecios(CardCode, csFormUID)

            HideOrShowFormItems(DocNum, csFormUID)

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
        Dim loText As SAPbouiCOM.EditText
        Dim loDS As SAPbouiCOM.UserDataSource
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid

        Try
            bindUserDataSources = 0

            loDS = coForm.DataSources.UserDataSources.Add("dsDate", SAPbouiCOM.BoDataType.dt_DATE)
            loText = coForm.Items.Item("4").Specific    'identifico mi caja de texto
            loText.DataBind.SetBound(True, "", "dsDate")    ' uno mi userdatasources a mi caja de fecha

            oGrid = coForm.Items.Item("2").Specific
            oDataTable = coForm.DataSources.DataTables.Add("DelPri")
            oGrid.DataTable = oDataTable

        Catch ex As Exception
            cSBOApplication.MessageBox("FrmTratamientoPedidos. Al crear los UserDataSources. " & ex.Message)
            bindUserDataSources = -1
        Finally
            loText = Nothing
            loDS = Nothing
            oDataTable = Nothing
            oGrid = Nothing
        End Try
    End Function


    '----- carga los ultimos precios de entrega segun el proveedor
    Public Function AgregarPrecios(ByVal CardCode As String, ByVal psFormUID As String)
        Dim oGrid As SAPbouiCOM.Grid
        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset

        Try

            coForm = cSBOApplication.Forms.Item(psFormUID)

            coForm.DataSources.UserDataSources.Item("dsDate").Value = Nothing

            oGrid = coForm.Items.Item("2").Specific
            oGrid.DataTable.Clear()

            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "call AgregarDelPrice('" & CardCode & "')"
            oGrid.DataTable.ExecuteQuery(stQuery)

            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(1).Editable = True
            oGrid.Columns.Item(2).Visible = False

            Return 0

        Catch ex As Exception

            MsgBox("FrmtekDel. fallo la carga previa de la forma AgregarLineas: " & ex.Message)

        Finally

            oGrid = Nothing

        End Try

    End Function


    Public Sub HideOrShowFormItems(ByVal DocNum As String, ByVal psFormUID As String)
        Dim loItem As SAPbouiCOM.Item
        Dim loButton As SAPbouiCOM.Button
        Dim stQuery As String = ""
        Dim oRecSet As SAPbobsCOM.Recordset

        Try

            coForm = cSBOApplication.Forms.Item(psFormUID)
            oRecSet = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuery = "call ShowButtonPriceDelivery('" & DocNum & "')"
            oRecSet.DoQuery(stQuery)

            If oRecSet.RecordCount = 0 Then

                loButton = coForm.Items.Item("3").Specific
                loButton.Item.Enabled = False

            Else

                loButton = coForm.Items.Item("3").Specific
                loButton.Item.Enabled = True

            End If

        Catch ex As Exception
            cSBOApplication.MessageBox("DocumentoSBO. agregar elementos a la forma. " & ex.Message)
        Finally
            loItem = Nothing
            loButton = Nothing
        End Try
    End Sub


End Class
