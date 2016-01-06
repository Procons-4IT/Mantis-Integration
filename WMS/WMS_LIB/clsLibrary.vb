Imports System.Configuration.ConfigurationManager
Imports System.IO
Imports System.Data.SqlClient
Imports SAPbobsCOM

Public Class clsLibrary

    Private oCompany As SAPbobsCOM.Company
    Dim myConnection As SqlConnection
    Dim oCommand As SqlCommand
    Dim oSqlAdap As SqlDataAdapter
    Dim oHeaderDt As DataTable
    Dim oDetailDt As DataTable
    Dim oLineDt_C As DataTable
    Dim sQuery As String = String.Empty

    Public Function mainFunction() As Boolean
        Dim _retVal As Boolean = True
        Try
            connectSAP()
            If Not IsNothing(oCompany) Then
                If oCompany.Connected Then
                    'import_Delivery()
                    'import_Return()
                    'import_Purchase()
                    'import_Transfer()
                    'import_InventoryReceipt()
                    'import_InventoryIssue()
                    import_InventoryCounting()
                End If
            End If
            'disconnectSAP()
            Return _retVal
        Catch ex As Exception
            traceService(ex.ToString())
        Finally
            disconnectSAP()
        End Try
        Return _retVal
    End Function

    Private Sub connectSAP()
        Try
            traceService("Company Connection")

            oCompany = New SAPbobsCOM.Company()
            Dim strCompany As String = System.Configuration.ConfigurationManager.AppSettings("CompanyDB").ToString
            oCompany.CompanyDB = strCompany

            oCompany.UserName = System.Configuration.ConfigurationManager.AppSettings("UserName").ToString
            oCompany.Password = System.Configuration.ConfigurationManager.AppSettings("Password").ToString
            oCompany.Server = System.Configuration.ConfigurationManager.AppSettings("Server").ToString
            oCompany.DbUserName = System.Configuration.ConfigurationManager.AppSettings("DbUserName").ToString
            oCompany.DbPassword = System.Configuration.ConfigurationManager.AppSettings("DbPassword").ToString

            If System.Configuration.ConfigurationManager.AppSettings("DBType").ToString = "2008" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf System.Configuration.ConfigurationManager.AppSettings("DBType").ToString = "2012" Then
                'oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            End If

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            oCompany.LicenseServer = System.Configuration.ConfigurationManager.AppSettings("LicenseServer").ToString

            If (0 <> oCompany.Connect()) Then
                traceService("Company Connection Failed")
                traceService(oCompany.GetLastErrorDescription())
                Exit Sub
            Else
                traceService("Connected successfully")
            End If

        Catch ex As Exception
            traceService(ex.ToString())
        End Try
    End Sub

    Private Sub disconnectSAP()
        Try
            If Not IsNothing(oCompany) Then
                If oCompany.Connected Then
                    oCompany.Disconnect()
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try
    End Sub

    Private Sub traceService(ByVal content As String)
        Try
            Dim strFile As String = "\WMS_SERVICE_" + System.DateTime.Now.ToString("yyyyMMdd") + ".txt"
            Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() & strFile

            If Not File.Exists(strPath) Then
                Dim fs As New FileStream(strPath, FileMode.Create, FileAccess.Write)
                Dim sw As New StreamWriter(fs)
                sw.BaseStream.Seek(0, SeekOrigin.[End])
                sw.WriteLine(content)
                sw.Flush()
                sw.Close()
            Else
                Dim fs As New FileStream(strPath, FileMode.Append, FileAccess.Write)
                Dim sw As New StreamWriter(fs)
                sw.BaseStream.Seek(0, SeekOrigin.[End])
                sw.WriteLine(content)
                sw.Flush()
                sw.Close()
            End If

        Catch ex As Exception
            'traceService(ex.ToString)
        End Try

    End Sub

    Public Function ExecuteDataSet(ByVal strQuery As String) As DataSet
        Dim _retVal As New DataSet()
        Dim ConnectionString As String = System.Configuration.ConfigurationManager.AppSettings("ConnectionString")
        myConnection = New SqlConnection(ConnectionString)
        Try
            myConnection.Open()
            oCommand = New SqlCommand()
            If myConnection.State = ConnectionState.Open Then
                oCommand.Connection = myConnection
                oCommand.CommandText = strQuery
                oCommand.CommandType = CommandType.Text
                oSqlAdap = New SqlDataAdapter(oCommand)
                oSqlAdap.Fill(_retVal)
            Else
                Throw New Exception("")
            End If
        Catch ex As Exception
            Throw ex
        Finally
            myConnection.Close()
            oCommand = Nothing
            oSqlAdap = Nothing
        End Try
        Return _retVal
    End Function

    Public Sub ExecuteNonQuery(ByVal str As String)
        Dim ConnectionString As String = System.Configuration.ConfigurationManager.AppSettings("ConnectionString")
        Dim oCommand As SqlCommand
        myConnection = New SqlConnection(ConnectionString)

        Try
            myConnection.Open()
            oCommand = New SqlCommand(str, myConnection)
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            myConnection.Close()
        End Try
    End Sub

    Private Sub import_Delivery()
        Try
            Dim oDelivery As SAPbobsCOM.Documents
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select T0.lvsboh_DocNum,T0.lvsboh_CustomerCode,T0.lvsboh_DocDate,T0.lvsboh_DocDueDate,T0.lvsboh_PostingDate, " & _
                        " T0.lvsboh_DocReference, ISNULL(T0.lvsboh_Remarks,'') As lvsboh_Remarks,  T0.lvsboh_Warehouse " & _
                        " From LVSBODocumentHeader T0 " & _
                        " Where T0.lvsboh_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oDelivery = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                        DocNum = Hrow("lvsboh_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oDelivery.CardCode = Hrow("lvsboh_CustomerCode").ToString().Trim()
                            oDelivery.DocDate = Hrow("lvsboh_DocDate")
                            oDelivery.TaxDate = Hrow("lvsboh_PostingDate")
                            oDelivery.DocDueDate = Hrow("lvsboh_DocDueDate")
                            oDelivery.NumAtCard = Hrow("lvsboh_DocReference")
                            oDelivery.Comments = Hrow("lvsboh_Remarks")

                            sQuery = " Select T1.lvsbol_ItemCode, T1.lvsbol_Quantity, T1.lvsbol_Price, T1.lvsbol_Discount, " & _
                                        " T1.lvsbol_BaseType,T1.lvsbol_BaseRef,T1.lvsbol_BaseLine " & _
                                        " From LVSBODocumentLines T1 " & _
                                        " Where T1.lvsbol_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True
                                        If intRow > 0 Then
                                            oDelivery.Lines.Add()
                                        End If
                                        oDelivery.Lines.SetCurrentLine(intRow)

                                        oDelivery.Lines.ItemCode = Drow("lvsbol_ItemCode").ToString().Trim()
                                        oDelivery.Lines.Quantity = Drow("lvsbol_Quantity")
                                        oDelivery.Lines.UnitPrice = Drow("lvsbol_Price")
                                        oDelivery.Lines.DiscountPercent = Drow("lvsbol_Discount")
                                        oDelivery.Lines.WarehouseCode = Hrow("lvsboh_Warehouse").ToString().Trim()
                                        If Not IsDBNull(Drow("lvsbol_BaseRef")) Then
                                            If Drow("lvsbol_BaseRef").ToString() <> "" Then
                                                oDelivery.Lines.BaseType = 17
                                                oDelivery.Lines.BaseEntry = Drow("lvsbol_BaseRef")
                                                oDelivery.Lines.BaseLine = Drow("lvsbol_BaseLine")
                                            End If
                                        End If
                                        intRow = intRow + 1
                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oDelivery.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVSBODocumentHeader" & _
                                               " Set lvsboh_Status = 1 " & _
                                               " Where lvsboh_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)
                                    If 1 = 2 Then
                                        'closeOpenOrders(DocNum)
                                    Else
                                        closeOpenOrders_Lines(DocNum)
                                    End If
                                Else
                                    'sQuery = "Update LVSBODocumentHeader" & _
                                    '           " Set lvsboh_SapMessage = '" & oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription() & "' " & _
                                    '           " Where lvsboh_DocNum = '" & DocNum & "'"
                                    'ExecuteNonQuery(sQuery)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            'sQuery = "Update LVSBODocumentHeader" & _
                            '                   " Set lvsboh_SapMessage = '" & ex.Message & "' " & _
                            '                   " Where lvsboh_DocNum = '" & DocNum & "'"
                            'ExecuteNonQuery(sQuery)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Public Function closeOpenOrders_Lines(ByVal strDocNum As String) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim ds As DataSet
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oOrder As SAPbobsCOM.Documents
            Dim intStatus As Integer

            sQuery = " Select T1.lvsbol_BaseRef,T1.lvsbol_BaseLine,T1.lvsbol_IsClosing " & _
                                       " From LVSBODocumentHeader T0 " & _
                                       " JOIN LVSBODocumentLines T1 On T0.lvsboh_DocNum = T1.lvsbol_DocNum " & _
                                       " Where T0.lvsboh_DocNum = '" & strDocNum & "' "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oLineDt_C = ds.Tables(0)
                If oLineDt_C.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oLineDt_C.Rows

                        sQuery = "Update LVSBODocumentLines" & _
                                              " Set lvsbol_Status = 1 " & _
                                              " Where lvsbol_DocNum = '" & strDocNum & "'" & _
                                              " And lvsbol_LineNo = '" & Hrow("lvsbol_BaseLine") & "'"
                        ExecuteNonQuery(sQuery)

                        If Not IsDBNull(Hrow("lvsbol_IsClosing")) Then
                            If Hrow("lvsbol_IsClosing").ToString() = "Y" Then
                                sQuery = "Select DocEntry,VisOrder From RDR1 Where DocEntry = '" & Hrow("lvsbol_BaseRef") & "' And LineStatus = 'O' " & _
                                    " And LineNum = '" & Hrow("lvsbol_BaseLine") & "'"
                                oRecordSet.DoQuery(sQuery)
                                If Not oRecordSet.EoF Then
                                    oOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                                    If oOrder.GetByKey(CInt(oRecordSet.Fields.Item("DocEntry").Value)) Then
                                        oOrder.Lines.SetCurrentLine(CInt(oRecordSet.Fields.Item("VisOrder").Value))
                                        If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                            oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                            intStatus = oOrder.Update()
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub import_Return()
        Try
            Dim oReturn As SAPbobsCOM.Documents
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select T0.lvrboh_DocNum,T0.lvrboh_CustomerCode,T0.lvrboh_DocDate,T0.lvrboh_DocDueDate,T0.lvrboh_PostingDate, " & _
                        " T0.lvrboh_DocReference, ISNULL(T0.lvrboh_Remarks,'') As lvrboh_Remarks,  T0.lvrboh_Warehouse " & _
                        " From LVRBODocumentHeader T0 " & _
                        " Where T0.lvrboh_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oReturn = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns)
                        DocNum = Hrow("lvrboh_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oReturn.CardCode = Hrow("lvrboh_CustomerCode").ToString().Trim()
                            oReturn.DocDate = Hrow("lvrboh_DocDate")
                            oReturn.TaxDate = Hrow("lvrboh_PostingDate")
                            oReturn.DocDueDate = Hrow("lvrboh_DocDueDate")
                            oReturn.NumAtCard = Hrow("lvrboh_DocReference")
                            oReturn.Comments = Hrow("lvrboh_Remarks")

                            sQuery = " Select T1.lvrbol_ItemCode, T1.lvrbol_Quantity, T1.lvrbol_Price, T1.lvrbol_Discount, " & _
                                        " T1.lvrbol_BaseType,T1.lvrbol_BaseRef,T1.lvrbol_BaseLine " & _
                                        " From LVRBODocumentLines T1 " & _
                                        " Where T1.lvrbol_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True
                                        If intRow > 0 Then
                                            oReturn.Lines.Add()
                                        End If
                                        oReturn.Lines.SetCurrentLine(intRow)

                                        oReturn.Lines.ItemCode = Drow("lvrbol_ItemCode").ToString().Trim()
                                        oReturn.Lines.Quantity = Drow("lvrbol_Quantity")
                                        oReturn.Lines.UnitPrice = Drow("lvrbol_Price")
                                        oReturn.Lines.DiscountPercent = Drow("lvrbol_Discount")
                                        oReturn.Lines.WarehouseCode = Hrow("lvrboh_Warehouse").ToString().Trim()
                                        If Not IsDBNull(Drow("lvrbol_BaseRef")) Then
                                            If Drow("lvrbol_BaseRef").ToString() <> "" Then
                                                oReturn.Lines.BaseType = 15
                                                oReturn.Lines.BaseEntry = Drow("lvrbol_BaseRef")
                                                oReturn.Lines.BaseLine = Drow("lvrbol_BaseLine")
                                            End If
                                        End If
                                        intRow = intRow + 1
                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oReturn.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVRBODocumentHeader" & _
                                               " Set lvrboh_Status = 1 " & _
                                               " Where lvrboh_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)
                                    'Close Delivery Lines Based on Flag...
                                    'closeOpenOrders_Lines(DocNum)
                                Else
                                    'sQuery = "Update LVRBODocumentHeader" & _
                                    '           " Set lvrboh_SapMessage = '" & oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription() & "' " & _
                                    '           " Where lvrboh_DocNum = '" & DocNum & "'"
                                    'ExecuteNonQuery(sQuery)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            'sQuery = "Update LVSBODocumentHeader" & _
                            '                   " Set lvsboh_SapMessage = '" & ex.Message & "' " & _
                            '                   " Where lvsboh_DocNum = '" & DocNum & "'"
                            'ExecuteNonQuery(sQuery)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Private Sub import_Purchase()
        Try
            Dim oGRPO As SAPbobsCOM.Documents
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select T0.lvpboh_DocNum,T0.lvpboh_CustomerCode,T0.lvpboh_DocDate,T0.lvpboh_DocDueDate,T0.lvpboh_PostingDate, " & _
                        " T0.lvpboh_DocReference, ISNULL(T0.lvpboh_Remarks,'') As lvpboh_Remarks,  T0.lvpboh_Warehouse " & _
                        " From LVPBODocumentHeader T0 " & _
                        " Where T0.lvpboh_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oGRPO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                        DocNum = Hrow("lvpboh_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oGRPO.CardCode = Hrow("lvpboh_CustomerCode").ToString().Trim()
                            oGRPO.DocDate = Hrow("lvpboh_DocDate")
                            oGRPO.TaxDate = Hrow("lvpboh_PostingDate")
                            oGRPO.DocDueDate = Hrow("lvpboh_DocDueDate")
                            oGRPO.NumAtCard = Hrow("lvpboh_DocReference")
                            oGRPO.Comments = Hrow("lvpboh_Remarks")

                            sQuery = " Select T1.lvpbol_ItemCode, T1.lvpbol_Quantity, T1.lvpbol_Price, T1.lvpbol_Discount, " & _
                                        " T1.lvpbol_BaseType,T1.lvpbol_BaseRef,T1.lvpbol_BaseLine " & _
                                        " From LVPBODocumentLines T1 " & _
                                        " Where T1.lvpbol_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True
                                        If intRow > 0 Then
                                            oGRPO.Lines.Add()
                                        End If
                                        oGRPO.Lines.SetCurrentLine(intRow)

                                        oGRPO.Lines.ItemCode = Drow("lvpbol_ItemCode").ToString().Trim()
                                        oGRPO.Lines.Quantity = Drow("lvpbol_Quantity")
                                        oGRPO.Lines.UnitPrice = Drow("lvpbol_Price")
                                        oGRPO.Lines.DiscountPercent = Drow("lvpbol_Discount")
                                        oGRPO.Lines.WarehouseCode = Hrow("lvpboh_Warehouse").ToString().Trim()
                                        If Not IsDBNull(Drow("lvpbol_BaseRef")) Then
                                            If Drow("lvpbol_BaseRef").ToString() <> "" Then
                                                oGRPO.Lines.BaseType = 22
                                                oGRPO.Lines.BaseEntry = Drow("lvpbol_BaseRef")
                                                oGRPO.Lines.BaseLine = Drow("lvpbol_BaseLine")
                                            End If
                                        End If
                                        intRow = intRow + 1
                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oGRPO.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVPBODocumentHeader" & _
                                               " Set lvpboh_Status = 1 " & _
                                               " Where lvpboh_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)
                                    closeOpenPurchaseOrders_Lines(DocNum) ' Close Purchase Lines.
                                Else
                                    'sQuery = "Update LVSBODocumentHeader" & _
                                    '           " Set lvpboh_SapMessage = '" & oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription() & "' " & _
                                    '           " Where lvpboh_DocNum = '" & DocNum & "'"
                                    'ExecuteNonQuery(sQuery)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            'sQuery = "Update LVSBODocumentHeader" & _
                            '                   " Set lvpboh_SapMessage = '" & ex.Message & "' " & _
                            '                   " Where lvpboh_DocNum = '" & DocNum & "'"
                            'ExecuteNonQuery(sQuery)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Public Function closeOpenPurchaseOrders_Lines(ByVal strDocNum As String) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim ds As DataSet
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oOrder As SAPbobsCOM.Documents
            Dim intStatus As Integer

            sQuery = " Select T1.lvpbol_BaseRef,T1.lvpbol_BaseLine,T1.lvpbol_IsClosing " & _
                                       " From LVPBODocumentHeader T0 " & _
                                       " JOIN LVPBODocumentLines T1 On T0.lvpboh_DocNum = T1.lvpbol_DocNum " & _
                                       " Where T0.lvpboh_DocNum = '" & strDocNum & "' "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oLineDt_C = ds.Tables(0)
                If oLineDt_C.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oLineDt_C.Rows

                        sQuery = "Update LVPBODocumentLines" & _
                                              " Set lvpbol_Status = 1 " & _
                                              " Where lvpbol_DocNum = '" & strDocNum & "'" & _
                                              " And lvpbol_LineNo = '" & Hrow("lvpbol_BaseLine") & "'"
                        ExecuteNonQuery(sQuery)

                        If Not IsDBNull(Hrow("lvpbol_IsClosing")) Then
                            If Hrow("lvpbol_IsClosing").ToString() = "Y" Then
                                sQuery = "Select DocEntry,VisOrder From POR1 Where DocEntry = '" & Hrow("lvpbol_BaseRef") & "' And LineStatus = 'O' " & _
                                    " And LineNum = '" & Hrow("lvpbol_BaseLine") & "'"
                                oRecordSet.DoQuery(sQuery)
                                If Not oRecordSet.EoF Then
                                    oOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                    If oOrder.GetByKey(CInt(oRecordSet.Fields.Item("DocEntry").Value)) Then
                                        oOrder.Lines.SetCurrentLine(CInt(oRecordSet.Fields.Item("VisOrder").Value))
                                        If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                            oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                            intStatus = oOrder.Update()
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub import_Transfer()
        Try
            Dim oTransfer As SAPbobsCOM.StockTransfer
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select Top 1 T0.lvtboh_DocNum,T0.lvtboh_DocDate,T0.lvtboh_DocDueDate,T0.lvtboh_PostingDate, " & _
                        " T0.lvtboh_DocReference, ISNULL(T0.lvtboh_Remarks,'') As lvtboh_Remarks,  T0.lvtboh_FWarehouse,  T0.lvtboh_TWarehouse " & _
                        " From LVTBODocumentHeader T0 " & _
                        " Where T0.lvtboh_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oTransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        DocNum = Hrow("lvtboh_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            ' oTransfer.CardCode = Hrow("lvpboh_CustomerCode").ToString().Trim()
                            oTransfer.DocDate = Hrow("lvtboh_DocDate")
                            oTransfer.TaxDate = Hrow("lvtboh_PostingDate")
                            'oTransfer.DocDueDate = Hrow("lvtboh_DocDueDate")
                            oTransfer.FromWarehouse = Hrow("lvtboh_FWarehouse")
                            'oTransfer.ToWarehouse = Hrow("lvtboh_TWarehouse")
                            'oTransfer.NumAtCard = Hrow("lvpboh_DocReference")
                            oTransfer.Comments = Hrow("lvtboh_Remarks")

                            sQuery = " Select T1.lvtbol_ItemCode, T1.lvtbol_Quantity, T1.lvtbol_Price, T1.lvtbol_Discount, " & _
                                        " T1.lvtbol_BaseType,T1.lvtbol_BaseRef,T1.lvtbol_BaseLine " & _
                                        " From LVTBODocumentLines T1 " & _
                                        " Where T1.lvtbol_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True

                                        If intRow > 0 Then
                                            oTransfer.Lines.Add()
                                        End If

                                        oTransfer.Lines.SetCurrentLine(intRow)

                                        oTransfer.Lines.ItemCode = Drow("lvtbol_ItemCode").ToString().Trim()
                                        oTransfer.Lines.Quantity = Drow("lvtbol_Quantity")
                                        oTransfer.Lines.UnitPrice = Drow("lvtbol_Price")
                                        oTransfer.Lines.DiscountPercent = Drow("lvtbol_Discount")
                                        oTransfer.Lines.WarehouseCode = Hrow("lvtboh_TWarehouse").ToString().Trim()
                                        If Not IsDBNull(Drow("lvtbol_BaseRef")) Then
                                            If Drow("lvtbol_BaseRef").ToString() <> "" Then
                                                oTransfer.Lines.BaseType = SAPbobsCOM.InvBaseDocTypeEnum.InventoryTransferRequest
                                                oTransfer.Lines.BaseEntry = Drow("lvtbol_BaseRef")
                                                oTransfer.Lines.BaseLine = Drow("lvtbol_BaseLine")
                                            End If
                                        End If

                                        intRow = intRow + 1

                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oTransfer.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVTBODocumentHeader" & _
                                               " Set lvtboh_Status = 1 " & _
                                               " Where lvtboh_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                    closeOpenInventoryTransfer_Lines(DocNum)

                                Else
                                    'sQuery = "Update LVSBODocumentHeader" & _
                                    '           " Set lvpboh_SapMessage = '" & oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription() & "' " & _
                                    '           " Where lvpboh_DocNum = '" & DocNum & "'"
                                    'ExecuteNonQuery(sQuery)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            'sQuery = "Update LVSBODocumentHeader" & _
                            '                   " Set lvpboh_SapMessage = '" & ex.Message & "' " & _
                            '                   " Where lvpboh_DocNum = '" & DocNum & "'"
                            'ExecuteNonQuery(sQuery)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Public Function closeOpenInventoryTransfer_Lines(ByVal strDocNum As String) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim ds As DataSet
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            'Dim oOrder As SAPbobsCOM.Documents
            'Dim intStatus As Integer

            sQuery = " Select T1.lvtbol_BaseRef,T1.lvtbol_BaseLine,T1.lvtbol_IsClosing " & _
                                       " From LVTBODocumentHeader T0 " & _
                                       " JOIN LVTBODocumentLines T1 On T0.lvtboh_DocNum = T1.lvtbol_DocNum " & _
                                       " Where T0.lvtboh_DocNum = '" & strDocNum & "' "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oLineDt_C = ds.Tables(0)
                If oLineDt_C.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oLineDt_C.Rows

                        sQuery = "Update LVTBODocumentLines" & _
                                              " Set lvtbol_Status = 1 " & _
                                              " Where lvtbol_DocNum = '" & strDocNum & "'" & _
                                              " And lvtbol_LineNo = '" & Hrow("lvtbol_BaseLine") & "'"
                        ExecuteNonQuery(sQuery)

                        'If Not IsDBNull(Hrow("lvpbol_IsClosing")) Then
                        '    If Hrow("lvpbol_IsClosing").ToString() = "Y" Then
                        '        sQuery = "Select DocEntry,VisOrder From POR1 Where DocEntry = '" & Hrow("lvpbol_BaseRef") & "' And LineStatus = 'O' " & _
                        '            " And LineNum = '" & Hrow("lvpbol_BaseLine") & "'"
                        '        oRecordSet.DoQuery(sQuery)
                        '        If Not oRecordSet.EoF Then
                        '            oOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                        '            If oOrder.GetByKey(CInt(oRecordSet.Fields.Item("DocEntry").Value)) Then
                        '                oOrder.Lines.SetCurrentLine(CInt(oRecordSet.Fields.Item("VisOrder").Value))
                        '                If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                        '                    oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                        '                    intStatus = oOrder.Update()
                        '                End If
                        '            End If
                        '        End If
                        '    End If
                        'End If
                    Next
                End If
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub import_InventoryReceipt()
        Try
            Dim oIReceipt As SAPbobsCOM.Documents
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select Top 1 T0.lvibo_DocNum,T0.lvibo_DocDate,T0.lvibo_DocDueDate,T0.lvibo_PostingDate, " & _
                        " T0.lvibo_DocReference, ISNULL(T0.lvibo_Remarks,'') As lvibo_Remarks,  T0.lvibo_Warehouse " & _
                        " From LVIBODocument T0 " & _
                        " Where T0.lvibo_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oIReceipt = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                        DocNum = Hrow("lvibo_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oIReceipt.DocDate = Hrow("lvibo_DocDate")
                            oIReceipt.TaxDate = Hrow("lvibo_PostingDate")
                            oIReceipt.Comments = Hrow("lvibo_Remarks")

                            sQuery = " Select T1.lvibo_ItemCode, T1.lvibo_Quantity, T1.lvibo_Price, T1.lvibo_Discount " & _
                                        " From LVIBODocument T1 " & _
                                        " Where T1.lvibo_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True

                                        If intRow > 0 Then
                                            oIReceipt.Lines.Add()
                                        End If

                                        oIReceipt.Lines.SetCurrentLine(intRow)

                                        oIReceipt.Lines.ItemCode = Drow("lvibo_ItemCode").ToString().Trim()
                                        oIReceipt.Lines.Quantity = Drow("lvibo_Quantity")
                                        oIReceipt.Lines.UnitPrice = Drow("lvibo_Price")
                                        oIReceipt.Lines.DiscountPercent = Drow("lvibo_Discount")
                                        oIReceipt.Lines.WarehouseCode = Hrow("lvibo_Warehouse").ToString().Trim()
                                        

                                        intRow = intRow + 1

                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oIReceipt.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVIBODocument" & _
                                               " Set lvibo_Status = 1 " & _
                                               " Where lvibo_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                Else
                                    'sQuery = "Update LVSBODocumentHeader" & _
                                    '           " Set lvpboh_SapMessage = '" & oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription() & "' " & _
                                    '           " Where lvpboh_DocNum = '" & DocNum & "'"
                                    'ExecuteNonQuery(sQuery)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            'sQuery = "Update LVSBODocumentHeader" & _
                            '                   " Set lvpboh_SapMessage = '" & ex.Message & "' " & _
                            '                   " Where lvpboh_DocNum = '" & DocNum & "'"
                            'ExecuteNonQuery(sQuery)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Private Sub import_InventoryIssue()
        Try
            Dim oIIssue As SAPbobsCOM.Documents
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select Top 1 T0.lvobo_DocNum,T0.lvobo_DocDate,T0.lvobo_DocDueDate,T0.lvobo_PostingDate, " & _
                        " T0.lvobo_DocReference, ISNULL(T0.lvobo_Remarks,'') As lvobo_Remarks,  T0.lvobo_Warehouse " & _
                        " From LVOBODocument T0 " & _
                        " Where T0.lvobo_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oIIssue = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        DocNum = Hrow("lvobo_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oIIssue.DocDate = Hrow("lvobo_DocDate")
                            oIIssue.TaxDate = Hrow("lvobo_PostingDate")
                            oIIssue.Comments = Hrow("lvobo_Remarks")

                            sQuery = " Select T1.lvobo_ItemCode, T1.lvobo_Quantity, T1.lvobo_Price, T1.lvobo_Discount " & _
                                        " From LVOBODocument T1 " & _
                                        " Where T1.lvobo_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True

                                        If intRow > 0 Then
                                            oIIssue.Lines.Add()
                                        End If

                                        oIIssue.Lines.SetCurrentLine(intRow)

                                        oIIssue.Lines.ItemCode = Drow("lvobo_ItemCode").ToString().Trim()
                                        oIIssue.Lines.Quantity = Drow("lvobo_Quantity")
                                        oIIssue.Lines.UnitPrice = Drow("lvobo_Price")
                                        oIIssue.Lines.DiscountPercent = Drow("lvobo_Discount")
                                        oIIssue.Lines.WarehouseCode = Hrow("lvobo_Warehouse").ToString().Trim()


                                        intRow = intRow + 1

                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oIIssue.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVOBODocument" & _
                                               " Set lvobo_Status = 1 " & _
                                               " Where lvobo_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                Else
                                    'sQuery = "Update LVSBODocumentHeader" & _
                                    '           " Set lvpboh_SapMessage = '" & oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription() & "' " & _
                                    '           " Where lvpboh_DocNum = '" & DocNum & "'"
                                    'ExecuteNonQuery(sQuery)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            'sQuery = "Update LVSBODocumentHeader" & _
                            '                   " Set lvpboh_SapMessage = '" & ex.Message & "' " & _
                            '                   " Where lvpboh_DocNum = '" & DocNum & "'"
                            'ExecuteNonQuery(sQuery)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Private Sub import_InventoryCounting()
        Try
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select Top 1 T0.lvcbo_DocNum,T0.lvcbo_CountDate, " & _
                        " T0.lvcbo_DocReference, ISNULL(T0.lvcbo_Remarks,'') As lvcbo_Remarks,T0.lvcbo_Warehouse " & _
                        " From LVCBODocument T0 " & _
                        " Where T0.lvcbo_Status = 0 "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        Dim oCS As SAPbobsCOM.CompanyService = oCompany.GetCompanyService()
                        Dim oICS As SAPbobsCOM.InventoryCountingsService = oCS.GetBusinessService(ServiceTypes.InventoryCountingsService)
                        Dim oIC As SAPbobsCOM.InventoryCounting = oICS.GetDataInterface(InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
                        DocNum = Hrow("lvcbo_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oIC.CountDate = Hrow("lvcbo_CountDate")
                            oIC.SingleCounterType = CounterTypeEnum.ctUser
                            oIC.SingleCounterID = oCompany.UserSignature

                            sQuery = " Select T1.lvcbo_ItemCode, T1.lvcbo_Quantity,T1.lvcbo_UOMEntry,T1.lvcbo_UOMName " & _
                                        " From LVCBODocument T1 " & _
                                        " Where T1.lvcbo_DocNum = '" & DocNum & "' "
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0

                                    Dim oICLS As SAPbobsCOM.InventoryCountingLines = oIC.InventoryCountingLines

                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True
                                        Dim oICL As SAPbobsCOM.InventoryCountingLine = oICLS.Add

                                        'If intRow > 0 Then
                                        '    oIIssue.Lines.Add()
                                        'End If

                                        ' oIIssue.Lines.SetCurrentLine(intRow)
                                        oICL.ItemCode = Drow("lvcbo_ItemCode").ToString().Trim()
                                        oICL.CountedQuantity = Drow("lvcbo_Quantity")
                                        oICL.Counted = BoYesNoEnum.tYES
                                        oICL.CountedQuantity = Drow("lvcbo_Quantity")
                                        oICL.WarehouseCode = Hrow("lvcbo_Warehouse").ToString().Trim()
                                        oICL.UoMCode = Drow("lvcbo_UOMEntry")


                                        'intRow = intRow + 1

                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim oICP As SAPbobsCOM.InventoryCountingParams = oICS.Add(oIC)

                                If oICP.DocumentEntry > 0 Then
                                    sQuery = "Update LVCBODocument" & _
                                               " Set lvcbo_Status = 1 " & _
                                               " Where lvcbo_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                Else
                                    'sQuery = "Update LVSBODocumentHeader" & _
                                    '           " Set lvpboh_SapMessage = '" & oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription() & "' " & _
                                    '           " Where lvpboh_DocNum = '" & DocNum & "'"
                                    'ExecuteNonQuery(sQuery)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            'sQuery = "Update LVSBODocumentHeader" & _
                            '                   " Set lvpboh_SapMessage = '" & ex.Message & "' " & _
                            '                   " Where lvpboh_DocNum = '" & DocNum & "'"
                            'ExecuteNonQuery(sQuery)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    'Public Sub CreateInvCounting()


    '    Try

    '        oIC.CountDate = DateTime.Now
    '        Dim oICLS As SAPbobsCOM.InventoryCountingLines = oIC.InventoryCountingLines
    '        Dim oICL As SAPbobsCOM.InventoryCountingLine = oICLS.Add
    '        oICL.ItemCode = "00111520102Y"
    '        oICL.CountedQuantity = 4
    '        oICL.WarehouseCode = "1"
    '        oICL.Counted = BoYesNoEnum.tYES
    '        Dim oICP As SAPbobsCOM.InventoryCountingParams = oICS.Add(oIC)
    '            Dim DocEntry As Integer = = oICP.DocumentEntry  
    '    Catch Ex As Exception
    '        MsgBox(Ex.Message)
    '    End Try
    'End Sub

End Class


'Public Function closeOpenOrders(ByVal strDocNum As String) As Boolean
'    Try
'        Dim _retVal As Boolean = False
'        Dim intDocEntry As Integer
'        Dim oRecordSet As SAPbobsCOM.Recordset
'        oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
'        Dim oOrder As SAPbobsCOM.Documents
'        Dim intStatus As Integer
'        sQuery = "Select Distinct DocEntry From ORDR Where DocNum = '" & strDocNum & "' And DocStatus = 'O' "
'        oRecordSet.DoQuery(sQuery)
'        If Not oRecordSet.EoF Then
'            While Not oRecordSet.EoF
'                Dim blnClose As Boolean = False
'                oOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
'                intDocEntry = CInt(oRecordSet.Fields.Item(0).Value)
'                If oOrder.GetByKey(intDocEntry) Then
'                    If oOrder.DocumentStatus = BoStatus.bost_Open Then
'                        blnClose = True
'                    End If
'                    If blnClose Then
'                        intStatus = oOrder.Close()
'                    End If
'                End If
'                oRecordSet.MoveNext()
'            End While
'        End If
'        _retVal = True
'        Return _retVal
'    Catch ex As Exception
'        Throw ex
'    End Try
'End Function