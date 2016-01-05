Imports System.IO
Public Class clsInvoice
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oStatic As SAPbouiCOM.StaticText
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public LocalDb, strDocEntry, Server, UID, ServerUid, ServerPwd, servertype, serverdb, str As String
    Dim strSelectedFilepath, sPath As String

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.linkedserverConnect() = True Then
            oForm = oApplication.Utilities.LoadForm(xml_Invoice, frm_Invoice)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            Databind_load(oForm)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            oForm.Freeze(False)
        End If
    End Sub
#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        agrid.Columns.Item(0).TitleObject.Caption = "Select"
        agrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item(0).Visible = False
        agrid.Columns.Item(1).TitleObject.Caption = "Invoice Sequence"
        agrid.Columns.Item(2).TitleObject.Caption = "customer code"
        agrid.Columns.Item(3).TitleObject.Caption = "Transaction Date"
        agrid.Columns.Item(4).TitleObject.Caption = "Quantity"
        agrid.Columns.Item(5).TitleObject.Caption = "Item code"
        agrid.Columns.Item(6).TitleObject.Caption = "Flag"
        agrid.Columns.Item(6).Visible = False
        agrid.Columns.Item(7).TitleObject.Caption = "Reference No"
        agrid.Columns.Item(8).TitleObject.Caption = "Remarks"
        agrid.Columns.Item(9).TitleObject.Caption = "Price"
        agrid.Columns.Item(10).TitleObject.Caption = "Customer Type"
        agrid.Columns.Item(11).TitleObject.Caption = "Agent Code"
        agrid.Columns.Item(12).TitleObject.Caption = "Code"
        agrid.Columns.Item(12).Visible = False
        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region
#Region "Databind"
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim PcvNofrom, PcvNoto, str1, strVoucherCondition As String
            Dim strVoucherNumber As String = ""
            Dim OGrid As SAPbouiCOM.Grid
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  "
            LoginDetails()
            linkedserverConnect()
            OGrid = aform.Items.Item("7").Specific
            dtTemp = OGrid.DataTable
            Dim dtFrom, dtTo As Date
            Dim strFrom, strTo As String
            strFrom = oApplication.Utilities.getEdittextvalue(aform, "5")
            strTo = oApplication.Utilities.getEdittextvalue(aform, "8")

            If oApplication.Utilities.getEdittextvalue(aform, "5") <> "" Then
                dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "5"))
            End If
            If strTo <> "" Then
                dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
            End If
            Dim strCondition As String
            'If strFrom <> "" Then
            '    strCondition = "convert(varchar(10),[InvDate],105) >='" & dtFrom.ToString("dd-MM-yyyy") & "'"
            'Else
            '    strCondition = " 1 =1"
            'End If
            'If strTo <> "" Then
            '    strCondition = strCondition & " and convert(varchar(10),[InvDate],105) <='" & dtTo.ToString("dd-MM-yyyy") & "'"
            'Else
            '    strCondition = strCondition & " and 2=2"
            'End If

            If strFrom <> "" Then
                '  2013-01-09 00:00:00.000
                strCondition = "[InvDate] >='" & dtFrom.ToString("yyyy-MM-dd") & " 00:00:00.000'"
            Else
                strCondition = " 1 =1"
            End If
            If strTo <> "" Then
                strCondition = strCondition & " and [InvDate] <='" & dtTo.ToString("yyyy-MM-dd") & " 23:59:00.000'"
            Else
                strCondition = strCondition & " and 2=2"
            End If
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  "
            str1 = "Select '',[InvSeq],[CustCode],[InvDate],[Quantity],[ItemNo],[Flag],[RefNo],[Remarks],[Price],[Invtype],[AgentCode],PK from  " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False' and " & strCondition
            dtTemp.ExecuteQuery(str1)
            OGrid.DataTable = dtTemp
            Formatgrid(OGrid)
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "  "
            aform.Freeze(False)
            linkedserverDisconnect()
        Catch ex As Exception
            linkedserverDisconnect()
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try

    End Sub

    Private Function Valiation_Table(ByVal aform As SAPbouiCOM.Form) As Boolean
        Try
            aform.Freeze(True)
            Dim PcvNofrom, PcvNoto, str1, strVoucherCondition As String
            Dim strVoucherNumber As String = ""
            Dim OGrid As SAPbouiCOM.Grid
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  "
            LoginDetails()
            linkedserverConnect()
            ' OGrid = aform.Items.Item("7").Specific
            ' dtTemp = OGrid.DataTable
            Dim dtFrom, dtTo As Date
            Dim strFrom, strTo As String
            Dim strGridGLCode, strAcctcode, Strselect, strcustCode, strItemCode As String
            Dim oTemp, oTemp1 As SAPbobsCOM.Recordset
            Dim blnRecordselected As Boolean = False
            Dim strCustomer, strItem As ArrayList
            strCustomer = New ArrayList
            strItem = New ArrayList
            strCustomer.Clear()
            strItem.Clear()
            Dim blnErrorflag As Boolean = False
            sPath = System.Windows.Forms.Application.StartupPath & "/ImportLog_Invoice.txt"
            If File.Exists(sPath) Then
                File.Delete(sPath)
            End If
            blnErrorflag = False
            strCustomerCodeFiler = " '1'"
            strItemCodeFilter = "'1'"

            strFrom = oApplication.Utilities.getEdittextvalue(aform, "5")
            strTo = oApplication.Utilities.getEdittextvalue(aform, "8")

            If oApplication.Utilities.getEdittextvalue(aform, "5") <> "" Then
                dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "5"))
            End If
            If strTo <> "" Then
                dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
            End If
            Dim strCondition As String
            'If strFrom <> "" Then
            '    strCondition = "convert(varchar(10),[InvDate],105) >='" & dtFrom.ToString("dd-MM-yyyy") & "'"
            'Else
            '    strCondition = " 1 =1"
            'End If
            'If strTo <> "" Then
            '    strCondition = strCondition & " and convert(varchar(10),[InvDate],105) <='" & dtTo.ToString("dd-MM-yyyy") & "'"
            'Else
            '    strCondition = strCondition & " and 2=2"
            'End If

            If strFrom <> "" Then
                '  2013-01-09 00:00:00.000
                strCondition = "[InvDate] >='" & dtFrom.ToString("yyyy-MM-dd") & " 00:00:00.000'"
            Else
                strCondition = " 1 =1"
            End If
            If strTo <> "" Then
                strCondition = strCondition & " and [InvDate] <='" & dtTo.ToString("yyyy-MM-dd") & " 23:59:00.000'"
            Else
                strCondition = strCondition & " and 2=2"
            End If
            oStatic = aform.Items.Item("stProcess").Specific
            WriteErrorlog("Processing Invoice Posting : Date From : " & strFrom & " To : " & strTo, sPath)
            oStatic.Caption = "Processing  "
            str1 = "Select '',[InvSeq],[CustCode],[InvDate],[Quantity],[ItemNo],[Flag],[RefNo],[Remarks],[Price],[Invtype],[AgentCode],PK from  " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False' and " & strCondition
            str1 = "Select [CustCode],[ItemNo],Count(*) from  " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False' and " & strCondition & " group by [CustCode],[ItemNo]"
            Dim otest As SAPbobsCOM.Recordset
            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest.DoQuery(str1)
            strCustomerCodeFiler = " '1'"
            strItemCodeFilter = "'1'"
            For intRow As Integer = 0 To otest.RecordCount - 1
                strCustCode = otest.Fields.Item(0).Value
                strItemCode = otest.Fields.Item(1).Value
                oStatic = aform.Items.Item("stProcess").Specific
                oStatic.Caption = "Processing  Validations.  Customer code  : " & strcustCode & " and ItemCode :  " & strItemCode
                oApplication.Utilities.Message("Processing Validation....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                otemp.DoQuery("Select * from OCRD where U_Z_CardCode='" & strCustCode & "'")
                If otemp.RecordCount <= 0 Then
                    '  oApplication.Utilities.Message("Customer code does not exists  " & strcustCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    blnErrorflag = True
                    WriteErrorlog(" ", sPath)
                    WriteErrorlog("Customer code does not exists  " & strCustCode, sPath)
                    If strCustCode <> "" Then
                        If strCustomerCodeFiler <> "" Then
                            strCustomerCodeFiler = strCustomerCodeFiler & ",'" & strCustCode & "'"
                        Else
                            strCustomerCodeFiler = "'" & strCustCode & "'"
                        End If
                    End If
                    ' Return False
                End If
                otemp.DoQuery("Select * from OITM where ItemCode='" & strItemCode & "'")
                If otemp.RecordCount <= 0 Then
                    WriteErrorlog("Item code does not exists  " & strItemCode, sPath)
                    blnErrorflag = True
                    If strItemCode <> "" Then
                        If strItemCodeFilter <> "" Then
                            strItemCodeFilter = strItemCodeFilter & ",'" & strItemCode & "'"
                        Else
                            strItemCodeFilter = "'" & strItemCode & "'"
                        End If
                    End If
                    ' Return False
                End If
                otest.MoveNext()
            Next

            'dtTemp.ExecuteQuery(str1)
            'OGrid.DataTable = dtTemp
            'Formatgrid(OGrid)

            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "  "
            aform.Freeze(False)
            linkedserverDisconnect()
            Return True
        Catch ex As Exception
            linkedserverDisconnect()
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try

    End Function

    Private Sub Databind_load(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim PcvNofrom, PcvNoto, str1, strVoucherCondition As String
            Dim strVoucherNumber As String = ""
            Dim OGrid As SAPbouiCOM.Grid
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  "
            LoginDetails()
            linkedserverConnect()
            OGrid = aform.Items.Item("7").Specific
            dtTemp = OGrid.DataTable
            Dim dtFrom, dtTo As Date
            Dim strFrom, strTo As String
            strFrom = oApplication.Utilities.getEdittextvalue(aform, "5")
            strTo = oApplication.Utilities.getEdittextvalue(aform, "8")

            If oApplication.Utilities.getEdittextvalue(aform, "5") <> "" Then
                dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "5"))
            End If
            If strTo <> "" Then
                dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
            End If
            Dim strCondition As String
            'If strFrom <> "" Then
            '    strCondition = "convert(varchar(10),[InvDate],105) >='" & dtFrom.ToString("dd-MM-yyyy") & "'"
            'Else
            '    strCondition = " 1 =1"
            'End If
            'If strTo <> "" Then
            '    strCondition = strCondition & " and convert(varchar(10),[InvDate],105) <='" & dtTo.ToString("dd-MM-yyyy") & "'"
            'Else
            '    strCondition = strCondition & " and 2=2"
            'End If

            If strFrom <> "" Then
                '  2013-01-09 00:00:00.000
                strCondition = "[InvDate] >='" & dtFrom.ToString("yyyy-MM-dd") & " 00:00:00.000'"
            Else
                strCondition = " 1 =2"
            End If
            If strTo <> "" Then
                strCondition = strCondition & " and [InvDate] <='" & dtTo.ToString("yyyy-MM-dd") & " 23:59:00.000'"
            Else
                strCondition = strCondition & " and 2=3"
            End If
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  "
            str1 = "Select '',[InvSeq],[CustCode],[InvDate],[Quantity],[ItemNo],[Flag],[RefNo],[Remarks],[Price],[Invtype],[AgentCode],PK from  " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False' and " & strCondition
            dtTemp.ExecuteQuery(str1)
            OGrid.DataTable = dtTemp
            Formatgrid(OGrid)
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "  "
            aform.Freeze(False)
            linkedserverDisconnect()
        Catch ex As Exception
            linkedserverDisconnect()
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try

    End Sub
#End Region

#Region "Write into ErrorLog File"
    Private Sub WriteErrorHeader(ByVal apath As String)
        Dim aSw As System.IO.StreamWriter
        Dim aMessage As String
        aMessage = "FileName : " & apath
        If File.Exists(apath) Then
        End If
        aSw = New StreamWriter(sPath, True)
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        If File.Exists(aPath) Then
        End If
        aSw = New StreamWriter(sPath, True)
        aMessage = Now.ToString("dd-MM-yyyy hh:mm") & "--> " & aMessage
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region
    'Private Function CreateInvoice(ByVal aform As SAPbouiCOM.Form) As Boolean
    '    Dim ORec, oTemp, oTemp1 As SAPbobsCOM.Recordset
    '    Dim Strselect, strsql2 As String
    '    Dim dblprice, dbldiscount As Double
    '    Dim intRow As Integer
    '    Dim RetVal As Long
    '    Dim count As Integer = 0
    '    ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    oGrid = oForm.Items.Item("7").Specific
    '    For intRow = 0 To oGrid.DataTable.Rows.Count - 1
    '        Strselect = oGrid.DataTable.GetValue(0, intRow)
    '        If Strselect = "Y" Then
    '            Dim oOrders As SAPbobsCOM.Documents
    '            oOrders = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
    '            oOrders.DocDate = Now
    '            oOrders.TaxDate = Now
    '            oOrders.DocDueDate = oGrid.DataTable.GetValue("InvDate", intRow)
    '            oOrders.NumAtCard = oGrid.DataTable.GetValue("RefNo", intRow)
    '            oOrders.CardCode = oGrid.DataTable.GetValue("CustCode", intRow)
    '            oOrders.Comments = oGrid.DataTable.GetValue("Remarks", intRow)
    '            oOrders.UserFields.Fields.Item("U_Z_INVType").Value = oGrid.DataTable.GetValue("Invtype", intRow)
    '            oOrders.UserFields.Fields.Item("U_Z_Agent").Value = oGrid.DataTable.GetValue("AgentCode", intRow)
    '            oOrders.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
    '            oOrders.Lines.SetCurrentLine(count)
    '            oOrders.Lines.TaxCode = "PA"
    '            oOrders.Lines.ItemCode = oGrid.DataTable.GetValue("ItemNo", intRow)
    '            oOrders.Lines.Quantity = oGrid.DataTable.GetValue("Quantity", intRow)
    '            oOrders.Lines.UnitPrice = oGrid.DataTable.GetValue("Price", intRow)
    '            RetVal = oOrders.Add()
    '            If (RetVal <> 0) Then
    '                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                Return False
    '            Else
    '                ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                strsql2 = "Update " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] set  Flag ='True' where CustCode='" & oGrid.DataTable.GetValue("CustCode", intRow) & "' and  AgentCode='" & oGrid.DataTable.GetValue("AgentCode", intRow) & "'"
    '                strsql2 = strsql2 & " and  PK='" & oGrid.DataTable.GetValue("PK", intRow) & "' and RefNo='" & oGrid.DataTable.GetValue("RefNo", intRow) & "' "
    '                ORec.DoQuery(strsql2)
    '                oApplication.Utilities.Message("Operation Completed Successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '            End If
    '        End If
    '    Next
    '    Return True
    'End Function

    Private Function CreateInvoice(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim ORec, oTemp, oTemp1, ORec1, oItemRS As SAPbobsCOM.Recordset
        Dim Strselect, strcarcode, strcardname, strquantity, stritemcode, Stritemname, IRefId, strcurrency, strRoute, strTripno As String
        Dim strRefno, strInvtype, strAgent, strdate, strqry2 As String
        Dim dblprice, dbldiscount As Double
        Dim intRow As Integer
        Dim dtEndDate As Date
        Dim RetVal As Long
        Dim count As Integer = 0

        ORec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oItemRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            ' strqry2 = "select CustCode,RefNo,Invtype,AgentCode,InvDate,count(*) from " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False'  group by CustCode,RefNo,Invtype,AgentCode,InvDate"

            Dim dtFrom, dtTo, dtPostingdate As Date
            Dim strFrom, strTo As String
            strFrom = oApplication.Utilities.getEdittextvalue(aform, "12")
            If strFrom = "" Then
                oApplication.Utilities.Message("Posting date missing..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                dtPostingdate = oApplication.Utilities.GetDateTimeValue(strFrom)
            End If

            If oApplication.Utilities.getEdittextvalue(aform, "5") <> "" Then
                dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "5"))
            End If

            strFrom = oApplication.Utilities.getEdittextvalue(aform, "5")
            strTo = oApplication.Utilities.getEdittextvalue(aform, "8")

            If strTo <> "" Then
                dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
            End If
            Dim strCondition As String
            'If strFrom <> "" Then
            '    strCondition = "convert(varchar(10),[InvDate],105) >='" & dtFrom.ToString("dd-MM-yyyy") & "'"
            'Else
            '    strCondition = " 1 =1"
            'End If
            'If strTo <> "" Then
            '    strCondition = strCondition & " and convert(varchar(10),[InvDate],105) <='" & dtTo.ToString("dd-MM-yyyy") & "'"
            'Else
            '    strCondition = strCondition & " and 2=2"
            'End If
            If strFrom <> "" Then
                '  2013-01-09 00:00:00.000
                strCondition = "[InvDate] >='" & dtFrom.ToString("yyyy-MM-dd") & " 00:00:00.000'"
            Else
                strCondition = " 1 =1"
            End If
            If strTo <> "" Then
                strCondition = strCondition & " and [InvDate] <='" & dtTo.ToString("yyyy-MM-dd") & " 23:59:00.000'"
            Else
                strCondition = strCondition & " and 2=2"
            End If

            If strCustomerCodeFiler <> "" Then
                strqry2 = "select CustCode,Invtype,AgentCode,count(*) from " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False'  and  " & strCondition & " and CustCode not in(" & strCustomerCodeFiler & ") group by CustCode,Invtype,AgentCode"
            Else
                strqry2 = "select CustCode,Invtype,AgentCode,count(*) from " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False'  and  " & strCondition & "  group by CustCode,Invtype,AgentCode"
            End If

            ORec.DoQuery(strqry2)
            For intLoop As Integer = 0 To ORec.RecordCount - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim strcustomer As String
                strcustomer = ORec.Fields.Item(0).Value
                ORec1.DoQuery("Select CardCode from OCRD where U_Z_CardCode='" & strcustomer & "'")
                oStatic = aform.Items.Item("stProcess").Specific
                oStatic.Caption = "Processing  Invoice Creation : Customer Code : " & strcustomer
                strcustomer = ORec1.Fields.Item(0).Value
                If strcustomer <> "" Then
                    ' strRefno = ORec.Fields.Item(1).Value
                    strInvtype = ORec.Fields.Item(1).Value
                    strAgent = ORec.Fields.Item(2).Value
                    count = 0
                    Dim oOrders As SAPbobsCOM.Documents
                    oOrders = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                    oOrders.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                    oOrders.DocDate = dtPostingdate ' Now.Date 'strdate
                    oOrders.DocDueDate = dtPostingdate 'Now.Date ' strdate
                    oOrders.TaxDate = dtPostingdate ' Now.Date ' strdate
                    '   oOrders.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                    oOrders.CardCode = strcustomer
                    ' oOrders.NumAtCard = strRefno
                    oOrders.UserFields.Fields.Item("U_Z_Agent").Value = strAgent
                    oOrders.UserFields.Fields.Item("U_Z_INVType").Value = strInvtype
                    Dim strqry As String
                    Dim blnLineExists As Boolean = False
                    If strItemCodeFilter = "" Then
                        strqry = "Select  * from " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False' and CustCode='" & strcustomer & "' "
                        strqry = strqry & " and Invtype='" & strInvtype & "' and AgentCode='" & strAgent & "'   and  " & strCondition & ""
                    Else
                        strqry = "Select  * from " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice] where [Flag]='False' and CustCode='" & strcustomer & "' "
                        strqry = strqry & " and Invtype='" & strInvtype & "' and AgentCode='" & strAgent & "'   and  " & strCondition & " and ItemNo not in (" & strItemCodeFilter & ")"
                    End If
                    oTemp1.DoQuery(strqry)
                    'oGrid = oForm.Items.Item("9").Specific
                    Dim strItemCode1 As String
                    For intRow = 0 To oTemp1.RecordCount - 1
                        ' strcarcode = oTemp1.Fields.Item("CustCode").Value
                        Try

                        
                            oStatic = aform.Items.Item("stProcess").Specific
                            oStatic.Caption = "Processing  Invoice Creation : Customer Code : " & strcustomer
                        Catch ex As Exception

                        End Try
                        strItemCode1 = oTemp1.Fields.Item("ItemNo").Value
                        oItemRS.DoQuery("Select * from OITM where ItemCode='" & strItemCode1 & "'")
                        If oItemRS.RecordCount > 0 Then
                            oApplication.Utilities.Message("Processing Invoice Creation....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            blnLineExists = True
                            oOrders.Comments = oTemp1.Fields.Item("Remarks").Value
                            oOrders.UserFields.Fields.Item("U_Z_InvSeqNo").Value = oTemp1.Fields.Item("InvSeq").Value.ToString()
                            If count > 0 Then
                                oOrders.Lines.Add()
                            End If
                            oOrders.Lines.SetCurrentLine(count)
                            oOrders.Lines.UserFields.Fields.Item("U_Z_Ref").Value = oTemp1.Fields.Item("RefNo").Value
                            oOrders.Lines.UserFields.Fields.Item("U_Z_TrnsDate").Value = oTemp1.Fields.Item("InvDate").Value
                            oOrders.Lines.ItemCode = oTemp1.Fields.Item("ItemNo").Value
                            oOrders.Lines.Quantity = oTemp1.Fields.Item("Quantity").Value
                            oOrders.Lines.UnitPrice = oTemp1.Fields.Item("Price").Value
                            ' oOrders.Lines.VatGroup = "X0"
                            count = count + 1
                            'End If
                        End If
                        oTemp1.MoveNext()
                    Next
                    If blnLineExists = True Then
                        RetVal = oOrders.Add()
                        If (RetVal <> 0) Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oStatic = aform.Items.Item("stProcess").Specific
                            oStatic.Caption = oApplication.Company.GetLastErrorDescription
                            'linkedserverDisconnect()
                            Return False
                        Else
                            Try
                                strSQL = "update " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Invoice]  set [Flag]='True' where CustCode='" & strcustomer & "' "
                                strSQL = strSQL & " and  Invtype='" & strInvtype & "' and AgentCode='" & strAgent & "' and ( " & strCondition & " ) and CustCode not in (" & strCustomerCodeFiler & ") and ItemNo not in (" & strItemCodeFilter & ")"
                                ORec1.DoQuery(strSQL)
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                End If
                ORec.MoveNext()
            Next
            oApplication.Utilities.Message("Operation Completed Successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Try
                oStatic = aform.Items.Item("stProcess").Specific
                oStatic.Caption = "Import completed"
            Catch ex As Exception
            End Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return True
    End Function

    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strGridGLCode, strAcctcode, Strselect, strcustCode, strItemCode As String
        Dim oTemp, oTemp1 As SAPbobsCOM.Recordset
        Dim blnRecordselected As Boolean = False
        Dim strCustomer, strItem As ArrayList
        strCustomer = New ArrayList
        strItem = New ArrayList
        strCustomer.Clear()
        strItem.Clear()
        Dim blnErrorflag As Boolean = False
        sPath = System.Windows.Forms.Application.StartupPath & "/ImportLog_Invoice.txt"
        If File.Exists(sPath) Then
            File.Delete(sPath)
        End If
        blnErrorflag = False
        strCustomerCodeFiler = " '1'"
        strItemCodeFilter = "'1'"


        Dim dtFrom, dtTo As Date
        Dim strFrom, strTo As String
        strFrom = oApplication.Utilities.getEdittextvalue(aform, "5")
        strTo = oApplication.Utilities.getEdittextvalue(aform, "8")

        If oApplication.Utilities.getEdittextvalue(aform, "5") <> "" Then
            dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "5"))
        End If
        If strTo <> "" Then
            dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
        End If
        WriteErrorlog("Processing Invoice Posting : Date From : " & strFrom & " To : " & strTo, sPath)
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            oApplication.Utilities.Message("Processing Validations....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            blnRecordselected = True
            strcustCode = oGrid.DataTable.GetValue("CustCode", intRow)
            strItemCode = oGrid.DataTable.GetValue("ItemNo", intRow)
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  Validations.  Customer code  : " & strcustCode & " and ItemCode :  " & strItemCode

            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strCustomer.Contains(strcustCode) = False Then
                oTemp.DoQuery("Select * from OCRD where U_Z_CardCode='" & strcustCode & "'")
                If oTemp.RecordCount <= 0 Then
                    '  oApplication.Utilities.Message("Customer code does not exists  " & strcustCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    blnErrorflag = True
                    WriteErrorlog(" ", sPath)
                    WriteErrorlog("Customer code does not exists  " & strcustCode, sPath)
                    If strcustCode <> "" Then
                        If strCustomerCodeFiler <> "" Then
                            strCustomerCodeFiler = strCustomerCodeFiler & ",'" & strcustCode & "'"
                        Else
                            strCustomerCodeFiler = "'" & strcustCode & "'"
                        End If
                    End If

                    ' Return False
                End If
                strCustomer.Add(strcustCode)
            End If
            oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strItem.Contains(strItemCode) = False Then
                oTemp1.DoQuery("Select * from OITM where ItemCode='" & strItemCode & "'")
                If oTemp1.RecordCount <= 0 Then
                    '  oApplication.Utilities.Message("Item code does not exists  " & strItemCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    blnErrorflag = True
                    WriteErrorlog("Item code does not exists  " & strItemCode, sPath)
                    If strItemCode <> "" Then
                        If strItemCodeFilter <> "" Then
                            strItemCodeFilter = strItemCodeFilter & ",'" & strItemCode & "'"
                        Else
                            strItemCodeFilter = "'" & strItemCode & "'"
                        End If
                    End If
                    ' Return False
                End If
                strItem.Add(strItemCode)
            End If
        Next
        If blnRecordselected = False Then
            oApplication.Utilities.Message("No row selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        oStatic = aform.Items.Item("stProcess").Specific
        oStatic.Caption = "Validation Completed"

        Return True
    End Function
#Region "GetServerDetails"
    Private Sub LoginDetails()
        Dim ORec As SAPbobsCOM.Recordset
        ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        str = "Select U_Z_SERVER,U_Z_SQLDB,U_Z_SQLUID,U_Z_SQLPWD, U_Z_LOCALSQLDB,U_Z_LOCUID from [@Z_INT_LOGIN]"
        ORec.DoQuery(str)
        If ORec.RecordCount > 0 Then
            Server = ORec.Fields.Item(0).Value
            serverdb = ORec.Fields.Item(1).Value
            ServerUid = ORec.Fields.Item(2).Value
            ServerPwd = ORec.Fields.Item(3).Value
            LocalDb = ORec.Fields.Item(4).Value
            UID = ORec.Fields.Item(5).Value
        Else
            oApplication.SBO_Application.SetStatusBarMessage("Please Enter the DB Credentails", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End If
    End Sub
#End Region
#Region "Linked Server Connection"
    Private Sub linkedserverConnect()
        Try
            LoginDetails()
            Dim ORec, ORec2 As SAPbobsCOM.Recordset
            ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ORec.DoQuery("sp_addlinkedserver  '" & Server & "'")
            ORec2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ORec2.DoQuery("sp_addlinkedsrvlogin '" & Server & "', 'false', '" & UID & "', '" & ServerUid & "', '" & ServerPwd & "'")
        Catch ex As Exception
            linkedserverDisconnect()
        End Try
    End Sub
#End Region
#Region "Linked Server  DisConnection"
    Private Sub linkedserverDisconnect()
        Try
            LoginDetails()
            Dim ORec, ORec2 As SAPbobsCOM.Recordset
            ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ORec.DoQuery("sp_DROPlinkedsrvlogin '" & Server & "',  '" & UID & "'")
            ORec2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ORec2.DoQuery("SP_DROPSERVER  '" & Server & "'")
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_CustInvoice Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1000003" Then
                                    oGrid = oForm.Items.Item("7").Specific
                                    If oApplication.SBO_Application.MessageBox("Do you want to import the Details?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    If oApplication.Utilities.linkedserverConnect() = True Then
                                        If oApplication.Utilities.getEdittextvalue(oForm, "12") = "" Then
                                            oApplication.Utilities.Message("Posting Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            Exit Sub
                                        End If
                                        ' If validation(oGrid, oForm) = True Then
                                        If Valiation_Table(oForm) = True Then
                                            linkedserverConnect()
                                            If CreateInvoice(oForm) = True Then
                                                linkedserverDisconnect()
                                                Databind(oForm)
                                            Else
                                                linkedserverDisconnect()
                                            End If
                                        End If

                                        'oStaticText = oForm.Items.Item("12").Specific
                                        'oStaticText.Caption = ""
                                        'objUtility.ShowMessage("Import process encounterd with some errors, please correct and reimport..")
                                        Dim x As System.Diagnostics.ProcessStartInfo
                                        x = New System.Diagnostics.ProcessStartInfo
                                        x.UseShellExecute = True
                                        sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
                                        x.FileName = sPath
                                        System.Diagnostics.Process.Start(x)
                                        x = Nothing
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "10" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to get the details ?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    Databind(oForm)
                                End If
                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Invoice
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
