Public Class clsPrepaid
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oStatic As SAPbouiCOM.StaticText
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public LocalDb, strDocEntry, Server, UID, ServerUid, ServerPwd, servertype, serverdb, str As String

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.linkedserverConnect() = True Then
            oForm = oApplication.Utilities.LoadForm(xml_Prepaid, frm_Prepaid)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.DataSources.UserDataSources.Add("AcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oEditText = oForm.Items.Item("5").Specific
            oEditText.DataBind.SetBound(True, "", "AcCode")
            oEditText.ChooseFromListUID = "CFL1"
            oEditText.ChooseFromListAlias = "Formatcode"
            Databind_Load(oForm)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            oForm.Freeze(False)
        End If
    End Sub
#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        agrid.Columns.Item(0).TitleObject.Caption = "Select"
        agrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item(0).Visible = False
        agrid.Columns.Item(1).TitleObject.Caption = "Inc.Payment Sequence"
        agrid.Columns.Item(2).TitleObject.Caption = "customer code"
        agrid.Columns.Item(3).TitleObject.Caption = "Payment Date"
        agrid.Columns.Item(4).TitleObject.Caption = "Flag"
        agrid.Columns.Item(4).Visible = False
        agrid.Columns.Item(5).TitleObject.Caption = "Reference No"
        agrid.Columns.Item(6).TitleObject.Caption = "Amount"
        agrid.Columns.Item(7).TitleObject.Caption = "Remarks"
        agrid.Columns.Item(8).TitleObject.Caption = "Payment Type"
        agrid.Columns.Item(9).TitleObject.Caption = "Check No"
        agrid.Columns.Item(10).TitleObject.Caption = "Bank"
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
            strFrom = oApplication.Utilities.getEdittextvalue(aform, "12")
            strTo = oApplication.Utilities.getEdittextvalue(aform, "9")

            If oApplication.Utilities.getEdittextvalue(aform, "12") <> "" Then
                dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "12"))
            End If
            If strTo <> "" Then
                dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
            End If
            Dim strCondition As String
            If strFrom <> "" Then
                '  2013-01-09 00:00:00.000
                strCondition = "[PreDate] >='" & dtFrom.ToString("yyyy-MM-dd") & " 00:00:00.000'"
            Else
                strCondition = " 1 =1"
            End If
            If strTo <> "" Then
                strCondition = strCondition & " and [PreDate] <='" & dtTo.ToString("yyyy-MM-dd") & " 23:59:00.000'"
            Else
                strCondition = strCondition & " and 2=2"
            End If
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  "
            str1 = "Select '',[PreSeq],[CustCode],[PreDate],[Flag],[RefNo],[Amount],[Remarks],[Pretype],[CheckNo],[Bank] from  " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Prepaid] where [Flag]='False' and " & strCondition

            dtTemp.ExecuteQuery(str1)
            oGrid.DataTable = dtTemp
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

    Private Sub Databind_Load(ByVal aform As SAPbouiCOM.Form)
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
            strFrom = oApplication.Utilities.getEdittextvalue(aform, "12")
            strTo = oApplication.Utilities.getEdittextvalue(aform, "9")

            If oApplication.Utilities.getEdittextvalue(aform, "12") <> "" Then
                dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "12"))
            End If
            If strTo <> "" Then
                dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
            End If
            Dim strCondition As String
            If strFrom <> "" Then
                '  2013-01-09 00:00:00.000
                strCondition = "[PreDate] >='" & dtFrom.ToString("yyyy-MM-dd") & " 00:00:00.000'"
            Else
                strCondition = " 1 =2"
            End If
            If strTo <> "" Then
                strCondition = strCondition & " and [PreDate] <='" & dtTo.ToString("yyyy-MM-dd") & " 23:59:00.000'"
            Else
                strCondition = strCondition & " and 2=3"
            End If

            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  "


            str1 = "Select '',[PreSeq],[CustCode],[PreDate],[Flag],[RefNo],[Amount],[Remarks],[Pretype],[CheckNo],[Bank] from  " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Prepaid] where [Flag]='False' and " & strCondition
            dtTemp.ExecuteQuery(str1)
            OGrid.DataTable = dtTemp
            Formatgrid(OGrid)
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = ""
            aform.Freeze(False)
            linkedserverDisconnect()
        Catch ex As Exception
            linkedserverDisconnect()
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try

    End Sub
#End Region
#Region "GetServerDetails"
    Private Sub LoginDetails()
        Dim ORec As SAPbobsCOM.Recordset
        ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        str = "Select U_Z_SERVER,U_Z_SQLDB,U_Z_SQLUID,U_Z_SQLPWD, U_Z_LOCALSQLDB,U_Z_LOCUID from [@Z_INT_LOGIN]"
        ORec.DoQuery(Str)
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
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strGridGLCode, strAcctcode, Strselect, strcustCode, strItemCode As String
        Dim oTemp, oTemp1 As SAPbobsCOM.Recordset
        Dim blnRecordselected As Boolean = False
        Dim strCustomer As ArrayList
        strCustomer = New ArrayList
        strAcctcode = oApplication.Utilities.getEdittextvalue(oForm, "5")
        If strAcctcode = "" Then
            oApplication.Utilities.Message("Cash Account Number can not be empty ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            oApplication.Utilities.Message("Processing validation...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            blnRecordselected = True
            strcustCode = oGrid.DataTable.GetValue("CustCode", intRow)

            oStatic = Form.Items.Item("stProcess").Specific
            oStatic.Caption = "Processing  Customer code : " & strcustCode


            If strCustomer.Contains(strcustCode) = False Then
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery("Select * from OCRD where U_Z_CardCode='" & strcustCode & "'")
                If oTemp.RecordCount <= 0 Then
                    oApplication.Utilities.Message("Customer code does not exists  " & strcustCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oStatic.Caption = "Customer code does not exits : " & strcustCode
                    Return False
                End If
                strCustomer.Add(strcustCode)
            End If
        Next
        oStatic = aform.Items.Item("stProcess").Specific
        oStatic.Caption = "Validation Completed"


        Return True
    End Function

    Private Function IncomingPayment(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim ORec, oTemp, oTemp1, ORec1 As SAPbobsCOM.Recordset
        Dim strRefno, strInvtype, strAgent, strAcctCode, strqry2 As String
        Dim intRow As Integer
        Dim dtEndDate As Date
        Dim RetVal As Long
        Dim count As Integer = 0
        ORec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try

            Dim dtFrom, dtTo, dtPostingdate As Date
            Dim strFrom, strTo, strCondition As String
            strFrom = oApplication.Utilities.getEdittextvalue(aform, "13")
            If strFrom = "" Then
                oApplication.Utilities.Message("Posting date missing..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                dtPostingdate = oApplication.Utilities.GetDateTimeValue(strFrom)
            End If

            If oApplication.Utilities.getEdittextvalue(aform, "12") <> "" Then
                dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aform, "12"))
            End If



            strFrom = oApplication.Utilities.getEdittextvalue(aform, "12")
            strTo = oApplication.Utilities.getEdittextvalue(aform, "9")
            If strTo <> "" Then
                dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
            End If
            If strFrom <> "" Then
                '  2013-01-09 00:00:00.000
                strCondition = "[PreDate] >='" & dtFrom.ToString("yyyy-MM-dd") & " 00:00:00.000'"
            Else
                strCondition = " 1 =1"
            End If
            If strTo <> "" Then
                strCondition = strCondition & " and [PreDate] <='" & dtTo.ToString("yyyy-MM-dd") & " 23:59:00.000'"
            Else
                strCondition = strCondition & " and 2=2"
            End If

            strqry2 = "select PreSeq,sum(Amount) as Amount,count(*) from " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Prepaid] where [Flag]='False' and " & strCondition & "  group by PreSeq"
            ORec.DoQuery(strqry2)
            For intLoop As Integer = 0 To ORec.RecordCount - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                strAcctCode = oApplication.Utilities.getEdittextvalue(aform, "5")
                oStatic = aform.Items.Item("stProcess").Specific
                oStatic.Caption = "Processing Import  ...."


                Dim vPay As SAPbobsCOM.Payments
                vPay = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
                vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
                strRefno = ORec.Fields.Item(0).Value
                Dim strqry As String
                strqry = "Select  * from " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Prepaid] where [Flag]='False'  and " & strCondition & "  and PreSeq='" & strRefno & "' "
                oTemp.DoQuery(strqry)
                For intRow = 0 To oTemp.RecordCount - 1
                    Dim strcustomer As String
                    Try
                        oStatic = aform.Items.Item("stProcess").Specific
                        oStatic.Caption = "Processing Import  ...."
                    Catch ex As Exception
                    End Try
                    oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strcustomer = oTemp.Fields.Item("CustCode").Value
                    ORec1.DoQuery("Select CardCode from OCRD where U_Z_CardCode='" & strcustomer & "'")
                    strcustomer = ORec1.Fields.Item(0).Value
                    vPay.DocDate = dtPostingdate
                    vPay.CardCode = strcustomer '  oTemp.Fields.Item("CustCode").Value
                    Dim strref As String
                    strref = oTemp.Fields.Item("RefNo").Value
                    'vPay.CounterReference = strref
                    ' vPay.UserFields.Fields.Item("U_Z_PreSeqNo").Value = strref
                    If strref.Length > 7 Then
                        strref = strref.Substring(0, 7)
                    End If
                    vPay.CounterReference = strref
                    If oTemp.Fields.Item("PreType").Value = "C" Then
                        vPay.CashAccount = oApplication.Utilities.getSAPAccountcode(strAcctCode) 'oTemp.Fields.Item("CheckNo").Value
                        vPay.CashSum = ORec.Fields.Item("Amount").Value
                    Else
                        vPay.CheckAccount = oApplication.Utilities.getSAPAccountcode(strAcctCode) 'oTemp.Fields.Item("CheckNo").Value
                        vPay.Checks.CheckSum = ORec.Fields.Item("Amount").Value
                        vPay.Checks.CheckNumber = oTemp.Fields.Item("CheckNo").Value
                    End If
                    vPay.DocDate = dtPostingdate ' oTemp.Fields.Item("PreDate").Value
                    vPay.Remarks = oTemp.Fields.Item("Remarks").Value
                    vPay.UserFields.Fields.Item("U_Z_PreSeqNo").Value = oTemp.Fields.Item("PreSeq").Value
                    vPay.UserFields.Fields.Item("U_Z_PreType").Value = oTemp.Fields.Item("Pretype").Value
                    oTemp.MoveNext()
                Next
                RetVal = vPay.Add()
                If (RetVal <> 0) Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    linkedserverDisconnect()
                    Return False
                Else
                    oApplication.Utilities.Message("Operation Completed Successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    strSQL = "update " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Prepaid]  set [Flag]='True' where PreSeq='" & strRefno & "' and " & strCondition
                    ORec1.DoQuery(strSQL)
                End If
                ORec.MoveNext()
            Next
            Try
                oStatic = aform.Items.Item("stProcess").Specific
                oStatic.Caption = "Import process Completed"
            Catch ex As Exception

            End Try
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return True
    End Function

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Prepaid Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                '    oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1000003" Then

                                    If oApplication.SBO_Application.MessageBox("Do you want to import the Details?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    If oApplication.Utilities.getEdittextvalue(oForm, "13") = "" Then
                                        oApplication.Utilities.Message("Posting Date missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                    oGrid = oForm.Items.Item("7").Specific
                                    If oApplication.Utilities.getEdittextvalue(oForm, "5") = "" Then
                                        oApplication.Utilities.Message("Account Code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                    If oApplication.Utilities.linkedserverConnect() = True Then
                                        linkedserverConnect()
                                        If IncomingPayment(oForm) = True Then
                                            linkedserverDisconnect()
                                            Databind(oForm)
                                        Else
                                            linkedserverDisconnect()
                                        End If
                                    End If
                                End If
                                If pVal.ItemUID = "10" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to get the details ?", , "Yes", "No") = 2 Then
                                        Exit Sub
                                    End If
                                    Databind(oForm)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim sCHFL_ID As String

                                Dim intChoice As Integer
                                Dim codebar, val1, val As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "5" Then
                                            val1 = oDataTable.GetValue("FormatCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val1)
                                        End If
                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try
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
                Case mnu_Prepaid
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
