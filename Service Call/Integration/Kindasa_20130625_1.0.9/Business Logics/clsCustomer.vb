Public Class clsCustomer
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
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private oStatic As SAPbouiCOM.StaticText
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Public LocalDb, strDocEntry, Server, UID, ServerUid, ServerPwd, servertype, serverdb, str, str1 As String

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.linkedserverConnect() = True Then
            oForm = oApplication.Utilities.LoadForm(xml_Customer, frm_Customer)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            Databind(oForm)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            oForm.Freeze(False)
        End If
        
    End Sub
#Region "FormatGrid"
    Private Sub Formatgrid(ByVal agrid As SAPbouiCOM.Grid)
        agrid.Columns.Item(0).TitleObject.Caption = "Select"
        agrid.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        agrid.Columns.Item(0).Visible = False
        agrid.Columns.Item(1).TitleObject.Caption = "Customer Code"
        agrid.Columns.Item(2).TitleObject.Caption = "Customer Name"
        agrid.Columns.Item(3).TitleObject.Caption = "Foreign Name"
        agrid.Columns.Item(4).TitleObject.Caption = "Flag"
        agrid.Columns.Item(4).Visible = False
        agrid.Columns.Item(5).TitleObject.Caption = "Type"
        agrid.Columns.Item(6).TitleObject.Caption = "Balance"
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
            LoginDetails()
            linkedserverConnect()
            OGrid = aform.Items.Item("9").Specific
            dtTemp = OGrid.DataTable
            str1 = "Select '',[CustCode],[CustName],[CustNameF],[Flag],[Type],[Balance] from  " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Customer] where [Flag]='False'"
            dtTemp.ExecuteQuery(str1)
            oGrid.DataTable = dtTemp
            Formatgrid(OGrid)
            aform.Freeze(False)
            linkedserverDisconnect()
        Catch ex As Exception
            linkedserverDisconnect()
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try

    End Sub
#End Region

    Public Function CreateCustomer(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim intRow As Integer
        Dim strsql2, Strselect, strCardCode As String
        Dim ORec, ORec1, oRec2, oTemp1 As SAPbobsCOM.Recordset
        Dim BP, BP1 As SAPbobsCOM.BusinessPartners
        Dim intsize As Integer
        BP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        Try
            ORec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            str1 = "Select [CustCode],[CustName],[CustNameF],[Flag],[Type],[Balance] from  " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Customer] where [Flag]='False'"
            ORec.DoQuery(str1)
            For intLoop As Integer = 0 To ORec.RecordCount - 1
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oRec2.DoQuery("Select isnull(CardCode,'') from OCRD where U_Z_CardCode='" & ORec.Fields.Item("CustCode").Value & "'")
                strCardCode = oRec2.Fields.Item(0).Value
                oStatic = aform.Items.Item("stProcess").Specific
                oStatic.Caption = "Processing  Customer code : " & strCardCode

                If strCardCode <> "" Then
                    If BP.GetByKey(strCardCode) Then
                        BP.CardCode = strCardCode
                        BP.CardName = ORec.Fields.Item("CustName").Value
                        BP.CardForeignName = ORec.Fields.Item("CustNameF").Value
                        BP.UserFields.Fields.Item("U_Z_CardCode").Value = ORec.Fields.Item("CustCode").Value
                        BP.UserFields.Fields.Item("U_Z_Type").Value = ORec.Fields.Item("Type").Value
                        If BP.Update() <> 0 Then
                            oApplication.Utilities.Message("failed to add Customer Updation :" & oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            oApplication.Utilities.Message("Customer Information Successfully Updated...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    Else
                        oTemp1.DoQuery("select series,*  from NNM1 where objectcode=2 and IsManual='N' and locked='N' and seriestype='B' and Series=(select DfltSeries  from ONNm where ObjectCode=2 and DocSubType='C')")
                        If oTemp1.RecordCount > 0 Then
                            BP.Series = oTemp1.Fields.Item(0).Value
                            intsize = oTemp1.Fields.Item("NumSize").Value
                        Else
                            BP.CardCode = ORec.Fields.Item("CustCode").Value
                        End If
                        ' BP.CardCode = strCardCode ' ORec.Fields.Item("CustCode").Value
                        BP.CardName = ORec.Fields.Item("CustName").Value
                        BP.CardForeignName = ORec.Fields.Item("CustNameF").Value
                        BP.UserFields.Fields.Item("U_Z_Type").Value = ORec.Fields.Item("Type").Value
                        BP.UserFields.Fields.Item("U_Z_CardCode").Value = ORec.Fields.Item("CustCode").Value
                        If BP.Add() <> 0 Then
                            oApplication.Utilities.Message("failed to add Customer Creation :" & oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            ' oApplication.Utilities.Message("Customer Information Successfully Created...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                Else

                    oTemp1.DoQuery("select series,*  from NNM1 where objectcode=2 and IsManual='N' and locked='N' and seriestype='B' and Series=(select DfltSeries  from ONNm where ObjectCode=2 and DocSubType='C')")
                    If oTemp1.RecordCount > 0 Then
                        BP.Series = oTemp1.Fields.Item(0).Value
                        intsize = oTemp1.Fields.Item("NumSize").Value
                    Else
                        'oTemp1.DoQuery("select  'C'+convert(varchar,max(replace(CardCode,'C','')+1)) from OCRD where CardCode like  'C%' and CardType='C'")
                        'If oTemp1.RecordCount > 0 Then
                        '    BP.CardCode = oTemp1.Fields.Item(0).Value
                        'End If
                        BP.CardCode = ORec.Fields.Item("CustCode").Value
                    End If
                    ' BP.CardCode = ORec.Fields.Item("CustCode").Value
                    BP.CardName = ORec.Fields.Item("CustName").Value
                    BP.CardForeignName = ORec.Fields.Item("CustNameF").Value
                    BP.UserFields.Fields.Item("U_Z_Type").Value = ORec.Fields.Item("Type").Value
                    BP.UserFields.Fields.Item("U_Z_CardCode").Value = ORec.Fields.Item("CustCode").Value
                    If BP.Add() <> 0 Then
                        oApplication.Utilities.Message("failed to add Customer Creation :" & oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    Else
                        '  oApplication.Utilities.Message("Customer Information Successfully Created...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If

                End If


                ORec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strsql2 = "Update " & "[" & Server & "]" & "." & serverdb & "." & "dbo" & "." & "[SAP_Customer] set  Flag ='True' where CustCode='" & ORec.Fields.Item("CustCode").Value & "'"
                ORec1.DoQuery(strsql2)
                ORec.MoveNext()
            Next
            oStatic = aform.Items.Item("stProcess").Specific
            oStatic.Caption = "Import Completed"


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return True
    End Function
    Private Function validation(ByVal aGrid As SAPbouiCOM.Grid) As Boolean
        Dim strGridGLCode, strAcctcode, Strselect As String
        Dim blnRecordselected As Boolean = False
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            Strselect = oGrid.DataTable.GetValue(0, intRow)
            If Strselect = "Y" Then
                Return True
            End If
        Next
        If blnRecordselected = False Then
            oApplication.Utilities.Message("No row selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
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
            If pVal.FormTypeEx = frm_Customer Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "3" Then
                                    If oApplication.Utilities.linkedserverConnect() = True Then
                                        linkedserverConnect()
                                        If CreateCustomer(oForm) = True Then
                                            linkedserverDisconnect()
                                            Databind(oForm)
                                        Else
                                            linkedserverDisconnect()
                                        End If
                                    End If
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
                Case mnu_Customer
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
