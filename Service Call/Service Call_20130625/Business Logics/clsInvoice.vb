Public Class clsInvoice
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Create AP Invoice"
    Private Function CreateAPinvoice(ByVal aDocEntry As Integer) As Boolean

        Dim oSalesOrder As SAPbobsCOM.ServiceCalls
        Dim oAPInvoice As SAPbobsCOM.ServiceCalls
        Dim oTempRec As SAPbobsCOM.Recordset
        Dim strNumAtCard, strCardCode As String
        Dim dtstartDate, dtEndDate As Date
        Dim strRecType As String
        Dim dtPurchaseDate As Date
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If oTempRec.RecordCount <= 0 Then
            oSalesOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
            oAPInvoice = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
            If oSalesOrder.GetByKey(aDocEntry) Then
                oTempRec.DoQuery("Select isnull(U_Z_Status,'N') from OSCS where statusid=" & oSalesOrder.Status)
                'If oSalesOrder.Status = -1 And oSalesOrder.UserFields.Fields.Item("U_Z_RecType").Value <> "N" Then
                If oTempRec.Fields.Item(0).Value = "Y" And oSalesOrder.UserFields.Fields.Item("U_Z_RecType").Value <> "N" Then
                    If oSalesOrder.UserFields.Fields.Item("U_Z_IsDuplicate").Value = "Y" Then
                        Return True
                    End If
                    oAPInvoice.CustomerCode = oSalesOrder.CustomerCode
                    oAPInvoice.InternalSerialNum = oSalesOrder.InternalSerialNum
                    oAPInvoice.ItemCode = oSalesOrder.ItemCode
                    oAPInvoice.ManufacturerSerialNum = oSalesOrder.ManufacturerSerialNum
                    oAPInvoice.Origin = oSalesOrder.Origin
                    oAPInvoice.Priority = oSalesOrder.Priority
                    oAPInvoice.ProblemType = oSalesOrder.ProblemType
                    oAPInvoice.Queue = oSalesOrder.Queue
                    Dim otec As SAPbobsCOM.Recordset
                    otec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    otec.DoQuery("select StatusID from OSCS where Name='Approved'")
                    If otec.RecordCount > 0 Then
                        oSalesOrder.Status = otec.Fields.Item(0).Value
                    Else
                        ' oSalesOrder.Status = -3
                    End If

                    If oSalesOrder.TechnicianCode > 0 Then
                        oAPInvoice.TechnicianCode = oSalesOrder.TechnicianCode
                    End If
                    oAPInvoice.Resolution = oSalesOrder.Resolution
                    If oSalesOrder.UserFields.Fields.Item("U_Z_RecType").Value = "D" Then
                        'oAPInvoice.ResolutionDate = DateAdd(DateInterval.Day, 1, oSalesOrder.ResolutionDate)
                    ElseIf oSalesOrder.UserFields.Fields.Item("U_Z_RecType").Value = "W" Then
                        'oAPInvoice.ResolutionDate = DateAdd(DateInterval.Day, 7, oSalesOrder.ResolutionDate)
                    Else
                        ' oAPInvoice.ResolutionDate = DateAdd(DateInterval.Month, 1, oSalesOrder.ResolutionDate)
                    End If
                    ' MsgBox(oSalesOrder.UserFields.Fields.Item("U_Z_RecType").Value)
                    oAPInvoice.AssigneeCode = oSalesOrder.AssigneeCode
                    oAPInvoice.CallType = oSalesOrder.CallType
                    oAPInvoice.ContactCode = oSalesOrder.ContactCode
                    oAPInvoice.ContractID = oSalesOrder.ContractID
                    oAPInvoice.CustomerName = oSalesOrder.CustomerName
                    'oAPInvoice.EntitledforService = oSalesOrder.EntitledforService
                    oAPInvoice.Description = oSalesOrder.Description
                    oAPInvoice.UserFields.Fields.Item("U_Z_RecType").Value = oSalesOrder.UserFields.Fields.Item("U_Z_RecType").Value
                    oAPInvoice.UserFields.Fields.Item("U_Z_IsDuplicate").Value = "N"
                    oAPInvoice.Subject = oSalesOrder.Subject
                    If oAPInvoice.Add <> 0 Then
                        'MsgBox(oApplication.Company.GetLastErrorDescription)
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Else
                        Dim strdocnum As String
                        oApplication.Company.GetNewObjectCode(strdocnum)
                        oApplication.SBO_Application.MessageBox("Recurring Service call Created successfully: " & strdocnum)
                        oApplication.Utilities.Message("Recurring Service call created successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Dim otest As SAPbobsCOM.Recordset
                        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        otest.DoQuery("Select isnull(U_Z_RecType,'D') 'RecType', * from OSCL where CallID=" & aDocEntry)
                        If otest.RecordCount > 0 Then
                            Dim strRecType1 As String
                            strRecType1 = otest.Fields.Item("RecType").Value
                            dtstartDate = otest.Fields.Item("StartDate").Value
                            If strRecType1 = "D" Then
                                dtstartDate = DateAdd(DateInterval.Day, 1, dtstartDate)
                            ElseIf strRecType1 = "W" Then
                                dtstartDate = DateAdd(DateInterval.Day, 7, dtstartDate)
                            ElseIf strRecType1 = "M" Then
                                dtstartDate = DateAdd(DateInterval.Month, 1, dtstartDate)
                            ElseIf strRecType1 = "Q" Then
                                dtstartDate = DateAdd(DateInterval.Day, 90, dtstartDate)
                            ElseIf strRecType1 = "4M" Then
                                dtstartDate = DateAdd(DateInterval.Day, 120, dtstartDate)
                            ElseIf strRecType1 = "S" Then
                                dtstartDate = DateAdd(DateInterval.Day, 180, dtstartDate)
                            ElseIf strRecType1 = "Y" Then
                                dtstartDate = DateAdd(DateInterval.Day, 360, dtstartDate)
                            ElseIf strRecType1 = "2Y" Then
                                dtstartDate = DateAdd(DateInterval.Day, 720, dtstartDate)
                            Else
                                dtstartDate = dtstartDate
                            End If
                            'dtEndDate = otest.Fields.Item("Enddate").Value
                            Try
                                Dim strstring As String
                                strstring = "Update OSCL set Startdate='" & dtstartDate.ToString("yyyy-MM-dd") & "',EndDate='" & dtstartDate.ToString("yyyy-MM-dd") & "' where CallID=" & strdocnum
                                otest.DoQuery(strstring)
                                strstring = "Update OSCL set starttime=0800,EndTime=0900,duration=1  , DurType='H' where CallID=" & strdocnum
                                otest.DoQuery(strstring)
                                otec.DoQuery("select StatusID from OSCS where Name='Open'")
                                If otec.RecordCount > 0 Then
                                    strstring = "Update OSCL set status=" & otec.Fields.Item(0).Value & " where CallID=" & strdocnum
                                    otest.DoQuery(strstring)
                                Else
                                    ' oSalesOrder.Status = -3
                                End If
                                strstring = "Update OSCL set U_Z_IsDuplicate='Y' where CallID=" & aDocEntry
                                otest.DoQuery(strstring)
                            Catch ex As Exception
                                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        End If
                    End If
                End If
            End If
        End If
        Return True
    End Function
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.MenuUID
            Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            Case "Dup"
                If pVal.BeforeAction = False Then
                    Dim strDocNum, strDocType As String
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    Dim bisObj As SAPbouiCOM.BusinessObject = Form.BusinessObject
                    Dim otest As SAPbobsCOM.Recordset
                    Dim str As Integer
                    Dim strString As String
                    Dim BP1 As SAPbobsCOM.ServiceCalls  '= oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    Select Case oForm.TypeEx
                        Case frm_ServiceCall
                            BP1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
                            If oApplication.Utilities.getEdittextvalue(oForm, "540000180") <> "" Then
                                'otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'strString = "SELECT *  FROM OSCL T0 WHERE  isnull(U_Z_IsDuplicate,'N')='N' and Status=-1 and  T0.[DocNum] =" & oApplication.Utilities.getEdittextvalue(oForm, "540000180")
                                'otest.DoQuery(strString)
                                'If otest.RecordCount > 0 Then
                                '    str = otest.Fields.Item("CallID").Value
                                '    ' str = CInt(oApplication.Utilities.getEdittextvalue(oForm, "540000180"))
                                '    CreateAPinvoice(str)
                                'End If
                                
                            End If
                    End Select

                    


                End If
            Case mnu_ADD

        End Select
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                Dim strDocNum, strDocType As String
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Dim bisObj As SAPbouiCOM.BusinessObject = Form.BusinessObject
                Dim uid As String = bisObj.Key
                Dim BP1 As SAPbobsCOM.ServiceCalls  '= oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                Select Case BusinessObjectInfo.FormTypeEx
                    Case frm_ServiceCall
                        BP1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
                End Select
                If BP1.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                    strDocNum = BP1.ServiceCallID
                    CreateAPinvoice(BP1.ServiceCallID)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        If oForm.TypeEx = frm_ServiceCall Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    '    Dim otest As SAPbobsCOM.Recordset
                    '    Dim strString As String
                    '    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    strString = "SELECT *  FROM OSCL T0 WHERE isnull(U_Z_RecType,'N') <> 'N' and  isnull(U_Z_IsDuplicate,'N')='N' and Status=-1 and  T0.[DocNum] =" & oApplication.Utilities.getEdittextvalue(oForm, "540000180")
                    '    otest.DoQuery(strString)
                    '    If otest.Fields.Item(0).Value > 0 Then
                    '        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    '        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    '        oCreationPackage.UniqueID = "Dup"
                    '        oCreationPackage.String = "Duplicate Service Call"
                    '        oCreationPackage.Enabled = True
                    '        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    '        oMenus = oMenuItem.SubMenus
                    '        oMenus.AddEx(oCreationPackage)
                    '    End If
                    'End If
                Catch ex As Exception
                    ' MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '    oApplication.SBO_Application.Menus.RemoveEx("Dup")
                    'End If
                Catch ex As Exception
                    ' MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.BeforeAction
                Case True
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "1" Then
                        oMode = pVal.FormMode
                    End If
                Case False
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                        Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Select Case pVal.ItemUID
                                Case "1"

                            End Select

                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            oCFLEvent = pVal
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region


End Class
