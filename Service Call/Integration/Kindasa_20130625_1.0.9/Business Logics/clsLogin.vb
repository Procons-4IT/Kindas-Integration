Public Class clsLogin
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
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Private objMatrix As SAPbouiCOM.Matrix
    Private objForm As SAPbouiCOM.Form
    Private oColumn As SAPbouiCOM.Column
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_Login, frm_Login)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        databind(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        oForm.Freeze(False)
    End Sub

#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim oTestRS As SAPbobsCOM.Recordset
            Dim oDBDataSrc As SAPbouiCOM.DBDataSource
            oTestRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTestRS.DoQuery("Select * from [@Z_INT_LOGIN]")
            If oTestRS.RecordCount > 0 Then
                oApplication.Utilities.setEdittextvalue(aForm, "4", oTestRS.Fields.Item(2).Value)
                oApplication.Utilities.setEdittextvalue(aForm, "6", oTestRS.Fields.Item(3).Value)
                oApplication.Utilities.setEdittextvalue(aForm, "8", oTestRS.Fields.Item(4).Value)
                oApplication.Utilities.setEdittextvalue(aForm, "10", oTestRS.Fields.Item(5).Value)
                oApplication.Utilities.setEdittextvalue(aForm, "12", oTestRS.Fields.Item(6).Value)
                oApplication.Utilities.setEdittextvalue(aForm, "14", oTestRS.Fields.Item(7).Value)
            End If
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim oTest As SAPbobsCOM.Recordset
        Dim strCode As String

        If oApplication.Utilities.getEdittextvalue(aform, "4") = "" Then
            oApplication.Utilities.Message("Linked server details missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        If oApplication.Utilities.getEdittextvalue(aform, "6") = "" Then
            oApplication.Utilities.Message("Linked server SQL DB missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oApplication.Utilities.getEdittextvalue(aform, "8") = "" Then
            oApplication.Utilities.Message("Linked server User ID missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oApplication.Utilities.getEdittextvalue(aform, "10") = "" Then
            oApplication.Utilities.Message("Linked server Password missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oApplication.Utilities.getEdittextvalue(aform, "12") = "" Then
            oApplication.Utilities.Message("Local server details missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oApplication.Utilities.getEdittextvalue(aform, "14") = "" Then
            oApplication.Utilities.Message("Local server User ID missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If


        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Delete from [@Z_INT_LOGIN]")

        oUserTable = oApplication.Company.UserTables.Item("Z_INT_LOGIN")
        strCode = oApplication.Utilities.getMaxCode("@Z_INT_LOGIN", "Code")
        oUserTable.Code = strCode
        oUserTable.Name = strCode
        oUserTable.UserFields.Fields.Item("U_Z_SERVER").Value = oApplication.Utilities.getEdittextvalue(aform, "4")
        oUserTable.UserFields.Fields.Item("U_Z_SQLDB").Value = oApplication.Utilities.getEdittextvalue(aform, "6")
        oUserTable.UserFields.Fields.Item("U_Z_SQLUID").Value = oApplication.Utilities.getEdittextvalue(aform, "8")
        oUserTable.UserFields.Fields.Item("U_Z_SQLPWD").Value = oApplication.Utilities.getEdittextvalue(aform, "10")
        oUserTable.UserFields.Fields.Item("U_Z_LOCALSQLDB").Value = oApplication.Utilities.getEdittextvalue(aform, "12")
        oUserTable.UserFields.Fields.Item("U_Z_LOCUID").Value = oApplication.Utilities.getEdittextvalue(aform, "14")
        If oUserTable.Add <> 0 Then
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        ElseIf oApplication.Utilities.linkedserverConnect() = True Then
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            databind(aform)
            Return True
        Else
            oApplication.Utilities.Message("Linked server connection failed. Please Check Linked server details...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
    End Function
#End Region





#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Login Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And pVal.ColUID = "V_0" And pVal.CharPressed <> 9 Then
                                    objMatrix = oForm.Items.Item("3").Specific
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" Then
                                    oForm.Freeze(True)
                                    objForm = oForm
                                    AddtoUDT1(oForm)
                                    oForm.Freeze(False)
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
                Case mnu_Login
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

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
