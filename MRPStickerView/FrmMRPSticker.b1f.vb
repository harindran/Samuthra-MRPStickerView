Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SAPbobsCOM
Imports System.IO
Imports System.Windows.Forms

Namespace MRPStickerView
    <FormAttribute("MRP", "FrmMRPSticker.b1f")>
    Friend Class FrmMRPSticker
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("100").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lFDate").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("txtfdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lTDate").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("txttodate").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lWhse").Specific, SAPbouiCOM.StaticText)
            Me.StaticText3 = CType(Me.GetItem("lStart").Specific, SAPbouiCOM.StaticText)
            Me.StaticText4 = CType(Me.GetItem("lEnd").Specific, SAPbouiCOM.StaticText)
            Me.StaticText5 = CType(Me.GetItem("lPCode").Specific, SAPbouiCOM.StaticText)
            Me.Button2 = CType(Me.GetItem("btnMul").Specific, SAPbouiCOM.Button)
            Me.EditText2 = CType(Me.GetItem("txtPCode").Specific, SAPbouiCOM.EditText)
            Me.Button3 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.Button)
            Me.EditText3 = CType(Me.GetItem("Serstart").Specific, SAPbouiCOM.EditText)
            Me.EditText4 = CType(Me.GetItem("Serend").Specific, SAPbouiCOM.EditText)
            Me.EditText5 = CType(Me.GetItem("txtwhse").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("MRP", 0)
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                Dim RptName As String = ""
                Dim objrs As SAPbobsCOM.Recordset
                objrs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                RptName = "select Distinct ""U_RName"",""Code"" from ""@MRP_DATA"" Order by ""Code"" Desc"
                objrs.DoQuery(RptName)
                If objrs.RecordCount > 0 Then
                    RptName = objrs.Fields.Item("U_RName").Value
                    PassingParametersToReport(RptName)
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Public Sub GetCrystalReportFile(ByVal RDOCCode As String, ByVal outFileName As String)
            Try
                Dim oBlobParams As SAPbobsCOM.BlobParams = objaddon.objcompany.GetCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams)

                oBlobParams.Table = "RDOC"
                oBlobParams.Field = "Template"
                oBlobParams.FileName = "D:\Chitra\HRMS\Rajesh\JAN13\ReportByVinod\New PaySlip.rpt"
                Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = oBlobParams.BlobTableKeySegments.Add()
                oKeySegment.Name = "DocCode"
                oKeySegment.Value = RDOCCode

                Dim oBlob As SAPbobsCOM.Blob = objaddon.objcompany.GetCompanyService().GetBlob(oBlobParams)
                Dim sContent As String = oBlob.Content

                Dim buf() As Byte = Convert.FromBase64String(sContent)
                Using oFile As New System.IO.FileStream(outFileName, System.IO.FileMode.Create)
                    oFile.Write(buf, 0, buf.Length)
                    oFile.Close()
                End Using
            Catch ex As Exception
                Throw ex
            End Try


        End Sub

        'Private Function LoadCrViewer(ByVal crxReport As String) As Boolean
        '    'CRAXDDRT.Report
        '    Dim SBOFormCreationParams As SAPbouiCOM.FormCreationParams
        '    Dim SBOCRViewer As SAPbouiCOM.ActiveX
        '    Dim SBOForm As SAPbouiCOM.Form
        '    Dim SBOItem As SAPbouiCOM.Item

        '    'Add CRViewer item
        '    SBOItem = SBOForm.Items.Add("XX_CR01", SAPbouiCOM.BoFormItemTypes.it_ACTIVE_X)
        '    SBOItem.Left = 0
        '    SBOItem.Top = 0
        '    SBOItem.Width = SBOForm.ClientWidth
        '    SBOItem.Height = SBOForm.ClientHeight

        '    ' Create the new activeX control
        '    SBOCRViewer = SBOItem.Specific

        '    SBOCRViewer.ClassID = "CrystalReports13.ActiveXReportViewer.1"

        '    Dim SBOCRViewerOBJ
        '    SBOCRViewerOBJ = SBOCRViewer.Object
        '    SBOCRViewerOBJ.EnablePrintButton = False

        '    Dim MyProcs() As Process
        '    Dim i, ID As Integer
        '    Dim a As System.IntPtr

        '    SBOCRViewerOBJ.ViewReport()

        '    SBOForm.Visible = True

        '    Return True

        'End Function

        Private Sub TestLayout()
            Try

                'Dim oCmpSrv As SAPbobsCOM.CompanyService
                Dim oReportLayoutService As ReportLayoutsService
                Dim oReportLayout As ReportLayout
                Dim oReportLayoutParam As ReportLayoutParams

                'Get report layout service
                'oCmpSrv = objaddon.objcompany.GetCompanyService
                oReportLayoutService = objaddon.objcompany.GetCompanyService.GetBusinessService(ServiceTypes.ReportLayoutsService)

                'Set parameters
                oReportLayoutParam = oReportLayoutService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams)
                oReportLayoutParam.LayoutCode = "RDR10003"
                'Get report layout
                oReportLayout = oReportLayoutService.GetReportLayout(oReportLayoutParam)

                'Add report layout
                oReportLayoutService.AddReportLayout(oReportLayout)
                Try
                    Dim oNewReportParams As ReportLayoutParams = oReportLayoutService.AddReportLayout(oReportLayout)
                    'newReportCode = oNewReportParams.LayoutCode
                Catch err As System.Exception
                    Dim errMessage As String = err.Message
                    Return
                End Try

                Dim rptFilePath As String = "D:\Chitra\HRMS\Rajesh\JAN13\ReportByVinod\New PaySlip.rpt"
                Dim oCompanyService As CompanyService = objaddon.objcompany.GetCompanyService()
                Dim oBlobParams As BlobParams = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), BlobParams)
                oBlobParams.Table = "RDOC"
                oBlobParams.Field = "Template"
                Dim oKeySegment As BlobTableKeySegment = oBlobParams.BlobTableKeySegments.Add()
                oKeySegment.Name = "DocCode"
                oKeySegment.Value = "RDR10003" 'newReportCode
                Dim oBlob As Blob = CType(oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob), Blob)
                Dim oFile As FileStream = New FileStream(rptFilePath, System.IO.FileMode.Open)
                Dim fileSize As Integer = CInt(oFile.Length)
                Dim buf As Byte() = New Byte(fileSize - 1) {}
                oFile.Read(buf, 0, fileSize)
                oFile.Close()
                oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)

                Try
                    oCompanyService.SetBlob(oBlobParams, oBlob)
                Catch ex As System.Exception
                    Dim errmsg As String = ex.Message
                    MsgBox(errmsg)
                End Try
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub

        Private Sub PassingParametersToReport(ByVal ReportName As String)
            Try
                Dim strQuery As String, FormID, UniqueID, FieldQuery As String
                Dim objCRform, objCRLine As SAPbouiCOM.Form
                Dim objrs, objRSField As SAPbobsCOM.Recordset
                objrs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objRSField = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                Dim oCombo As SAPbouiCOM.ComboBox
                Dim oEdit As SAPbouiCOM.EditText
                Dim oMatrix As SAPbouiCOM.Matrix
                Dim chkbox As SAPbouiCOM.CheckBox
                Dim ProdCode() As String
                Dim FormExist As Boolean = False
                Dim IRow As Integer = 1
                strQuery = "Select ""MenuUID"" from OCMN where ""Name""='" & ReportName & "' and ""Type""='C';  "
                objrs.DoQuery(strQuery)
                FormID = "410000100"
                For Each uid As SAPbouiCOM.Form In objaddon.objapplication.Forms
                    If uid.TypeEx = FormID Then
                        FormExist = True
                        UniqueID = objaddon.objapplication.Forms.GetForm(FormID, 0).UniqueID.ToString
                        objaddon.objapplication.Forms.Item(UniqueID).Close()
                        Exit For
                    End If
                Next
                objaddon.objapplication.ActivateMenuItem(objrs.Fields.Item("MenuUID").Value)

                objCRform = objaddon.objapplication.Forms.ActiveForm 'GetForm("410000100", 0)
                objCRform.Visible = False
                FieldQuery = "select * from ""@MRP_DATA"" Order by ""Code"" Desc "
                objRSField.DoQuery(FieldQuery)
                If objRSField.RecordCount > 0 Then
                    oEdit = objCRform.Items.Item(objRSField.Fields.Item("U_FDate").Value).Specific  'FromDate
                    oEdit.Value = EditText0.Value

                    oEdit = objCRform.Items.Item(objRSField.Fields.Item("U_TDate").Value).Specific  'ToDate
                    oEdit.Value = EditText1.Value

                    oCombo = objCRform.Items.Item(objRSField.Fields.Item("U_Whse").Value).Specific  'Warehouse
                    oCombo.Select(EditText5.Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

                    oCombo = objCRform.Items.Item(objRSField.Fields.Item("U_SStart").Value).Specific  'Serial Starts
                    oCombo.Select(EditText3.Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

                    oCombo = objCRform.Items.Item(objRSField.Fields.Item("U_SEnd").Value).Specific  'Serial Ends
                    oCombo.Select(EditText4.Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

                    ProdCode = EditText2.Value.Split(New Char() {","c})
                    objCRform.Items.Item(objRSField.Fields.Item("U_BtnProd").Value).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    objCRLine = objaddon.objapplication.Forms.ActiveForm
                    objCRLine.Visible = False
                    'objform = objaddon.objapplication.Forms.GetForm("410000003", 0)
                    oMatrix = objCRLine.Items.Item(objRSField.Fields.Item("U_ProdMat").Value).Specific
                    For i As Integer = 0 To ProdCode.Length - 1
                        For j As Integer = 1 To oMatrix.VisualRowCount
                            chkbox = oMatrix.Columns.Item(objRSField.Fields.Item("U_ChkProdVal").Value).Cells.Item(j).Specific
                            If oMatrix.Columns.Item(objRSField.Fields.Item("U_ProdVal").Value).Cells.Item(j).Specific.String = ProdCode(i) Then
                                chkbox.Checked = True
                                Exit For
                            End If
                        Next
                    Next
                    objCRLine.Items.Item(objRSField.Fields.Item("U_ProdOK").Value).Click()
                    'objform = objaddon.objapplication.Forms.GetForm("410000100", 0)
                    objCRform.Items.Item("1").Click()
                Else
                    objaddon.objapplication.StatusBar.SetText("Please update the MRP UDT...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
                objrs = Nothing
                objRSField = Nothing
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents EditText2 As SAPbouiCOM.EditText

        Private Sub Button2_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                Dim FromDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Dim ToDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'ProductQuery = "Select Distinct 'N' ""Select"",D.""Code"", D.""Name"" From OSRI A INNER JOIN SRI1 B on A.""SysSerial""=B.""SysSerial"" and A.""ItemCode""=B.""ItemCode""  join OITM C on C.""ItemCode""=B.""ItemCode"" join ""@PRODGRP"" D on D.""Code""=C.""U_ProductGroup"" "
                'ProductQuery += vbCrLf + " where B.""DocDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' and A.""WhsCode""='" & EditText5.Value & "' and  A.""Status""=0 and C.""U_ProductGroup""<>'';"
                ProductQuery = "Select Distinct 'N' ""Select"",D.""Code"", D.""Name"" From OSRI A INNER JOIN SRI1 B on A.""SysSerial""=B.""SysSerial"" and A.""ItemCode""=B.""ItemCode""  join OITM C on C.""ItemCode""=B.""ItemCode"" join ""@PRODGRP"" D on D.""Code""=C.""U_ProductGroup"" "
                ProductQuery += vbCrLf + " where A.""InDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' and A.""WhsCode""='" & EditText5.Value & "' and  A.""Status""=0 and C.""U_ProductGroup""<>'';"
                FrmMultiSel = objaddon.objapplication.Forms.ActiveForm
                If Not objaddon.FormExist("MULPROD") Then
                    Dim Multiselect As New FrmMultiSelect
                    Multiselect.Show()
                End If

            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents Button3 As SAPbouiCOM.Button

        Private Sub Button3_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try
                objform.Freeze(True)
                EditText0.Value = ""
                EditText1.Value = ""
                EditText5.Value = ""
                EditText3.Value = ""
                EditText4.Value = ""
                EditText2.Value = ""
                objform.ActiveItem = "txtfdate"
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try

        End Sub
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents EditText4 As SAPbouiCOM.EditText

        Private Sub EditText3_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText3.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            If EditText5.Value = "" Then
                objaddon.objapplication.SetStatusBarMessage("Please Select Warehouse...", SAPbouiCOM.BoMessageTime.bmt_Short, False) : BubbleEvent = False : Exit Sub
            End If
            If EditText2.Value = "" Then
                objaddon.objapplication.SetStatusBarMessage("Please Select Product Code...", SAPbouiCOM.BoMessageTime.bmt_Short, False) : BubbleEvent = False : Exit Sub
            End If
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("SerialS")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                Dim Query As String
                Dim ProdCode() As String
                Dim PCode As String = "'"
                Link_Value = "MRP"
                'oCFL.SetConditions(oEmptyConds)
                'oConds = oCFL.GetConditions()
                'oCond = oConds.Add()
                'oCond.Alias = "WhsCode"
                'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                'oCond.CondVal = EditText5.Value
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                'oCond = oConds.Add()
                'oCond.Alias = "CreateDate"
                'oCond.CondVal = EditText0.Value
                'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_BETWEEN
                'oCond.CondEndVal = EditText1.Value
                'oCFL.SetConditions(oConds)
                Dim rsetCFL As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                Dim FromDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Dim ToDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                ProdCode = EditText2.Value.Split(New Char() {","c})
                For i As Integer = 0 To ProdCode.Length - 1
                    PCode += ProdCode(i) + "','"
                Next
                PCode = PCode.Remove(PCode.Length - 2)
                'Query = "Select Distinct A.""IntrSerial"" From OSRI A Left JOIN SRI1 B on A.""SysSerial""=B.""SysSerial"" and A.""ItemCode""=B.""ItemCode"" Left join OITM C on C.""ItemCode""=B.""ItemCode"" where "
                'Query += vbCrLf + "B.""DocDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' and A.""WhsCode""='" & EditText5.Value & "' and A.""Status""=0 and C.""U_ProductGroup"" in (" & PCode & ") Group BY A.""IntrSerial"" order by A.""IntrSerial"""
                Query = "Select Distinct A.""IntrSerial"" From OSRI A Left JOIN SRI1 B on A.""SysSerial""=B.""SysSerial"" and A.""ItemCode""=B.""ItemCode"" Left join OITM C on C.""ItemCode""=B.""ItemCode"" where "
                Query += vbCrLf + "A.""InDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' and A.""WhsCode""='" & EditText5.Value & "' and A.""Status""=0 and C.""U_ProductGroup"" in (" & PCode & ") Group BY A.""IntrSerial"" order by A.""IntrSerial"""
                rsetCFL.DoQuery(Query)
                rsetCFL.MoveFirst()
                For i As Integer = 1 To rsetCFL.RecordCount
                    If i = (rsetCFL.RecordCount) Then
                        oCond = oConds.Add()
                        oCond.Alias = "IntrSerial"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                    Else
                        oCond = oConds.Add()
                        oCond.Alias = "IntrSerial"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    rsetCFL.MoveNext()
                Next
                If rsetCFL.RecordCount = 0 Then
                    oCond = oConds.Add()
                    oCond.Alias = "IntrSerial"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = ""
                End If
                oCFL.SetConditions(oConds)
                
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub
        Private WithEvents EditText5 As SAPbouiCOM.EditText

        Private Sub EditText5_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText5.ChooseFromListBefore
            Try
                If EditText0.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Please Select FromDate...", SAPbouiCOM.BoMessageTime.bmt_Short, False) : BubbleEvent = False : Exit Sub
                End If
                If EditText1.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Please Select ToDate...", SAPbouiCOM.BoMessageTime.bmt_Short, False) : BubbleEvent = False : Exit Sub
                End If
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_Whse")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim Query As String
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                Dim rsetCFL As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                Dim FromDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Dim ToDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'Query = "select distinct ""WhsCode"" from SRI1 where ""DocDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "'"
                'Query = "select distinct ""WhsCode"" from OSRI where ""InDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "'"
                'Query = "Select Distinct A.""WhsCode"" From OSRI A Left JOIN SRI1 B on A.""SysSerial""=B.""SysSerial"" and A.""ItemCode""=B.""ItemCode"" where B.""DocDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' and  A.""Status""=0 "
                Query = "Select Distinct A.""WhsCode"" From OSRI A Left JOIN SRI1 B on A.""SysSerial""=B.""SysSerial"" and A.""ItemCode""=B.""ItemCode"" where A.""InDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' and  A.""Status""=0 "
                rsetCFL.DoQuery(Query)
                rsetCFL.MoveFirst()
                For i As Integer = 1 To rsetCFL.RecordCount
                    If i = (rsetCFL.RecordCount) Then
                        oCond = oConds.Add()
                        oCond.Alias = "WhsCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL   'U_SalesWhs
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        oCond = oConds.Add()
                        oCond.Alias = "U_MRPPrint"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "Yes"
                    Else
                        oCond = oConds.Add()
                        oCond.Alias = "WhsCode"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        oCond = oConds.Add()
                        oCond.Alias = "U_MRPPrint"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "Yes"
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    rsetCFL.MoveNext()
                Next
                If rsetCFL.RecordCount = 0 Then
                    oCond = oConds.Add()
                    oCond.Alias = "WhsCode"
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                    oCond.CondVal = ""
                End If
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Choose FromList Filter Global Fun. Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try


        End Sub

        Private Sub EditText4_ChooseFromListBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles EditText4.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            If EditText5.Value = "" Then
                objaddon.objapplication.SetStatusBarMessage("Please Select Warehouse...", SAPbouiCOM.BoMessageTime.bmt_Short, False) : BubbleEvent = False : Exit Sub
            End If
            If EditText2.Value = "" Then
                objaddon.objapplication.SetStatusBarMessage("Please Select Product Code...", SAPbouiCOM.BoMessageTime.bmt_Short, False) : BubbleEvent = False : Exit Sub
            End If

            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("SerialE")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                Dim Query As String
                'oCFL.SetConditions(oEmptyConds)
                'oConds = oCFL.GetConditions()
                'oCond = oConds.Add()
                'oCond.Alias = "WhsCode"
                'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                'oCond.CondVal = EditText5.Value
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                'oCond = oConds.Add()
                'oCond.Alias = "CreateDate"
                'oCond.CondVal = EditText0.Value
                'oCond.Operation = SAPbouiCOM.BoConditionOperation.co_BETWEEN
                'oCond.CondEndVal = EditText1.Value
                'oCFL.SetConditions(oConds)
                Link_Value = "MRP"
                Dim ProdCode() As String
                Dim PCode As String = "'"
                Dim rsetCFL As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                Dim FromDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Dim ToDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                ProdCode = EditText2.Value.Split(New Char() {","c})
                For i As Integer = 0 To ProdCode.Length - 1
                    PCode += ProdCode(i) + "','"
                Next
                PCode = PCode.Remove(PCode.Length - 2)
                'Query = "Select Distinct A.""IntrSerial"" From OSRI A Left JOIN SRI1 B on A.""SysSerial""=B.""SysSerial"" and A.""ItemCode""=B.""ItemCode"" Left join OITM C on C.""ItemCode""=B.""ItemCode"" where "
                'Query += vbCrLf + "B.""DocDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' and A.""WhsCode""='" & EditText5.Value & "' and  A.""Status""=0 and C.""U_ProductGroup"" in (" & PCode & ") Group BY A.""IntrSerial"" order by A.""IntrSerial"" "
                Query = "Select Distinct A.""IntrSerial"" From OSRI A Left JOIN SRI1 B on A.""SysSerial""=B.""SysSerial"" and A.""ItemCode""=B.""ItemCode"" Left join OITM C on C.""ItemCode""=B.""ItemCode"" where "
                Query += vbCrLf + "A.""InDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' and A.""WhsCode""='" & EditText5.Value & "' and  A.""Status""=0 and C.""U_ProductGroup"" in (" & PCode & ") Group BY A.""IntrSerial"" order by A.""IntrSerial"" "
                rsetCFL.DoQuery(Query)
                rsetCFL.MoveFirst()
                For i As Integer = 1 To rsetCFL.RecordCount
                    If i = (rsetCFL.RecordCount) Then
                        oCond = oConds.Add()
                        oCond.Alias = "IntrSerial"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                    Else
                        oCond = oConds.Add()
                        oCond.Alias = "IntrSerial"
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    rsetCFL.MoveNext()
                Next
                If rsetCFL.RecordCount = 0 Then
                    oCond = oConds.Add()
                    oCond.Alias = "IntrSerial"
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.CondVal = ""
                End If
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Sub EditText3_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText3.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText3.Value = pCFL.SelectedObjects.Columns.Item("IntrSerial").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub EditText4_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText4.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText4.Value = pCFL.SelectedObjects.Columns.Item("IntrSerial").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub EditText5_ChooseFromListAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText5.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        EditText5.Value = pCFL.SelectedObjects.Columns.Item("WhsCode").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                Dim objRS, objRS1, objRS2 As SAPbobsCOM.Recordset
                Dim StrQuery, GetQuery, GetNull, ErrorMessage As String
                Dim ItemCode As String = ""
                Dim ELog As Boolean = False
                Dim ErrorCount As Integer = 0
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objRS1 = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objRS2 = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                If EditText0.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("From Date is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText1.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("To Date is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText5.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Warehouse Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText2.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Product Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText3.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Starting Serial is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText4.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Ending Serial is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If

                Dim FromDate As Date = Date.ParseExact(EditText0.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                Dim ToDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'Getting ItemCodes based on the serial
                GetQuery = "select distinct ""ItemCode"" from OSRN where ""DistNumber"" between '" & EditText3.Value & "' and '" & EditText4.Value & "' and ""InDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and  '" & ToDate.ToString("yyyyMMdd") & "' "
                objRS.DoQuery(GetQuery)
                'Concatenating itemcodes
                If objRS.RecordCount > 0 Then
                    For i As Integer = 0 To objRS.RecordCount - 1
                        If i = 0 Then
                            ItemCode = "'" & objRS.Fields.Item("ItemCode").Value & "'"
                        Else
                            ItemCode = ItemCode & ",'" & objRS.Fields.Item("ItemCode").Value & "'"
                        End If
                        objRS.MoveNext()
                    Next
                    'Validating Price
                    StrQuery = "Select ""ItemCode"",case when ifnull(""Price"",0)=0 then 'Price' else '0' end as ""Price"" from ITM1 where ""PriceList""='2' and ""ItemCode"" in (" & ItemCode & ")"
                    objRS2.DoQuery(StrQuery)
                    If objRS2.RecordCount > 0 Then
                        For Price As Integer = 0 To objRS2.RecordCount - 1
                            If objRS2.Fields.Item("Price").Value <> "0" Then
                                ErrorCount += 1
                                objaddon.objapplication.SetStatusBarMessage("Please fill the MRP ...Item Code: " & objRS2.Fields.Item("ItemCode").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            End If
                            objRS2.MoveNext()
                        Next
                        If ErrorCount > 0 Then
                            objaddon.objapplication.SetStatusBarMessage("Items without price in the MRP price list while generating the report....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        End If
                        ErrorCount = 0
                    End If
                    'Validating UDF fields of itemmaster
                    GetNull = "Select distinct T0.""ItemCode"",T0.""SalPackUn"",T0.""U_ProductGroup"",case  when ifnull(T0.""U_MinBore"",'')='' then 'Minimum Bore Size(mm)' else '0' end as ""U_MinBore"", "
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_MRPTit1"",'')='' then 'MRP Title 1' else '0' end as ""U_MRPTit1"","
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_MRPTit2"",'')='' then 'MRP Title 2' else '0' end as ""U_MRPTit2"","
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_MRPtit3"",'')='' then 'MRP Title 3' else '0' end as ""U_MRPtit3"","
                    GetNull += vbCrLf + " case when ifnull(T0.""U_kWHP"",'')='' then 'kW/HP' else '0' end as ""U_kWHP"","
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_Phase"",'')='' then 'Phase' else '0' end as ""U_Phase"","
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_Stages"",0)=0 then 'No. of Stages' else '0' end as ""U_Stages"","
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_MRPType"",'')='' then 'MRP Type' else '0' end as ""U_MRPType"","
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_DelSiz"",0)=0 then 'Delivery Size(mm)' else '0' end as ""U_DelSiz"" , "
                    GetNull += vbCrLf + " case when ifnull(T2.""MnfDate"",'')='' then 'Manufacturing Date' else '0' end as ""MnfDate"", "
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_GrossWetP"",0)=0 then 'Gross Weight (Kg.) Pump' else '0' end as ""U_GrossWetP"","
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_GrossWetMot"",0)=0 then 'Gross Weight (Kg.) Motor' else '0' end as ""U_GrossWetMot"","
                    GetNull += vbCrLf + "  case when ifnull(T0.""U_GrossWetPK"",0)=0 then 'Gross Weight (Kg.) Dual Pk' else '0' end as ""U_GrossWetPK""  "
                    GetNull += vbCrLf + " from OITM T0 join ITM1 T1 on T0.""ItemCode""=T1.""ItemCode"" join OSRN T2 on T1.""ItemCode""=T2.""ItemCode"" where T0.""ItemCode"" in (" & ItemCode & ") "
                    GetNull += vbCrLf + "  and T2.""DistNumber"" between '" & EditText3.Value & "' and '" & EditText4.Value & "' "
                    objRS1.DoQuery(GetNull)
                    If objRS1.RecordCount > 0 Then
                        ErrorMessage = ""
                        For Rec As Integer = 0 To objRS1.RecordCount - 1
                            If objRS1.Fields.Item("U_MinBore").Value <> "0" Or objRS1.Fields.Item("U_MRPTit1").Value <> "0" Or objRS1.Fields.Item("U_MRPTit2").Value <> "0" Or objRS1.Fields.Item("U_MRPtit3").Value <> "0" Or objRS1.Fields.Item("U_kWHP").Value <> "0" Or objRS1.Fields.Item("U_Phase").Value <> "0" Or objRS1.Fields.Item("U_Stages").Value <> "0" Or objRS1.Fields.Item("U_DelSiz").Value <> "0" Or objRS1.Fields.Item("U_MRPType").Value <> "0" Or objRS1.Fields.Item("MnfDate").Value <> "0" Or objRS1.Fields.Item("U_GrossWetP").Value <> "0" Or objRS1.Fields.Item("U_GrossWetMot").Value <> "0" Or objRS1.Fields.Item("U_GrossWetPK").Value <> "0" Then
                                'ErrorCount += 1
                                ErrorMessage = "ItemCode: " & objRS1.Fields.Item("ItemCode").Value.ToString
                            End If
                            If ErrorMessage <> "" Then
                                If objRS1.Fields.Item("U_MinBore").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_MinBore").Value.ToString
                                End If
                                If objRS1.Fields.Item("U_MRPType").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_MRPType").Value.ToString
                                End If
                                If objRS1.Fields.Item("U_MRPTit1").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_MRPTit1").Value.ToString
                                End If
                                If objRS1.Fields.Item("U_MRPTit2").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_MRPTit2").Value.ToString
                                End If
                                If objRS1.Fields.Item("U_MRPtit3").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_MRPtit3").Value.ToString
                                End If
                                If objRS1.Fields.Item("U_kWHP").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_kWHP").Value.ToString
                                End If
                                If objRS1.Fields.Item("U_Phase").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_Phase").Value.ToString
                                End If
                                If objRS1.Fields.Item("U_Stages").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_Stages").Value.ToString
                                End If
                                If objRS1.Fields.Item("U_DelSiz").Value <> "0" Then
                                    ErrorCount += 1
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_DelSiz").Value.ToString
                                End If
                            End If
                            If (objRS1.Fields.Item("U_ProductGroup").Value = "S4" Or objRS1.Fields.Item("U_ProductGroup").Value = "S6") And objRS1.Fields.Item("SalPackUn").Value = "0.5" Then
                                If objRS1.Fields.Item("U_GrossWetP").Value <> "0" Then
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_GrossWetP").Value.ToString
                                    ErrorCount += 1
                                End If
                                If objRS1.Fields.Item("U_GrossWetMot").Value <> "0" Then
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_GrossWetMot").Value.ToString
                                    ErrorCount += 1
                                End If
                            End If
                            If ((objRS1.Fields.Item("U_ProductGroup").Value = "S4" Or objRS1.Fields.Item("U_ProductGroup").Value = "S6") And objRS1.Fields.Item("SalPackUn").Value = "1") Or (objRS1.Fields.Item("U_ProductGroup").Value = "OWS" Or objRS1.Fields.Item("U_ProductGroup").Value = "SSP") Then
                                If objRS1.Fields.Item("U_GrossWetPK").Value <> "0" Then
                                    ErrorMessage += " Field: " & objRS1.Fields.Item("U_GrossWetPK").Value.ToString
                                    ErrorCount += 1
                                End If
                            End If
                            If objRS1.Fields.Item("MnfDate").Value <> "0" Then
                                ErrorMessage += " Field: " & objRS1.Fields.Item("MnfDate").Value.ToString
                                ErrorCount += 1
                            End If

                            If ErrorCount > 0 Then
                                objaddon.objapplication.StatusBar.SetText("Please fill the fields of the " & ErrorMessage.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                ErrorMessage = ""
                                ErrorCount = 0
                                ELog = True
                            End If
                            objRS1.MoveNext()
                        Next
                        If ELog Then
                            objaddon.objapplication.SetStatusBarMessage("Please review the system messages log & update the mentioned field values...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        ErrorCount = 0
                    End If

                    objRS = Nothing
                    objRS1 = Nothing
                    objRS2 = Nothing
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try

        End Sub

        Private Sub EditText4_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText4.LostFocusAfter
            Try
                objform.ActiveItem = "txtfdate"
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button2_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button2.ClickBefore
            Try
                If EditText0.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Please Select From Date to get Product Code...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText1.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Please Select To Date to get Product Code...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText5.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Please Select Warehouse Code to get Product Code...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

    End Class
End Namespace
