Imports SAPbouiCOM
Namespace MRPStickerView

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    'Case "SUBCTPO"
                    '    SubContractingPO_MenuEvent(pVal, BubbleEvent)
                    'Case "SUBBOM"
                    '    SubContractingBOM_MenuEvent(pVal, BubbleEvent)
                    'Case "65211"
                    '    ProductionOrder_MenuEvent(pVal, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                Else
                    Select Case pval.MenuUID
                        Case "1281"
                        Case Else
                    End Select
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "SubContractingPO"

        Private Sub SubContractingPO_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0, Matrix2, Matrix4, Matrix3, Matrix1 As SAPbouiCOM.Matrix
            Dim FolderInput, FolderOutput, FolderScrap, FolderRelDoc, FolderCosting As SAPbouiCOM.Folder
            Dim FolderID As String = ""
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("MtxinputN").Specific
                Matrix2 = objform.Items.Item("mtxreldoc").Specific
                Matrix4 = objform.Items.Item("MtxCosting").Specific
                Matrix3 = objform.Items.Item("mtxoutput").Specific
                Matrix1 = objform.Items.Item("mtxscrap").Specific
                FolderInput = objform.Items.Item("flrinput").Specific
                FolderOutput = objform.Items.Item("flroutput").Specific
                FolderScrap = objform.Items.Item("flrscrap").Specific
                FolderRelDoc = objform.Items.Item("flrreldoc").Specific
                FolderCosting = objform.Items.Item("flrcosting").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1292"
                        Case "1293"
                            'For i As Integer = Matrix0.VisualRowCount To 1 Step -1
                            '    Matrix0.Columns.Item("#").Cells.Item(i).Specific.String = i
                            'Next
                            'Matrix0.Columns.Item("#").Cells.Item(Matrix0.VisualRowCount - 1).Specific.String = "2" 'Matrix4.VisualRowCount + 1
                            'objaddon.objglobalmethods.RemoveLastrow(Matrix4, "#")
                            'objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                    End Select
                Else
                    If FolderInput.Selected = True Then
                        FolderID = "MtxinputN"
                    ElseIf FolderOutput.Selected = True Then
                        FolderID = "mtxoutput"
                    ElseIf FolderScrap.Selected = True Then
                        FolderID = "mtxscrap"
                    ElseIf FolderRelDoc.Selected = True Then
                        FolderID = "mtxreldoc"
                    ElseIf FolderCosting.Selected = True Then
                        FolderID = "MtxCosting"
                    End If
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("txtdocnum").Enabled = True
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtstat").Enabled = True
                            objform.Items.Item("txtcode").Enabled = True
                            objform.Items.Item("txtctper").Enabled = True
                            objform.Items.Item("txtsitem").Enabled = True
                            objform.Items.Item("docdate").Enabled = True
                            objform.Items.Item("deldate").Enabled = True
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtGINo").Enabled = True
                            objform.Items.Item("txtGRNo").Enabled = True
                            objform.Items.Item("TxtInvTr").Enabled = True
                            Matrix0.Item.Enabled = False
                            Matrix1.Item.Enabled = False
                            Matrix2.Item.Enabled = False
                            Matrix3.Item.Enabled = False
                            Matrix4.Item.Enabled = False
                            'objform.ActiveItem = "txtdocnum"
                            objform.Items.Item("txtdocnum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                objform = objaddon.objapplication.Forms.GetForm("SUBCTPO", 0)
                            End If
                            Exit Sub
                        Case "1282" ' Add Mode
                            objform.Items.Item("btngendoc").Enabled = False
                            objform.Items.Item("btnload").Enabled = False
                            objform.Items.Item("BtnView").Enabled = False
                            objform.Items.Item("BtnInv").Enabled = False
                            objform.Items.Item("BtnGIssue").Enabled = False
                            objform.Items.Item("btnOutput").Enabled = False
                            objform.Items.Item("BtnScrap").Enabled = False
                            objform.Items.Item("posdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objform.Items.Item("docdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                            FolderScrap.Item.Click(BoCellClickType.ct_Regular)
                            FolderInput.Item.Click(BoCellClickType.ct_Regular)
                            objform.Items.Item("txtdocnum").Specific.string = objaddon.objglobalmethods.GetNextDocNum_Value("@MIPL_OPOR")
                            objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_OPOR")

                        Case "1288", "1289", "1290", "1291"
                            For j = 1 To Matrix4.VisualRowCount
                                If Matrix4.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And Matrix4.Columns.Item("Stat").Cells.Item(j).Specific.String = "C" Then
                                    Matrix4.CommonSetting.SetRowEditable(j, False)
                                End If
                            Next
                            For j = 1 To Matrix3.VisualRowCount
                                If Matrix3.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And Matrix3.Columns.Item("Colstat").Cells.Item(j).Specific.String = "C" Then
                                    Matrix3.CommonSetting.SetRowEditable(j, False)
                                End If
                            Next
                            For j = 1 To Matrix1.VisualRowCount
                                If Matrix1.Columns.Item("Code").Cells.Item(j).Specific.String <> "" And Matrix1.Columns.Item("Colstat").Cells.Item(j).Specific.String = "C" Then
                                    Matrix1.CommonSetting.SetRowEditable(j, False)
                                End If
                            Next

                        Case "1293"
                            Select Case FolderID
                                Case "MtxinputN"
                                    DeleteRow(Matrix0, "@MIPL_POR1")
                                Case "mtxoutput"
                                    DeleteRow(Matrix3, "@MIPL_POR2")
                                Case "mtxscrap"
                                    DeleteRow(Matrix1, "@MIPL_POR3")
                                Case "mtxreldoc"
                                    DeleteRow(Matrix2, "@MIPL_POR4")
                                Case "MtxCosting"
                                    DeleteRow(Matrix4, "@MIPL_POR5")
                            End Select
                        Case "1292"
                            Select Case FolderID
                                Case "MtxinputN"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                                    'objaddon.objglobalmethods.SetCellEdit(Matrix0, True)
                                Case "mtxoutput"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "Code", "#")
                                Case "mtxscrap"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix1, "Code", "#")
                                Case "mtxreldoc"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "Code", "#")
                                Case "MtxCosting"
                                    objaddon.objglobalmethods.Matrix_Addrow(Matrix4, "Code", "#")
                            End Select
                        Case "1304" 'Refresh
                            Dim OnHand As String
                            objform.Mode = BoFormMode.fm_UPDATE_MODE
                            Select Case FolderID
                                Case "MtxinputN"
                                    For i As Integer = 1 To Matrix0.VisualRowCount
                                        OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String & "' and ""WhsCode""='" & Matrix0.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")
                                        Matrix0.Columns.Item("Instock").Cells.Item(i).Specific.String = OnHand
                                    Next
                                    'Case "mtxoutput"
                                    '    For i As Integer = 1 To Matrix3.VisualRowCount
                                    '        OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix3.Columns.Item("Code").Cells.Item(i).Specific.String & "' and ""WhsCode""='" & Matrix3.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")
                                    '        Matrix3.Columns.Item("Instock").Cells.Item(i).Specific.String = OnHand
                                    '    Next
                                    'Case "mtxscrap"
                                    '    For i As Integer = 1 To Matrix1.VisualRowCount
                                    '        OnHand = objaddon.objglobalmethods.getSingleValue("select ""OnHand"" from  oitw where ""ItemCode""='" & Matrix1.Columns.Item("Code").Cells.Item(i).Specific.String & "' and ""WhsCode""='" & Matrix1.Columns.Item("Whse").Cells.Item(i).Specific.String & "'")
                                    '        Matrix1.Columns.Item("Instock").Cells.Item(i).Specific.String = OnHand
                                    '    Next
                            End Select
                            If objform.Mode = BoFormMode.fm_UPDATE_MODE Then
                                objform.Items.Item("1").Click()
                            End If

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub


        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub
#End Region



        Private Sub SubContractingBOM_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Dim odbdsDetails As SAPbouiCOM.DBDataSource

            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIPL_BOM1")
                Matrix0 = objform.Items.Item("mtxBOM").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1293"
                            If Matrix0.VisualRowCount = 1 Then BubbleEvent = False
                        Case "1292"
                            'Try
                            '    If Matrix0.VisualRowCount > 0 Then
                            '        If odbdsDetails.GetValue("U_Itemcode", Matrix0.VisualRowCount - 1) = "" Then Exit Sub
                            '        objform.Freeze(True)
                            '        odbdsDetails.InsertRecord(odbdsDetails.Size)
                            '        odbdsDetails.SetValue("LineId", Matrix0.VisualRowCount, Matrix0.VisualRowCount + 1)
                            '        Matrix0.LoadFromDataSource()
                            '        objform.Freeze(False)
                            '    End If
                            'Catch ex As Exception

                            'End Try
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode                           
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtcode").Enabled = True
                            objform.Items.Item("mtxBOM").Enabled = False
                            objform.Items.Item("txtentry").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Case "1282" ' Add Mode
                            objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.getSingleValue("select Count(*)+1 ""DocEntry"" from ""@MIPL_OBOM""")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                        Case "1288", "1289", "1290", "1291"
                            'objform.Items.Item("btngendoc").Enabled = True
                            objaddon.objapplication.Menus.Item("1300").Activate()
                        Case "1293"
                            For i As Integer = Matrix0.VisualRowCount To 1 Step -1
                                Matrix0.Columns.Item("#").Cells.Item(i).Specific.String = i
                            Next
                            If objform.Mode = BoFormMode.fm_OK_MODE Then objform.Mode = BoFormMode.fm_UPDATE_MODE
                            objform.Update()
                            objform.Refresh()
                        Case "1292"
                            Try
                                If Matrix0.VisualRowCount > 0 Then
                                    If odbdsDetails.GetValue("U_Itemcode", Matrix0.VisualRowCount - 1) = "" Then Exit Sub
                                    objform.Freeze(True)
                                    odbdsDetails.InsertRecord(odbdsDetails.Size)
                                    odbdsDetails.SetValue("LineId", Matrix0.VisualRowCount, Matrix0.VisualRowCount + 1)
                                    Matrix0.LoadFromDataSource()
                                    objform.Freeze(False)
                                End If
                            Catch ex As Exception

                            End Try
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub ProductionOrder_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm

                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode 

                        Case "1282" ' Add Mode
                        Case "1288", "1289", "1290", "1291", "1304"
                           
                        Case "1293"

                        Case "1292"

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

    End Class
End Namespace