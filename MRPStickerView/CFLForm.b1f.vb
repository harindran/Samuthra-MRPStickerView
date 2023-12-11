Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace MRPStickerView
    <FormAttribute("10093", "CFLForm.b1f")>
    Friend Class CFLForm
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()
            Try
                Dim oMatrix As SAPbouiCOM.Matrix
                If Link_Value <> "-1" Then
                    objform = objaddon.objapplication.Forms.GetForm("10093", 0)
                    objform.Items.Item("5").Visible = False
                    oMatrix = objform.Items.Item("7").Specific
                    Link_Value = "-1"
                End If

            Catch ex As Exception

            End Try
        End Sub
       
    End Class
End Namespace
