Imports CrystalDecisions.CrystalReports.Engine
Public Class frm_report
    Dim i As Integer

    Private Sub frm_report_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        CrystalReportViewer1.Refresh()
    End Sub

    Public Sub New(ByRef doc As ReportDocument, ByRef ds As DataSet)
        InitializeComponent()

        doc.SetDataSource(ds)

        CrystalReportViewer1.ReportSource = doc
        CrystalReportViewer1.DisplayGroupTree = False

    End Sub

    Public Sub New(ByRef doc As ReportDocument, ByRef ds As DataSet, ByVal parameter() As Object, ByVal displayname() As Object)
        InitializeComponent()

        doc.SetDataSource(ds)

        For i = 0 To parameter.Length - 1
            doc.SetParameterValue(parameter(i), displayname(i))
        Next

        CrystalReportViewer1.ReportSource = doc
        CrystalReportViewer1.DisplayGroupTree = False
    End Sub
End Class