﻿Imports Oracle.DataAccess.Client
Imports Oracle.DataAccess.Types
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form3

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
        Dim objcr As New CrystalReport2
        Try
            objcr.SetDataSource(ds)
            Me.CrystalReportViewer1.ReportSource = objcr
            Me.CrystalReportViewer1.Refresh()
            objcr.PrintOptions.PaperOrientation = PaperOrientation.Landscape
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class