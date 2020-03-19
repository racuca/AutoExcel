Imports AutoExcel.Core

Public Class Form1

    Private Sub btnHelloWorld_Click(sender As Object, e As EventArgs) Handles btnHelloWorld.Click
        Dim path As String

        Dim excelbinder As New ExcelBinder()
        excelbinder.NewDocument()
        excelbinder.Visible = True
        excelbinder.SetValue("A1", "Hello world")
        path = My.Computer.FileSystem.CombinePath(Environment.CurrentDirectory, "testhello1.xlsx")
        excelbinder.SaveDocument(path)
        excelbinder.CloseDocument()
        excelbinder.QuitDocument()
    End Sub
End Class
