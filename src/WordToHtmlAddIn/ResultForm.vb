Imports System.Windows.Forms

Public Class ResultForm

    Public Overloads Function ShowDialog(resultHtmlText As String) As DialogResult
        Me.HtmlResultTextBox.Text = resultHtmlText
        Return Me.ShowDialog
    End Function

End Class