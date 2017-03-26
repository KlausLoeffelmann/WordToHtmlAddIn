Imports System.Windows.Forms

Public Class ProgressForm

    Private myMaxChars As Integer

    Public Property Progress As Integer
        Get
            Return Me.ConversionProgressBar.Value
        End Get
        Set(value As Integer)
            Me.ConversionProgressBar.Value = value
            Application.DoEvents()
        End Set
    End Property

    Public Property MaxChars As Integer
        Get
            Return Me.ConversionProgressBar.Maximum
        End Get
        Set(value As Integer)
            Me.ConversionProgressBar.Maximum = value
        End Set
    End Property

End Class