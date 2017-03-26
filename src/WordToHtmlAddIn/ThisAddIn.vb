Public Class ThisAddIn

    Private Shared myAddInReference As ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        myAddInReference = Me
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        myAddInReference = Nothing
    End Sub

    Public Shared ReadOnly Property AddInReference As ThisAddIn
        Get
            Return myAddInReference
        End Get
    End Property

End Class
