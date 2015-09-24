Public Class SaveGridEventArgs
    Inherits EventArgs

    Public Property filename As String

    Public Sub New(_filename As String)

        MyBase.New()

        Me.filename = _filename

    End Sub

End Class
