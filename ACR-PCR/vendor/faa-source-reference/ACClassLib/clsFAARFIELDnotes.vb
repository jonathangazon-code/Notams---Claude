Public Class clsFAARFIELDnotes


    Public WorkingDirectory As String
    Public Thickness() As Single


    Public Sub setThickness()
        ReDim Thickness(4)
        Thickness(1) = 4.2
        Thickness(2) = 4.3
        Thickness(3) = 4.4
        Thickness(4) = 4.5

    End Sub



    Public Property Title() As String
        Get
            Return WorkingDirectory
        End Get
        Set(ByVal Value As String)
            WorkingDirectory = Value
        End Set
    End Property



End Class
