Imports ICAOModels.Interfaces



Public Class Metric
    Implements IMeasurmentSystem

    Public Property Name As String Implements IMeasurmentSystem.Name

    ''' <summary>
    ''' This is the constructor for the Metric class.
    ''' It initializes the Name variable to "Metric".
    ''' </summary>
    Sub New()
        Name = "Metric"
    End Sub

    Public Function GetValue(us As String, si As String) As String Implements IMeasurmentSystem.GetValue
        Return si
    End Function

    Public Function GetValue(us As Double, si As Double) As Double Implements IMeasurmentSystem.GetValue
        Return si
    End Function

    Public Sub SetValue(value As Double, convert As Double, ByRef us As Double, ByRef si As Double) Implements IMeasurmentSystem.SetValue
        si = value
        us = si / convert
    End Sub

    Public Sub SetTemperatureValue(value As Double, ByRef us As Double, ByRef si As Double) Implements IMeasurmentSystem.SetTemperatureValue
        si = value
        us = (value * 1.8) + 32
    End Sub
End Class
