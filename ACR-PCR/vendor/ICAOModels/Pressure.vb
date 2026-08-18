Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces
<DataContract>
Public Class Pressure
    Implements IDimensionalProperty
    <DataMember>
    Dim us As Double
    <DataMember>
    Dim si As Double

    Dim _setMeasurementSystem As IMeasurmentSystem

    Public ReadOnly Property UsCustomary As Double Implements IDimensionalProperty.UsCustomary
        Get
            Return us
        End Get
    End Property

    Public ReadOnly Property Metric As Double Implements IDimensionalProperty.Metric
        Get
            Return si
        End Get
    End Property

    ''' <summary>
    ''' Constructor for the Pressure class.
    ''' Uses a value and measurement system as parameters.
    ''' </summary>
    ''' <param name="value"> The value to represent the pressure. </param>
    ''' <param name="measurementSystem"> The measurement system used. </param>
    Public Sub New(value As Double, measurementSystem As IMeasurmentSystem)
        measurementSystem.SetValue(value, 6.89476, us, si)
    End Sub

    Public Function GetValue(measurementSystem As IMeasurmentSystem) As Double Implements IDimensionalProperty.GetValue
        Return measurementSystem.GetValue(us, si)
    End Function
End Class