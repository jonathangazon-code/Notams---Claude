Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces
<DataContract>
<KnownType(GetType(Length))>
Public Class Length
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
    ''' This is an alternate constructor for the Length class.
    ''' It uses the value to represent the length as well as a measurement system.
    ''' Sets the value from feet to meters.
    ''' </summary>
    ''' <param name="value">A double variable meant to represent the length.</param>
    ''' <param name="measurementSystem">This is the measurement system that is being used.(Metric or UsCustomary)</param>
    Public Sub New(value As Double, measurementSystem As IMeasurmentSystem)
        measurementSystem.SetValue(value, 25.4, us, si)
    End Sub

    Public Function GetValue(measurementSystem As IMeasurmentSystem) As Double Implements IDimensionalProperty.GetValue
        Return measurementSystem.GetValue(us, si)
    End Function
End Class
