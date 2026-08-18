Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces
<DataContract>
Public Class Modulus
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
    ''' This is the constructor  for the modulus class.
    ''' </summary>
    ''' <param name="value">This is the modulus value stored as a double.</param>
    ''' <param name="measurementSystem">This is the measurement system that is used for the modulus.</param>
    Public Sub New(value As Double, measurementSystem As IMeasurmentSystem)
        measurementSystem.SetValue(value, 0.00689476, us, si)
    End Sub

    Public Function GetValue(measurementSystem As IMeasurmentSystem) As Double Implements IDimensionalProperty.GetValue
        Return measurementSystem.GetValue(us, si)
    End Function
End Class
