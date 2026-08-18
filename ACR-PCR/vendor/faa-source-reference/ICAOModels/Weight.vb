Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces
<DataContract>
<KnownType(GetType(Weight))>
Public Class Weight
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

    Private Sub New()

    End Sub

    Public Sub New(value As Double, measurementSystem As IMeasurmentSystem)
        measurementSystem.SetValue(value, 0.453592, us, si)  'lbs to kg
    End Sub

    Public Function GetValue(measurementSystem As IMeasurmentSystem) As Double Implements IDimensionalProperty.GetValue
        Return measurementSystem.GetValue(us, si)
    End Function
End Class
