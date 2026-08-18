Namespace Interfaces
    Public Interface IDimensionalProperty
        ReadOnly Property UsCustomary As Double
        ReadOnly Property Metric As Double
        Function GetValue(measurementSystem As IMeasurmentSystem) As Double
    End Interface
End Namespace