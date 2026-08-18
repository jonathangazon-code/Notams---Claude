Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Text
Imports ICAOModels
Imports ICAOModels.Interfaces

Public Class ICAOFactory
    Implements IICAOFactory

    Public Function CreateUsCustomary() As IMeasurmentSystem Implements IICAOFactory.CreateUsCustomary
        Return New UsCustomary()
    End Function

    Public Function CreateMetric() As IMeasurmentSystem Implements IICAOFactory.CreateMetric
        Return New Metric()
    End Function


    Public Function CreateThickness(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty Implements IICAOFactory.CreateThickness
        Return New Thickness(value, measurementSystem)
    End Function

    Public Function CreateArea(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty Implements IICAOFactory.CreateArea
        Return New Area(value, measurementSystem)
    End Function

    Public Function CreatePressure(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty Implements IICAOFactory.CreatePressure
        Return New Pressure(value, measurementSystem)
    End Function

    Public Function CreateModulus(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty Implements IICAOFactory.CreateModulus
        Return New Modulus(value, measurementSystem)
    End Function

    Public Function CreateLength(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty Implements IICAOFactory.CreateLength
        Return New Length(value, measurementSystem)
    End Function

    Public Function CreateWeight(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty Implements IICAOFactory.CreateWeight
        Return New Weight(value, measurementSystem)
    End Function

    Public Function CreateCoordinates(x As IDimensionalProperty, y As IDimensionalProperty) As ICoordinates Implements IICAOFactory.CreateCoordinates
        Return New LengthCoordinates(x, y)
    End Function

End Class
