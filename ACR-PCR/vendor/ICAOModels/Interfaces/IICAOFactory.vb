Imports System.Collections.ObjectModel

Namespace Interfaces
    Public Interface IICAOFactory

#Region "Measurement Systems"
        ''' <summary>
        ''' Create US Customary MeasurementSystem object
        ''' </summary>
        ''' <returns></returns>
        Function CreateUsCustomary() As IMeasurmentSystem
        ''' <summary>
        ''' Create SI/Metric MeasurementSystem object
        ''' </summary>
        ''' <returns></returns>
        Function CreateMetric() As IMeasurmentSystem
#End Region

#Region "Dimensional Properties"
        ''' <summary>
        ''' Stores thickness data (inches/mm)
        ''' </summary>
        ''' <param name="value">Thickness value</param>
        ''' <param name="measurementSystem">US Cutomary or metric</param>
        ''' <returns></returns>
        Function CreateThickness(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty

        ''' <summary>
        ''' Stores area data (inches/mm)
        ''' </summary>
        ''' <param name="value">Area value</param>
        ''' <param name="measurementSystem">US Cutomary or metric</param>
        ''' <returns></returns>
        Function CreateArea(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty

        ''' <summary>
        ''' Stores subgrade reaction data (psi/in or kpa/m)
        ''' </summary>
        ''' <param name="value">Subgrade reaction value</param>
        ''' <param name="measurementSystem">US Cutomary or metric</param>
        ''' <returns></returns>
        Function CreatePressure(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty
        ''' <summary>
        ''' Stores modulues data (psi or kpa)
        ''' </summary>
        ''' <param name="value">Modulus value</param>
        ''' <param name="measurementSystem">US Cutomary or metric</param>
        ''' <returns></returns>
        Function CreateModulus(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty

        Function CreateLength(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty

        Function CreateCoordinates(x As IDimensionalProperty, y As IDimensionalProperty) As ICoordinates

        Function CreateWeight(value As Double, measurementSystem As IMeasurmentSystem) As IDimensionalProperty


#End Region
    End Interface
End Namespace