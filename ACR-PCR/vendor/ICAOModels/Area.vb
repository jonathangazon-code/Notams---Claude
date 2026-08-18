Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces

<DataContract>
Public Class Area
    Implements IDimensionalProperty
    Implements INotifyPropertyChanged

    <DataMember>
    Dim us As Double
    <DataMember>
    Dim si As Double

    Dim _setMeasurementSystem As IMeasurmentSystem
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    ' This method is called by the Set accessor of each property.
    ' The CallerMemberName attribute that is applied to the optional propertyName
    ' parameter causes the property name of the caller to be substituted as an argument.
    Private Sub OnPropertyChanged(<CallerMemberName()> Optional ByVal info As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(info))
    End Sub

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
    ''' This is the defualt constructor for the Area class.
    ''' </summary>
    Private Sub New()

    End Sub
    ''' <summary>
    ''' The constructor for Area class.
    ''' Sets the value while using a double and measurement system.
    ''' Specifically for inches squared to millimeter squared.
    ''' </summary>
    ''' <param name="value">A double variable meant to represent the area.</param>
    ''' <param name="measurementSystem">The measurement system used for the area.</param>
    Public Sub New(value As Double, measurementSystem As IMeasurmentSystem)
        measurementSystem.SetValue(value, 25.4 * 25.4, us, si)
    End Sub

    Public Function GetValue(measurementSystem As IMeasurmentSystem) As Double Implements IDimensionalProperty.GetValue
        Return measurementSystem.GetValue(us, si)
    End Function
End Class
