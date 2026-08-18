Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces


<DataContract>
Public Class Thickness
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

    Public Property Metric As Double Implements IDimensionalProperty.Metric
        Get
            Return si
        End Get
        Set
            If Not Value <> Nothing Then
                OnPropertyChanged("ProfileImage")
                OnPropertyChanged("TotalThicknessUpdate")
                OnPropertyChanged("Myvisibility")
                OnPropertyChanged("Myvisibility2")
                OnPropertyChanged("Layername1")
                OnPropertyChanged("Mymargin")
                OnPropertyChanged("Mymargin2")

            End If
        End Set
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    Private Sub New()

    End Sub
    ''' <summary>
    ''' This is the constructor for the thickness class.
    ''' Sets the values from inches to millimeters.
    ''' </summary>
    ''' <param name="value">This is the actual thickness value as a double.</param>
    ''' <param name="measurementSystem">This is the measurement system that the thickness is using.</param>
    Public Sub New(value As Double, measurementSystem As IMeasurmentSystem)
        measurementSystem.SetValue(value, 25.4, us, si)
    End Sub

    Public Function GetValue(measurementSystem As IMeasurmentSystem) As Double Implements IDimensionalProperty.GetValue
        Return measurementSystem.GetValue(us, si)
    End Function
End Class
