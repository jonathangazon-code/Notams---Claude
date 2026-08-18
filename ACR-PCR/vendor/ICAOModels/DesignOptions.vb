Imports System.ComponentModel
Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces
'Imports Telerik.Windows.Controls

<DataContract>
<KnownType(GetType(Thickness))>
<KnownType(GetType(UsCustomary))>
<KnownType(GetType(Metric))>
Public Class DesignOptions
    Implements INotifyPropertyChanged
    Implements IDesignOptions

    Private _MeasurementSystem As IMeasurmentSystem
    <DataMember>
    Public Property MeasurementSystem As IMeasurmentSystem Implements IDesignOptions.MeasurementSystem
        Get
            Return _MeasurementSystem
        End Get
        Set(value As IMeasurmentSystem)
            _MeasurementSystem = value
            OnPropertyChanged(NameOf(MeasurementSystem))
        End Set
    End Property

    Private _CdfTolerance As Double
    <DataMember>
    Public Property CdfTolerance As Double Implements IDesignOptions.CdfTolerance
        Get
            Return _CdfTolerance
        End Get
        Set(value As Double)
            _CdfTolerance = value
            OnPropertyChanged(NameOf(CdfTolerance))
        End Set
    End Property

    Private _LifeTolerance As Double
    <DataMember>
    Public Property LifeTolerance As Double Implements IDesignOptions.LifeTolerance
        Get
            Return _LifeTolerance
        End Get
        Set(value As Double)
            _LifeTolerance = value
            OnPropertyChanged(NameOf(LifeTolerance))
        End Set
    End Property

    Private _CalculateHmaCdf As Boolean
    <DataMember>
    Public Property CalculateHmaCdf As Boolean Implements IDesignOptions.CalculateHmaCdf
        Get
            Return _CalculateHmaCdf
        End Get
        Set(value As Boolean)
            _CalculateHmaCdf = value
            OnPropertyChanged(NameOf(CalculateHmaCdf))
        End Set
    End Property

    Private _AlternateSubgrade As Boolean
    <DataMember>
    Public Property AlternateSubgrade As Boolean Implements IDesignOptions.AlternateSubgrade
        Get
            Return _AlternateSubgrade
        End Get
        Set(value As Boolean)
            _AlternateSubgrade = value
            OnPropertyChanged(NameOf(AlternateSubgrade))
        End Set
    End Property

    Private _AutomaticFlexibleBaseDesign As Boolean
    <DataMember>
    Public Property AutomaticFlexibleBaseDesign As Boolean Implements IDesignOptions.AutomaticFlexibleBaseDesign
        Get
            Return _AutomaticFlexibleBaseDesign
        End Get
        Set(value As Boolean)
            _AutomaticFlexibleBaseDesign = value
            OnPropertyChanged(NameOf(AutomaticFlexibleBaseDesign))
        End Set
    End Property

    Private _SectionParameterN As Double
    <DataMember>
    Public Property SectionParameterN As Double Implements IDesignOptions.SectionParameterN
        Get
            Return _SectionParameterN
        End Get
        Set(value As Double)
            _SectionParameterN = value
            OnPropertyChanged(NameOf(SectionParameterN))
        End Set
    End Property

    Private _AllowPartiallyBonded As Boolean
    <DataMember>
    Public Property AllowPartiallyBonded As Boolean Implements IDesignOptions.AllowPartiallyBonded
        Get
            Return _AllowPartiallyBonded
        End Get
        Set(value As Boolean)
            _AllowPartiallyBonded = value
            OnPropertyChanged(NameOf(AllowPartiallyBonded))
        End Set
    End Property
    Private _PCAConversion As Boolean
    <DataMember>
    Public Property PCAConversion As Boolean Implements IDesignOptions.PCAConversion
        Get
            Return _PCAConversion
        End Get
        Set(value As Boolean)
            _PCAConversion = value
            OnPropertyChanged(NameOf(PCAConversion))
        End Set
    End Property

    Private _Outfile As Boolean
    <DataMember>
    Public Property Outfile As Boolean Implements IDesignOptions.Outfile
        Get
            Return _Outfile
        End Get
        Set(value As Boolean)
            _Outfile = value
            OnPropertyChanged(NameOf(Outfile))
        End Set
    End Property

    Private _ThickPccOverlay As Boolean
    <DataMember>
    Public Property ThickPccOverlay As Boolean Implements IDesignOptions.ThickPccOverlay
        Get
            Return _ThickPccOverlay
        End Get
        Set(value As Boolean)
            _ThickPccOverlay = value
            OnPropertyChanged(NameOf(ThickPccOverlay))
        End Set
    End Property

    Private _CrackPropogation As IDimensionalProperty
    <DataMember>
    Public Property CrackPropogation As IDimensionalProperty Implements IDesignOptions.CrackPropogation
        Get
            Return _CrackPropogation
        End Get
        Set(value As IDimensionalProperty)
            _CrackPropogation = value
            OnPropertyChanged(NameOf(CrackPropogation))
        End Set
    End Property

    Private _ComputeCompaction As Boolean
    <DataMember>
    Public Property ComputeCompaction As Boolean Implements IDesignOptions.ComputeCompaction
        Get
            Return _ComputeCompaction
        End Get
        Set(value As Boolean)
            _ComputeCompaction = value
            OnPropertyChanged(NameOf(ComputeCompaction))
        End Set
    End Property

    Private _ACROptions As Boolean
    <DataMember>
    Public Property ACROptions As Boolean Implements IDesignOptions.ACROptions
        Get
            Return _ACROptions
        End Get
        Set(value As Boolean)
            _ACROptions = value
            OnPropertyChanged(NameOf(ACROptions))
        End Set
    End Property

    Private Property CrackPropogationValidation As IValidationRange
    Private Property CdfToleranceValidation As IValidationRange
    Private Property LifeToleranceValidation As IValidationRange
    Private Property SectionParameterNValidation As IValidationRange


    ''' <summary>
    ''' This is the default constructor for the DesignOptions class.
    ''' </summary>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' This is one of the alternate constructors for the DesignOptions class.
    ''' This one is meant to initialize the values using a factory object.
    ''' </summary>
    ''' <param name="factory">This is the factory object that is passed in to get values from.</param>
    Public Sub New(factory As IICAOFactory)
        CdfTolerance = 0.005
        LifeTolerance = 0.4
        SectionParameterN = 16
        AutomaticFlexibleBaseDesign = True
        Outfile = False
        ComputeCompaction = True
        ThickPccOverlay = True
    End Sub

    ''' <summary>
    ''' This is another contructor for the DesignOptions class.
    ''' It uses an already existing DesignOptions object to get its values.
    ''' </summary>
    ''' <param name="factory">This is the factory object that is being used.</param>
    ''' <param name="existing">This is the already existing DesignOptions object.</param>
    Public Sub New(factory As IICAOFactory, existing As IDesignOptions)
        MeasurementSystem = existing.MeasurementSystem
        CdfTolerance = existing.CdfTolerance
        LifeTolerance = existing.LifeTolerance
        CalculateHmaCdf = existing.CalculateHmaCdf
        AlternateSubgrade = existing.AlternateSubgrade
        AutomaticFlexibleBaseDesign = existing.AutomaticFlexibleBaseDesign
        SectionParameterN = existing.SectionParameterN
        AllowPartiallyBonded = existing.AllowPartiallyBonded
        PCAConversion = existing.PCAConversion
        Outfile = existing.Outfile
        ACROptions = existing.ACROptions
        CrackPropogation = factory.CreateThickness(existing.CrackPropogation.GetValue(MeasurementSystem), MeasurementSystem)
        ComputeCompaction = existing.ComputeCompaction
        ThickPccOverlay = existing.ThickPccOverlay
    End Sub

    Private Sub Clone()

    End Sub

    ' INotifyPropertyChanged in place of Telerik ViewModelBase
    Private ReadOnly _argsCache As Dictionary(Of String, PropertyChangedEventArgs) = New Dictionary(Of String, PropertyChangedEventArgs)()

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub OnPropertyChanged(ByVal propertyName As String)
        If _argsCache IsNot Nothing Then
            If Not _argsCache.ContainsKey(propertyName) Then _argsCache(propertyName) = New PropertyChangedEventArgs(propertyName)
            NotifyChange(_argsCache(propertyName))
        End If
    End Sub

    Private Sub NotifyChange(ByVal e As PropertyChangedEventArgs)
        RaiseEvent PropertyChanged(Me, e)
    End Sub

End Class
