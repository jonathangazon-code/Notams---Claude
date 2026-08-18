Imports System.ComponentModel
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization
Imports System.Windows.Forms
Imports ICAOModels.Interfaces



<DataContract>
<KnownType(GetType(AirplaneInfo))>
<KnownType(GetType(Length))>
<KnownType(GetType(Weight))>
<KnownType(GetType(LengthCoordinates))>
<KnownType(GetType(UsCustomary))>
<KnownType(GetType(Metric))>
Public Class AirplaneInfo
    Implements IAirplaneInfo
    Implements INotifyPropertyChanged

    <DataMember>
    Public Property Manufacturer As String Implements IAirplaneInfo.Manufacturer
    <DataMember>
    Public Property Name As String Implements IAirplaneInfo.Name
    <DataMember>
    Public Property _GrossWeight As Weight
    Private Property _MgPercent As Single

    <DataMember>
    Public Property MgPercent As Single Implements IAirplaneInfo.MgPercent
        Get
            Return _MgPercent
        End Get
        Set(value As Single)
            _MgPercent = value
            OnPropertyChanged(NameOf(MgPercent))
        End Set
    End Property

    Private Property _MgPercentPCN As Single
    <DataMember>
    Public Property MgPercentPCN As Single Implements IAirplaneInfo.MgPercentPCN
        Get
            Return _MgPercentPCN
        End Get
        Set(value As Single)
            _MgPercentPCN = value
            OnPropertyChanged(NameOf(MgPercentPCN))
        End Set
    End Property

    Private Property _RunMgPercent As Single
    <DataMember>
    Public Property RunMgPercent As Single Implements IAirplaneInfo.RunMgPercent
        Get
            Return _RunMgPercent
        End Get
        Set(value As Single)
            _RunMgPercent = value
            OnPropertyChanged(NameOf(RunMgPercent))
        End Set
    End Property

    <DataMember>
    Public Property Gear As String Implements IAirplaneInfo.Gear
    <DataMember>
    Public Property NumberGear As Integer Implements IAirplaneInfo.NumberGear
    <DataMember(Order:=1)>
    Public Property Tt As Thickness Implements IAirplaneInfo.Tt
    <DataMember>
    Public Property D2 As Thickness Implements IAirplaneInfo.D2
    <DataMember>
    Public Property d1 As Thickness Implements IAirplaneInfo.d1
    <DataMember>
    Public Property dmax As Thickness Implements IAirplaneInfo.dmax
    <DataMember>
    Public Property d3 As Thickness Implements IAirplaneInfo.d3
    <DataMember>
    Public Property D4 As Thickness Implements IAirplaneInfo.D4
    <DataMember>
    Public Property d5 As Thickness Implements IAirplaneInfo.d5
    <DataMember>
    Public Property D6 As Thickness Implements IAirplaneInfo.D6
    <DataMember>
    Public Property aa As Thickness Implements IAirplaneInfo.aa
    <DataMember>
    Public Property Ts As Single Implements IAirplaneInfo.Ts
    <DataMember>
    Public Property Tg As Single Implements IAirplaneInfo.Tg
    <DataMember>
    Public Property Xcoorodinates() As Single Implements IAirplaneInfo.Xcoordinates
    <DataMember>
    Public Property Ycoordinates() As Single Implements IAirplaneInfo.Ycoordinates
    <DataMember(Order:=2)>
    Public Property B As Thickness Implements IAirplaneInfo.B
    <DataMember>
    Public Property TBT As Thickness Implements IAirplaneInfo.TBT
    <DataMember>
    Public Property NumberWheels As Integer Implements IAirplaneInfo.NumberWheels
    <DataMember>
    Public Property WheelCoordinates As List(Of ICoordinates) Implements IAirplaneInfo.WheelCoordinates
    <DataMember>
    Public Property EvaluationPoints As List(Of ICoordinates) Implements IAirplaneInfo.EvaluationPoints
    <DataMember>
    Public Property GX As Single Implements IAirplaneInfo.GX
    Public Property GY As Single Implements IAirplaneInfo.GY
    <DataMember>
    Public Property NumberTireTracks As Integer Implements IAirplaneInfo.NumberTireTrack
    <DataMember>
    Public Property TireTrackX As List(Of IDimensionalProperty) Implements IAirplaneInfo.TireTrackX
    <DataMember>
    Public Property _NumberDepartures As Integer
    <DataMember>
    Public Property gNewAnnualdepartue As Integer Implements IAirplaneInfo.gNewAnnualdepartue
    Private Property _TotalDepartures As Int64
    <DataMember>
    Public Property GearOrientation As Integer Implements IAirplaneInfo.GearOrientation

    <DataMember>
    Public Property TotalDepartures As Int64 Implements IAirplaneInfo.TotalDepartures
        Get
            Return _TotalDepartures
        End Get
        Set(value As Int64)
            _TotalDepartures = value
            OnPropertyChanged(NameOf(TotalDepartures))
        End Set
    End Property

    Private Property _cp As Pressure
    <DataMember>
    Public Property Cp As Pressure Implements IAirplaneInfo.Cp
        Get
            Return _cp
        End Get
        Set(value As Pressure)
            _cp = value
            OnPropertyChanged(NameOf(Cp))
        End Set
    End Property
    <DataMember>
    Public Property _AnnualGrowth As Single
    <DataMember>
    Public Property IsBelly As Boolean Implements IAirplaneInfo.IsBelly
    <DataMember>
    Public Property Deprecated As Boolean Implements IAirplaneInfo.Deprecated
    Private Property _Cdf As Single
    <DataMember>
    Public Property Cdf As Single Implements IAirplaneInfo.Cdf
        Get
            Return _Cdf
        End Get
        Set(value As Single)
            _Cdf = value
            OnPropertyChanged(NameOf(Cdf))
        End Set
    End Property
    Private Property _CdfAircraftMax As Single
    <DataMember>
    Public Property CdfAircraftMax As Single Implements IAirplaneInfo.CdfAircraftMax
        Get
            Return _CdfAircraftMax
        End Get
        Set(value As Single)
            _CdfAircraftMax = value
            OnPropertyChanged(NameOf(CdfAircraftMax))
        End Set
    End Property
    <DataMember>
    Public Property CdfAc As Single Implements IAirplaneInfo.CdfAc
    <DataMember>
    Public Property CdfAircraftMaxAc As Single Implements IAirplaneInfo.CdfAircraftMaxAc
    <DataMember>
    Public Property CdfHma As Single Implements IAirplaneInfo.CdfHma
    <DataMember>
    Public Property CdfAircraftMaxHma As Single Implements IAirplaneInfo.CdfAircraftMaxHma
    <DataMember>
    Public Property Cdf401 As Single Implements IAirplaneInfo.Cdf401
    <DataMember>
    Public Property CdfAircraftMax401 As Single Implements IAirplaneInfo.CdfAircraftMax401
    <DataMember>
    Public Property CdfSub As Single Implements IAirplaneInfo.CdfSub
    <DataMember>
    Public Property CdfAircraftMaxSub As Single Implements IAirplaneInfo.CdfAircraftMaxSub
    Private Property _CtoP As Single
    <DataMember>
    Public Property CtoP As Single Implements IAirplaneInfo.CtoP
        Get
            Return _CtoP
        End Get
        Set(value As Single)
            _CtoP = value
            OnPropertyChanged(NameOf(CtoP))
        End Set
    End Property
    <DataMember>
    Public Property Dtemp As Single Implements IAirplaneInfo.Dtemp
    <DataMember>
    Public Property CtoPAc As Single Implements IAirplaneInfo.CtoPAc
    <DataMember>
    Public Property CtoPHma As Single Implements IAirplaneInfo.CtoPHma
    <DataMember>
    Public Property CtoP401 As Single Implements IAirplaneInfo.CtoP401
    <DataMember>
    Public Property CtoPSub As Single Implements IAirplaneInfo.CtoPSub
    <DataMember>
    Public Property TireWidth As Thickness Implements IAirplaneInfo.TireWidth
    <DataMember>
    Public Property TireLength As Thickness Implements IAirplaneInfo.TireLength
    <DataMember>
    Public Property TireArea As Area Implements IAirplaneInfo.TireArea
    <DataMember>
    Public Property TirePressureF As Thickness Implements IAirplaneInfo.TirePressureF
    <DataMember>
    Public Property AircraftNumber As Integer Implements IAirplaneInfo.AircraftNumber
    <DataMember>
    Public Property Coverage As Double Implements IAirplaneInfo.Coverage
    Public Property CriticalAnnualDeparture As Double Implements IAirplaneInfo.CriticalAnnualDeparture
    Private Property _ACRThick As Thickness

    <DataMember>
    Public Property ACRThick As Thickness Implements IAirplaneInfo.ACRThick
        Get
            Return _ACRThick
        End Get
        Set(value As Thickness)
            _ACRThick = value
            OnPropertyChanged(NameOf(ACRThick))
        End Set
    End Property
    Private Property _ACRThick1 As Thickness
    <DataMember>
    Public Property ACRThick1 As Thickness Implements IAirplaneInfo.ACRThick1
        Get
            Return _ACRThick1
        End Get
        Set(value As Thickness)
            _ACRThick1 = value
            OnPropertyChanged(NameOf(ACRThick1))
        End Set
    End Property
    Private Property _ACRThick2 As Thickness
    <DataMember>
    Public Property ACRThick2 As Thickness Implements IAirplaneInfo.ACRThick2
        Get
            Return _ACRThick2
        End Get
        Set(value As Thickness)
            _ACRThick2 = value
            OnPropertyChanged(NameOf(ACRThick2))
        End Set
    End Property
    Private Property _ACRThick3 As Thickness
    <DataMember>
    Public Property ACRThick3 As Thickness Implements IAirplaneInfo.ACRThick3
        Get
            Return _ACRThick3
        End Get
        Set(value As Thickness)
            _ACRThick3 = value
            OnPropertyChanged(NameOf(ACRThick3))
        End Set
    End Property
    Private Property _ACRThick4 As Thickness
    <DataMember>
    Public Property ACRThick4 As Thickness Implements IAirplaneInfo.ACRThick4
        Get
            Return _ACRThick4
        End Get
        Set(value As Thickness)
            _ACRThick4 = value
            OnPropertyChanged(NameOf(ACRThick4))
        End Set
    End Property
    <DataMember>
    Public Property PCRThick As Double Implements IAirplaneInfo.PCRThick
    <DataMember>
    Public Property ACRThickMGW As Double Implements IAirplaneInfo.ACRThickMGW
    Private Property _ACRB As Double
    <DataMember>
    Public Property ACRB As Double Implements IAirplaneInfo.ACRB
        Get
            Return _ACRB
        End Get
        Set(value As Double)
            _ACRB = value
            OnPropertyChanged(NameOf(ACRB))
        End Set
    End Property
    <DataMember>
    Public Property gNewGL As Weight Implements IAirplaneInfo.gNewGL
    Dim _factory As ICAOFactory
    <DataMember>
    Public Property Tv As Single Implements IAirplaneInfo.Tv
    Public Property DefaultGrossWeight As Weight Implements IAirplaneInfo.DefaultGrossWeight
    Public Property DefaultCp As Pressure Implements IAirplaneInfo.DefaultCp
    <DataMember>
    Public Property CDFGraphData As List(Of Single) Implements IAirplaneInfo.CDFGraphData
    <DataMember>
    Public Property PCRNumber As Single Implements IAirplaneInfo.PCRNumber
    <DataMember>
    Public Property ACRCoverage As Integer Implements IAirplaneInfo.ACRCoverage
    Private Property _ACRB1 As Single

    <DataMember>
    Public Property ACRB1 As Single Implements IAirplaneInfo.ACRB1
        Get
            Return _ACRB1
        End Get
        Set(value As Single)
            _ACRB1 = value
            OnPropertyChanged(NameOf(ACRB1))
        End Set
    End Property
    Private Property _ACRB2 As Single
    <DataMember>
    Public Property ACRB2 As Single Implements IAirplaneInfo.ACRB2
        Get
            Return _ACRB2
        End Get
        Set(value As Single)
            _ACRB2 = value
            OnPropertyChanged(NameOf(ACRB2))
        End Set
    End Property
    Private Property _ACRB3 As Single
    <DataMember>
    Public Property ACRB3 As Single Implements IAirplaneInfo.ACRB3
        Get
            Return _ACRB3
        End Get
        Set(value As Single)
            _ACRB3 = value
            OnPropertyChanged(NameOf(ACRB3))
        End Set
    End Property
    Private Property _ACRB4 As Single
    <DataMember>
    Public Property ACRB4 As Single Implements IAirplaneInfo.ACRB4
        Get
            Return _ACRB4
        End Get
        Set(value As Single)
            _ACRB4 = value
            OnPropertyChanged(NameOf(ACRB4))
        End Set
    End Property

    ''' <summary>
    ''' This is the default constructor for the AirplaneInfo class.
    ''' </summary>
    Public Sub New()

    End Sub

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    ' This method is called by the Set accessor of each property.
    ' The CallerMemberName attribute that is applied to the optional propertyName
    ' parameter causes the property name of the caller to be substituted as an argument.
    Private Sub OnPropertyChanged(<CallerMemberName()> Optional ByVal info As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(info))
    End Sub
    ''' <summary>
    ''' This is the alternate constructor for the AirplaneInfo class.
    ''' It uses an already existing instance AirplaneInfo object and factory.
    ''' </summary>
    ''' <param name="existing">This is the already existing airplane info.</param>
    ''' <param name="factory">This the factory object that is passed in.</param>
    Public Sub New(existing As IAirplaneInfo, factory As ICAOFactory)
        Dim coordinateSystem = factory.CreateUsCustomary()
        _factory = factory
        Manufacturer = existing.Manufacturer
        Name = existing.Name
        Cp = existing.Cp
        DefaultCp = existing.Cp
        _GrossWeight = existing.GrossWeight
        DefaultGrossWeight = existing.GrossWeight
        MgPercent = existing.MgPercent
        MgPercentPCN = existing.MgPercentPCN
        RunMgPercent = existing.MgPercent
        GearOrientation = existing.GearOrientation
        CDFGraphData = New List(Of Single)
        For i = 1 To 42
            CDFGraphData.Add(0)
        Next

        Gear = existing.Gear
        NumberGear = existing.NumberGear
        Tt = factory.CreateThickness(existing.Tt.UsCustomary, coordinateSystem)
        Ts = existing.Ts
        Tg = existing.Tg
        B = factory.CreateThickness(existing.B.UsCustomary, coordinateSystem)
        NumberWheels = existing.NumberWheels
        WheelCoordinates = New List(Of ICoordinates)
        For Each wheelCoordinate In existing.WheelCoordinates
            WheelCoordinates.Add(factory.CreateCoordinates(factory.CreateLength(wheelCoordinate.X.UsCustomary, coordinateSystem), factory.CreateLength(wheelCoordinate.Y.UsCustomary, coordinateSystem)))
        Next

        EvaluationPoints = New List(Of ICoordinates)
        For Each evaluationPoint In existing.EvaluationPoints
            EvaluationPoints.Add(factory.CreateCoordinates(factory.CreateLength(evaluationPoint.X.UsCustomary, coordinateSystem), factory.CreateLength(evaluationPoint.Y.UsCustomary, coordinateSystem)))
        Next

        NumberTireTracks = existing.NumberTireTrack
        TireTrackX = New List(Of IDimensionalProperty)
        If Not existing.TireTrackX Is Nothing Then

            For Each tireTrack In existing.TireTrackX
                TireTrackX.Add(factory.CreateLength(tireTrack.UsCustomary, coordinateSystem))
            Next
        End If

        AnnualGrowth = existing.AnnualGrowth
        NumberDepartures = existing.NumberDepartures
        TotalDepartures = NumberDepartures * 20
        Dim Contactarea As Single
        Contactarea = (Convert.ToSingle(_GrossWeight.UsCustomary) * MgPercent / Convert.ToSingle(Cp.UsCustomary)) / NumberWheels / 2

        IsBelly = existing.IsBelly
        Deprecated = existing.Deprecated
        TireArea = factory.CreateArea(Contactarea, New UsCustomary)
        TireWidth = factory.CreateThickness(CSng(2 * System.Math.Sqrt(Contactarea / (1.6 * 3.14159))), New UsCustomary)
        TireLength = factory.CreateThickness(TireWidth.UsCustomary * 1.6, coordinateSystem)

        Coverage = existing.Coverage
        Dim aircraftsbelly As New List(Of IAirplaneInfo)

    End Sub
    Public Property NumberDepartures As Integer Implements IAirplaneInfo.NumberDepartures
        Get
            Return _NumberDepartures
        End Get
        Set(value As Integer)
            _NumberDepartures = value
            If value > 100000 Then
                MessageBox.Show("Maximum allowable number for Annual Departure is 100,000. ")
                _NumberDepartures = 1200
            ElseIf value < 0 Then
                MessageBox.Show("Minimum allowable value for Annual Departures is 0")
                _NumberDepartures = 0
                TotalDepartures = 0
            Else
                _NumberDepartures = value
            End If
            OnPropertyChanged(NameOf(NumberDepartures))
        End Set
    End Property
    Public Property AnnualGrowth As Single Implements IAirplaneInfo.AnnualGrowth
        Get
            Return _AnnualGrowth
        End Get
        Set(value As Single)
            _AnnualGrowth = value
            If _AnnualGrowth < -10 Or _AnnualGrowth > 10 Then
                MessageBox.Show("Annual Growth allowed range is -10 to 10.")
                _AnnualGrowth = 0
            Else
            End If
            OnPropertyChanged(NameOf(AnnualGrowth))
        End Set
    End Property

    Public Property GrossWeight As Weight Implements IAirplaneInfo.GrossWeight
        Get
            Return _GrossWeight
        End Get
        Set(value As Weight)


            Dim NWMax As Single
            Dim NWMin As Single
            Dim NV As Single
            Dim NWMaxMetric As Single
            Dim NWMinMetric As Single
            Dim NVMetric As Single
            Dim Contactarea As Single
            Dim factory As IICAOFactory
            If Not _GrossWeight Is Nothing Then


                Contactarea = (Convert.ToSingle(_GrossWeight.UsCustomary) / Convert.ToSingle(Cp.UsCustomary))
                NV = System.Convert.ToSingle(value.UsCustomary)
                NVMetric = System.Convert.ToSingle(value.Metric)
                If Not DefaultGrossWeight Is Nothing Then
                    NWMax = System.Convert.ToSingle(DefaultGrossWeight.UsCustomary) * 1.25
                    NWMin = System.Convert.ToSingle(DefaultGrossWeight.UsCustomary) * 0.6
                    NWMaxMetric = System.Convert.ToSingle(DefaultGrossWeight.Metric) * 1.25
                    NWMinMetric = System.Convert.ToSingle(DefaultGrossWeight.Metric) * 0.6
                Else
                    DefaultGrossWeight = _GrossWeight
                    NWMax = System.Convert.ToSingle(DefaultGrossWeight.UsCustomary) * 1.25
                    NWMin = System.Convert.ToSingle(DefaultGrossWeight.UsCustomary) * 0.6
                    NWMaxMetric = System.Convert.ToSingle(DefaultGrossWeight.Metric) * 1.25
                    NWMinMetric = System.Convert.ToSingle(DefaultGrossWeight.Metric) * 0.6
                End If


                If NV > NWMax Then
                    MessageBox.Show("Allowed Gross Taxi weight for this airplane is " + NWMin.ToString + " lb" + " (" + NWMinMetric.ToString + " Kg" + ")" + " to " + NWMax.ToString + " lb" + " (" + NWMaxMetric.ToString + " Kg" + ")" + ".")
                    value = _GrossWeight
                ElseIf NV < NWMin Then
                    MessageBox.Show("Allowed Gross Taxi weight for this airplane is " + NWMin.ToString + " lb" + " (" + NWMinMetric.ToString + " Kg" + ")" + " to " + NWMax.ToString + " lb" + " (" + NWMaxMetric.ToString + " Kg" + ")" + ".")
                    value = _GrossWeight

                Else

                    _GrossWeight = value

                End If
            Else
                _GrossWeight = value
            End If

            Dim K As Double
            If Contactarea <> 0 Then
                K = _GrossWeight.UsCustomary / Contactarea
            End If
            If Not _factory Is Nothing Then
                Cp = _factory.CreatePressure(_GrossWeight.UsCustomary * MgPercent / NumberWheels / (DefaultGrossWeight.UsCustomary * MgPercent / NumberWheels / DefaultCp.UsCustomary), New UsCustomary)

            End If

            OnPropertyChanged(NameOf(Cp))
            OnPropertyChanged(NameOf(GrossWeight))

            OnPropertyChanged("SelectedAirplane")
            Call Updatecursorposition()
            CtoP = 0
            OnPropertyChanged("CtoP")
        End Set
    End Property
    Public Overrides Function ToString() As String
        Return Name
    End Function

    Sub Updatecursorposition()
        Dim X As Double
        Dim Y As Double
        X = Cursor.Position.X
        Y = Cursor.Position.Y
        Cursor.Position = Nothing
        Cursor.Position = New Point(X, Y)


    End Sub

End Class
