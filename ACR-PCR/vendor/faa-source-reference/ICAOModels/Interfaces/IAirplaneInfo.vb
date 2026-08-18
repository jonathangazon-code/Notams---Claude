Imports System.Drawing

Namespace Interfaces
    Public Interface IAirplaneInfo
        Property Manufacturer As String
        Property Name As String
        Property GrossWeight As Weight
        Property DefaultGrossWeight As Weight
        Property gNewGL As Weight
        Property MgPercent As Single
        Property MgPercentPCN As Single
        Property RunMgPercent As Single
        Property Cp As Pressure
        Property Gear As String
        Property NumberGear As Integer
        Property Tt As Thickness
        Property Ts As Single
        Property Tg As Single
        Property B As Thickness
        Property TBT As Thickness
        Property NumberWheels As Integer
        Property WheelCoordinates As List(Of ICoordinates)
        'Property NumberEvaluationPoints As Integer
        Property EvaluationPoints As List(Of ICoordinates)

        Property NumberTireTrack As Integer
        Property TireTrackX As List(Of IDimensionalProperty)

        Property Xcoordinates() As Single
        Property Ycoordinates() As Single
        Property NumberDepartures As Integer
        Property gNewAnnualdepartue As Integer
        Property TotalDepartures As Int64
        Property AnnualGrowth As Single
        Property IsBelly As Boolean
        Property Deprecated As Boolean
        Property Cdf As Single

        Property CdfAircraftMax As Single
        Property CdfAc As Single
        Property CdfAircraftMaxAc As Single
        Property CdfHma As Single
        Property CdfAircraftMaxHma As Single
        Property Cdf401 As Single
        Property CdfAircraftMax401 As Single
        Property CdfSub As Single
        Property CdfAircraftMaxSub As Single

        Property CtoP As Single
        Property Dtemp As Single
        Property CtoPAc As Single
        Property CtoPHma As Single
        Property CtoP401 As Single
        Property CtoPSub As Single
        Property Coverage As Double
        Property PCRThick As Double
        Property ACRThickMGW As Double
        Property ACRThick As Thickness
        Property ACRThick1 As Thickness
        Property ACRThick2 As Thickness
        Property ACRThick3 As Thickness
        Property ACRThick4 As Thickness
        Property ACRB As Double

        Property CriticalAnnualDeparture As Double
        Property TireWidth As Thickness
        Property TireLength As Thickness
        Property TireArea As Area
        Property D2 As Thickness
        Property d1 As Thickness
        Property d3 As Thickness
        Property D4 As Thickness
        Property d5 As Thickness
        Property D6 As Thickness
        Property aa As Thickness
        Property dmax As Thickness
        Property GX As Single
        Property GY As Single

        Property TirePressureF As Thickness
        Property AircraftNumber As Integer

        Property Tv As Single
        Property CDFGraphData As List(Of Single)
        Property PCRNumber As Single
        Property ACRCoverage As Integer
        Property ACRB1 As Single
        Property ACRB2 As Single
        Property ACRB3 As Single
        Property ACRB4 As Single

        'Property AcrSubCat1 As String
        'Property AcrSubCat2 As String
        'Property AcrSubCat3 As String
        'Property AcrSubCat4 As String

        Property DefaultCp As Pressure
        Property GearOrientation As Integer

    End Interface
End Namespace