Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net.Mime
Imports System.Runtime.Serialization
Imports System.Text
Imports ICAOModels
Imports ICAOModels.Interfaces
Imports Microsoft.VisualBasic.FileIO
Imports System.Text.RegularExpressions

Public Class clsAC
    Dim Temp As Single
    Dim Temp1 As Single


    Public Const BellyExt As String = " Belly"
    Public Const kPaTopsi As Double = 0.1450377438
    Public Const cmToin As Double = 0.3937008
    Public Const kgTolb As Double = 2.2046225

    Public ExternalLibraryIndex As Short


    Public Structure AircraftCharacteristics
        Dim libACName As String
        Dim libGL As Single
        Dim libMGpcnt As Single
        Dim libMGpcntPCN As Single 'ikawa PPPP
        Dim libCP As Single

        Dim libGear As String
        Dim libIGear As Short
        Dim libGearOrientation As Short
        Dim libTT As Single
        Dim libTS As Single
        Dim libTG As Single
        Dim libB As Single
        Dim libTV As Single '2017.05.25 izydor

        Dim libNTires As Short
        Dim libNTires1 As Short
        Dim libNTires2 As Short
        Dim libWheelCoordinate As Short
        Dim libNWheels As Short
        Dim libTX() As Single
        Dim libTY() As Single

        Dim libNEVPTS As Short
        Dim libNEVPTS1 As Short
        Dim libNEVPTS2 As Short
        Dim libEVPTX() As Single
        Dim libEVPTY() As Single

        Dim libNGroups As Short
        Dim libWheelsInGroup() As Short
        Dim libEvalPointsInGroup() As Short

        Dim libNTTrack As Short
        Dim libXAC() As Single
        Dim libTireContactArea As Single '(?)
        Dim libIsBelly As Boolean
    End Structure

    'Get directory of aircraft.xml file 
    Private _XMLFileLocation As String
    Public Property XMLFileLocation As String
        Get
            Return _XMLFileLocation
        End Get
        Set(ByVal value As String)
            _XMLFileLocation = value
        End Set
    End Property

    'Get directory of UDA directory 
    Private _UDAFileLocation As String
    Public Property UDAFileLocation As String
        Get
            Return _UDAFileLocation
        End Get
        Set(ByVal value As String)
            _UDAFileLocation = value
        End Set
    End Property

    Sub InitACLib(ByRef AC() As AircraftCharacteristics,
        ByRef LibACGroup() As Short,
        ByRef LibACGroupName() As String,
        ByRef libNAC As Short,
        ByRef NBelly As Short,
        ByRef NLibACGroups As Short,
        ByVal ExternalLibraryActive As Boolean)
        ' IG = aircraft group number.
        ' IA = aircraft number.
        ' NBelly = number of dual-wheel belly gear aircraft, like MD-11.
        ' The aircraft data is accessed by aircraft number, e.g.
        ' GL(I) = libGL(LibIndex(I))
        ' where I = index of aircraft within the current list
        ' and LibIndex(I) = IA
        ' Aircraft numbers within a group can be found using:
        ' StartGroupNo = LibACGroup(J):  EndGroupNo = LibACGroup(J + 1) - 1
        ' IA = StartGroupNo + IOffset - 1
        ' If IA > EndGroupNo Then NotInCurrentGroup
        ' where J = current group number = ILibACGroup
        ' IOffset = offset into current group.
        ' libNAC = number of aircraft in the library.
        ' NLibACGroups = number of aircraft groups in the library.

        ' GL = aircraft gross load (lb) (typically taxi weight).
        ' TX, TY = wheel coordinates on gear (in), TX = lateral, TY = forward.
        ' CP = contained (inflation) pressure (psi).
        ' NEVPTS = number of stress evaluation points.
        ' EVPTX, EVPTY = coordinates of stress evaluation points (in),
        '                X = lateral, Y = forward.
        ' GEAR$ = military gear type designation.
        ' IGEAR = LEDNEW gear type designation.
        ' TT, TS, TG, B = gear dimensions (definitions vary with gear type).
        ' NTTRACK = number of wheel tracks.
        ' XAC = lateral coordinates of the wheel tracks.

        'Library Aircraft aircraft.xml file
        Dim aircraftLibrary As SignedAircraftLibrary
        Dim ser As New DataContractSerializer(GetType(SignedAircraftLibrary))
        Dim xmlLibraryPath = XMLFileLocation
        Dim textReader As TextReader = New StreamReader(xmlLibraryPath)
        Dim xml As String = textReader.ReadToEnd()
        textReader.Close()

        Dim stream As New MemoryStream(Encoding.UTF8.GetBytes(xml))
        aircraftLibrary = DirectCast(ser.ReadObject(stream), SignedAircraftLibrary)

        Dim aircraftManufacturer As List(Of IAirplaneInfo) = aircraftLibrary.Airplanes.GroupBy(Function(x) x.Manufacturer).Select(Function(x) x.First).ToList()
        Dim orderedByAircraftNumber As List(Of IAirplaneInfo) = aircraftLibrary.Airplanes.OrderBy(Function(x) x.AircraftNumber).ToList()

        '' UDA directory
        Dim path As String
        Try
            path = UDAFileLocation
        Catch ex As Exception
            path = Nothing
        End Try

        Dim airplaneLibrary As List(Of AirplaneInfo) = New List(Of AirplaneInfo)

        Dim fi As FileInfo
        Dim di As New DirectoryInfo(path)
        Dim filelist As New List(Of String)

        Dim AAA As Integer

        Dim NumberofUserDefined As Integer = 0

        Try
            For Each fi In di.GetFiles
                If fi.Name.Contains("(UDA)") Then
                    filelist.Add(fi.Name)
                    NumberofUserDefined += 1
                End If
            Next
        Catch ex As Exception

        End Try

        For AAA = 0 To NumberofUserDefined - 1
            Dim loadPath As String = path + "\" + filelist.Item(AAA).ToString
            ser = New DataContractSerializer(GetType(List(Of AirplaneInfo)))
            textReader = New System.IO.StreamReader(loadPath.ToString())
            xml = textReader.ReadToEnd()
            textReader.Close()
            Dim xmlReplaced As String = Regex.Replace(xml, "http://schemas.datacontract.org/2004/07/\w+", "http://schemas.datacontract.org/2004/07/ICAOModels")
            stream = New MemoryStream(Encoding.UTF8.GetBytes(xmlReplaced))

            Dim airplaneInfos As List(Of AirplaneInfo)
            Dim airplaneinfo As AirplaneInfo

            Try
                airplaneInfos = DirectCast(ser.ReadObject(stream), List(Of AirplaneInfo))
                If airplaneInfos.Count > 0 Then
                    For Each airplaneinfo In airplaneInfos
                        airplaneLibrary.Add(airplaneinfo)
                    Next
                End If

            Catch ex As Exception

            End Try
        Next

        Dim IA, I, IG As Short
        Dim C, M As Single   ' Linear equation.

        NLibACGroups = Convert.ToInt16(aircraftManufacturer.Count())
        libNAC = Convert.ToInt16(aircraftLibrary.Airplanes.Count + airplaneLibrary.Count)

        ReDim AC(libNAC)
        ReDim LibACGroup(NLibACGroups - 1)
        ReDim LibACGroupName(NLibACGroups - 1)

        'add AC Group name to LibACGroupName
        For l = 0 To aircraftManufacturer.Count - 1
            LibACGroupName(l) = (aircraftManufacturer(l).Manufacturer)
        Next

        IA = 0
        IG = 0

        Dim previousGroup As String = ""
        Dim airplane As IAirplaneInfo

        'Library AircraftCharacteristics
        For Each airplane In orderedByAircraftNumber
            'IA = IA + 1S
            AC(IA).libACName = airplane.Name
            Dim check As Integer = 0
            If airplane.Name.Contains("MD-11") Then
                check = 1
            End If
            AC(IA).libCP = Convert.ToSingle(airplane.Cp.UsCustomary)
            AC(IA).libGL = Convert.ToSingle(airplane.GrossWeight.UsCustomary)
            AC(IA).libMGpcnt = Convert.ToSingle(airplane.MgPercent)

            AC(IA).libMGpcntPCN = Convert.ToSingle(airplane.MgPercentPCN)
            AC(IA).libGear = airplane.Gear
            AC(IA).libIGear = Convert.ToInt16(airplane.NumberGear)
            AC(IA).libGearOrientation = Convert.ToInt16(airplane.GearOrientation)
            AC(IA).libTT = Convert.ToSingle(airplane.Tt.UsCustomary)
            AC(IA).libTS = Convert.ToSingle(airplane.Ts)
            AC(IA).libTV = Convert.ToSingle(airplane.Tv)
            AC(IA).libTG = Convert.ToSingle(airplane.Tg)
            AC(IA).libB = Convert.ToSingle(airplane.B.UsCustomary)

            AC(IA).libNTires = Convert.ToInt16(airplane.NumberWheels)
            AC(IA).libNWheels = Convert.ToInt16(airplane.WheelCoordinates.Count)
            AC(IA).libWheelCoordinate = Convert.ToInt16(airplane.WheelCoordinates.Count)
            AC(IA).libNEVPTS = Convert.ToInt16(airplane.EvaluationPoints.Count)
            AC(IA).libNTTrack = Convert.ToInt16(airplane.NumberTireTrack)
            AC(IA).libIsBelly = airplane.IsBelly

            If AC(IA).libGear = "X" Then
                AC(IA).libNTires1 = Convert.ToInt16(airplane.NumberWheels / 2)
                If AC(IA).libNTires1 = 0 Then
                    AC(IA).libNTires1 = 1
                End If
                AC(IA).libNEVPTS1 = Convert.ToInt16(airplane.EvaluationPoints.Count)
                AC(IA).libNGroups = 1
                AC(I).libXAC = Nothing
            End If

            'ReDim AC(IA).libTX(AC(IA).libNTires)
            'ReDim AC(IA).libTY(AC(IA).libNTires)
            ReDim AC(IA).libTX(AC(IA).libWheelCoordinate)
            ReDim AC(IA).libTY(AC(IA).libWheelCoordinate)
            ReDim AC(IA).libXAC(AC(IA).libNTTrack)
            ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
            ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
            Dim K As Boolean = False
            If AC(IA).libACName = "F-15C" Then
                K = True

            End If

            'For j As Integer = 0 To airplane.NumberWheels - 1
            '    AC(IA).libTX(j + 1) = Convert.ToSingle(airplane.WheelCoordinates(j).X.UsCustomary)
            '    AC(IA).libTY(j + 1) = Convert.ToSingle(airplane.WheelCoordinates(j).Y.UsCustomary)
            'Next

            For j As Integer = 0 To airplane.WheelCoordinates.Count - 1
                AC(IA).libTX(j + 1) = Convert.ToSingle(airplane.WheelCoordinates(j).X.UsCustomary)
                AC(IA).libTY(j + 1) = Convert.ToSingle(airplane.WheelCoordinates(j).Y.UsCustomary)
            Next

            For j As Integer = 0 To airplane.EvaluationPoints.Count - 1
                AC(IA).libEVPTX(j + 1) = Convert.ToSingle(airplane.EvaluationPoints(j).X.UsCustomary)
                AC(IA).libEVPTY(j + 1) = Convert.ToSingle(airplane.EvaluationPoints(j).Y.UsCustomary)
            Next
            If airplane.NumberTireTrack > 0 Then
                For j As Integer = 0 To airplane.NumberTireTrack - 1
                    AC(IA).libXAC(j + 1) = Convert.ToSingle(airplane.TireTrackX(j).UsCustomary)
                Next
            End If
            IA = IA + 1S

            ''add number of aircraft in each group to LibACGroup
            For h As Integer = 0 To aircraftManufacturer.Count - 1
                If airplane.Manufacturer = aircraftManufacturer(h).Manufacturer Then
                    LibACGroup(h) += 1S
                End If
                IG = IG + 1S
            Next

            'Dim ManufacturerChecker As Integer = 0
            'For I = 0 To IG
            '    If airplane.Manufacturer = LibACGroupName(IG) Then
            '        ManufacturerChecker = 1
            '    End If
            'Next


            'If ManufacturerChecker = 1 Then
            '    LibACGroup(IG) = IA
            '    LibACGroupName(IG) = airplane.Manufacturer
            '    IG = IG + 1S
            '    previousGroup = airplane.Manufacturer
            'End If
        Next

        'UDA characteristics
        For Each airplane In airplaneLibrary
            'IA = IA + 1S
            AC(IA).libACName = airplane.Name
            AC(IA).libCP = Convert.ToSingle(airplane.Cp.UsCustomary)
            AC(IA).libGL = Convert.ToSingle(airplane.GrossWeight.UsCustomary)
            AC(IA).libMGpcnt = Convert.ToSingle(airplane.MgPercent)
            AC(IA).libMGpcntPCN = Convert.ToSingle(airplane.MgPercentPCN)
            AC(IA).libGear = airplane.Gear
            AC(IA).libIGear = Convert.ToInt16(airplane.NumberGear)
            AC(IA).libGearOrientation = Convert.ToInt16(airplane.GearOrientation)
            AC(IA).libTT = Convert.ToSingle(airplane.Tt.UsCustomary)
            AC(IA).libTS = Convert.ToSingle(airplane.Ts)
            AC(IA).libTV = Convert.ToSingle(airplane.Tv)
            AC(IA).libTG = Convert.ToSingle(airplane.Tg)
            AC(IA).libB = Convert.ToSingle(airplane.B.UsCustomary)

            AC(IA).libNTires = Convert.ToInt16(airplane.NumberWheels)
            AC(IA).libNWheels = Convert.ToInt16(airplane.WheelCoordinates.Count)
            AC(IA).libWheelCoordinate = Convert.ToInt16(airplane.WheelCoordinates.Count)

            AC(IA).libNEVPTS = Convert.ToInt16(airplane.EvaluationPoints.Count)
            AC(IA).libNTTrack = Convert.ToInt16(airplane.NumberTireTrack)

            If AC(IA).libGear = "X" Then
                AC(IA).libNTires1 = Convert.ToInt16(airplane.NumberWheels / 2)
                If AC(IA).libNTires1 = 0 Then
                    AC(IA).libNTires1 = 1
                End If
                AC(IA).libNEVPTS1 = Convert.ToInt16(airplane.EvaluationPoints.Count)
                AC(IA).libNGroups = 1
                AC(I).libXAC = Nothing
            End If

            'ReDim AC(IA).libTX(AC(IA).libNTires)
            'ReDim AC(IA).libTY(AC(IA).libNTires)
            ReDim AC(IA).libTX(AC(IA).libWheelCoordinate)
            ReDim AC(IA).libTY(AC(IA).libWheelCoordinate)
            ReDim AC(IA).libXAC(AC(IA).libNTTrack)
            ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
            ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)

            'For j As Integer = 0 To airplane.NumberWheels - 1
            '    AC(IA).libTX(j + 1) = Convert.ToSingle(airplane.WheelCoordinates(j).X.UsCustomary)
            '    AC(IA).libTY(j + 1) = Convert.ToSingle(airplane.WheelCoordinates(j).Y.UsCustomary)
            'Next

            For j As Integer = 0 To airplane.WheelCoordinates.Count - 1
                AC(IA).libTX(j) = Convert.ToSingle(airplane.WheelCoordinates(j).X.UsCustomary)
                AC(IA).libTY(j) = Convert.ToSingle(airplane.WheelCoordinates(j).Y.UsCustomary)
            Next

            For j As Integer = 0 To airplane.EvaluationPoints.Count - 1
                AC(IA).libEVPTX(j + 1) = Convert.ToSingle(airplane.EvaluationPoints(j).X.UsCustomary)
                AC(IA).libEVPTY(j + 1) = Convert.ToSingle(airplane.EvaluationPoints(j).Y.UsCustomary)
            Next

            If airplane.NumberTireTrack > 0 Then
                For j As Integer = 0 To airplane.NumberTireTrack - 1
                    AC(IA).libXAC(j + 1) = Convert.ToSingle(airplane.TireTrackX(j).UsCustomary)
                Next
            End If

            If AC(IA).libACName.Contains("UDA") Then

                If airplane.NumberWheels = 3 Then
                    ReDim AC(IA).libTX(AC(IA).libNTires * 2)
                    ReDim AC(IA).libTY(AC(IA).libNTires * 2)
                    AC(IA).libNTires = Convert.ToInt16(airplane.NumberWheels * 2)
                    AC(IA).libMGpcnt = Convert.ToSingle(airplane.MgPercent * 2)


                    For j As Integer = 0 To airplane.NumberWheels * 2 - 1
                        AC(IA).libTX(j + 1) = Convert.ToSingle(airplane.WheelCoordinates(j).X.UsCustomary)
                        AC(IA).libTY(j + 1) = Convert.ToSingle(airplane.WheelCoordinates(j).Y.UsCustomary)
                    Next
                End If
            End If
            IA = IA + 1S

            'add number of aircraft in "External Library" to LibACGroup
            LibACGroup(LibACGroup.Count - 2) += 1S
        Next

        libNAC = IA
        Temp = 0
        Temp1 = 0
        'NLibACGroups = IG
        NBelly = 0

        ' Set derived data where possible.
        For IA = 0 To libNAC - 1S
            '   Dual-tandem belly gear added. GFH 04/22/03.
            If AC(IA).libGear$ = "H" Then  ' Set belly gear data.

                If AC(IA).libACName = "IL86" Then
                    IA = IA
                End If

                Dim cont1 As Boolean
                'Call Check_H_type(IA)
                cont1 = False
                Call Check_H_type(AC, IA, I, NBelly, libNAC, cont1)
                If cont1 Then Continue For


                NBelly = NBelly + 1S
                I = libNAC + NBelly
                ReDim Preserve AC(I)
                AC(I).libACName$ = AC(IA).libACName$ & BellyExt$
                AC(I).libGL = AC(IA).libGL
                AC(I).libMGpcnt = CSng(0.95 - 2.0! * AC(IA).libMGpcnt)
                If AC(IA).libACName$ = "DC-10-30" Or AC(IA).libACName$ = "DC10-30/40" _
                   Or AC(IA).libACName$ = "KC-10" Then
                    AC(I).libCP = 153.0!
                ElseIf AC(IA).libACName$ = "MD-11" Then
                    AC(I).libCP = 180.0!
                    'ElseIf AC(IA).libACName$ = "A340-500/600" Then
                ElseIf Mid(AC(IA).libACName$, 1, 8) = "A340-500" Or Mid(AC(IA).libACName$, 1, 8) = "A340-600" Then
                    AC(I).libMGpcnt = AC(IA).libMGpcnt ' Special case.
                    AC(I).libCP = 222.0!
                    'ElseIf AC(IA).libACName$ = "A340-200/300" Then
                ElseIf Mid(AC(IA).libACName$, 1, 8) = "A340-200" Or Mid(AC(IA).libACName$, 1, 8) = "A340-300" Then
                    AC(I).libCP = 158
                    'ElseIf Mid(AC(IA).libACName$, 1, 8) = "A350-900" Then
                    '    AC(I).libCP = 240.8
                    'AC(I).libMGpcnt = 0.4
                Else
                    AC(I).libCP = AC(IA).libCP * 2.0! * AC(I).libMGpcnt / AC(IA).libMGpcnt
                End If

                AC(I).libGear$ = "HB"
                AC(I).libIGear = 2      ' Same as single-wheel or C-130, two gears, for coverage to pass.
                AC(I).libTT = AC(IA).libTG : AC(I).libTS = 0.0!
                AC(IA).libTG = 0.0!
                AC(I).libTG = 0.0! : AC(I).libB = 0.0!
                AC(I).libNTires = 2
                ReDim AC(I).libTX(AC(I).libNTires)
                ReDim AC(I).libTY(AC(I).libNTires)
                AC(I).libTX(1) = CSng(-AC(I).libTT * 0.5) : AC(I).libTY(1) = 0.0!
                AC(I).libTX(2) = -AC(I).libTX(1) : AC(I).libTY(2) = 0.0!


                If AC(IA).libACName = "IL86" Then
                    IA = IA
                    AC(I).libCP = AC(IA).libCP
                End If


                If Mid(AC(IA).libACName$, 1, 8) = "A340-500" Or Mid(AC(IA).libACName$, 1, 8) = "A340-600" _
                 Or AC(IA).libACName = "IL86" Then
                    AC(I).libB = AC(IA).libB
                    AC(I).libNTires = 4
                    ReDim AC(I).libTX(AC(I).libNTires)
                    ReDim AC(I).libTY(AC(I).libNTires)
                    AC(I).libTX(1) = CSng(-AC(I).libTT * 0.5) : AC(I).libTY(1) = 0.0!
                    AC(I).libTX(2) = CSng(AC(I).libTT * 0.5) : AC(I).libTY(2) = 0.0!
                    AC(I).libTX(3) = AC(I).libTX(1) : AC(I).libTY(3) = AC(I).libB
                    AC(I).libTX(4) = AC(I).libTX(2) : AC(I).libTY(4) = AC(I).libB
                End If

                AC(I).libNTTrack = 2
                ReDim AC(I).libXAC(AC(I).libNTTrack)
                AC(I).libXAC(1) = CSng(-AC(I).libTT * 0.5)
                AC(I).libXAC(2) = CSng(AC(I).libTT * 0.5)
                AC(I).libNEVPTS = 4
                ReDim AC(I).libEVPTX(AC(I).libNEVPTS)
                ReDim AC(I).libEVPTY(AC(I).libNEVPTS)
                AC(I).libEVPTX(1) = CSng(AC(I).libTT * 0.5) : AC(I).libEVPTY(1) = 0.0!
                AC(I).libEVPTX(2) = CSng(AC(I).libTT * 0.375) : AC(I).libEVPTY(2) = 0.0!
                AC(I).libEVPTX(3) = CSng(AC(I).libTT * 0.25) : AC(I).libEVPTY(3) = 0.0!
                AC(I).libEVPTX(4) = 0.0! : AC(I).libEVPTY(4) = 0.0!
                If Mid(AC(IA).libACName$, 1, 8) = "A340-500" Or Mid(AC(IA).libACName$, 1, 8) = "A340-600" _
                 Or Mid(AC(IA).libACName$, 1, 4) = "IL86" Then
                    AC(I).libNEVPTS = 8 ' Dual-tandem belly gear.
                    ReDim AC(I).libEVPTX(AC(I).libNEVPTS)
                    ReDim AC(I).libEVPTY(AC(I).libNEVPTS)
                    Temp = CSng(AC(I).libTT * 0.5)
                    C = CSng(0.5604 * AC(I).libTT - 0.2637 * AC(I).libB)
                    If C < 0.0! Then C = 0.0!
                    M = -C / Temp
                    AC(I).libEVPTX(1) = Temp
                    AC(I).libEVPTY(1) = 0.0!
                    AC(I).libEVPTX(2) = CSng(Temp * 0.8)
                    AC(I).libEVPTY(2) = C + M * AC(I).libEVPTX(2)
                    AC(I).libEVPTX(3) = CSng(Temp * 0.6)
                    AC(I).libEVPTY(3) = C + M * AC(I).libEVPTX(3)
                    AC(I).libEVPTX(4) = CSng(Temp * 0.4)
                    AC(I).libEVPTY(4) = C + M * AC(I).libEVPTX(4)
                    AC(I).libEVPTX(5) = CSng(Temp * 0.2)
                    AC(I).libEVPTY(5) = C + M * AC(I).libEVPTX(5)
                    AC(I).libEVPTX(6) = 0.0!
                    AC(I).libEVPTY(6) = C
                    AC(I).libEVPTX(7) = 0.0!
                    AC(I).libEVPTY(7) = CSng((AC(I).libB * 0.5 - C) * 0.5 + C)
                    AC(I).libEVPTX(8) = 0.0!
                    AC(I).libEVPTY(8) = CSng(AC(I).libB * 0.5)
                End If
            End If
        Next IA

        For IA = 1 To I

            If AC(IA).libGear = "A" Or AC(IA).libGear = "B" Then
                If AC(IA).libACName = "Truck Axle Single" Then
                    AC(IA).libTT = 96.0!
                End If
                AC(IA).libNTires = 1
                ReDim AC(IA).libTX(AC(IA).libNTires)
                ReDim AC(IA).libTY(AC(IA).libNTires)
                AC(IA).libTX(1) = 0.0! : AC(IA).libTY(1) = 0.0!
                AC(IA).libTS = 0.0! : AC(IA).libTG = 0.0!
                AC(IA).libB = 0.0!
                If AC(IA).libGear = "A" Then ' 1 wheel, 1 gear.
                    AC(IA).libNTTrack = 1
                    ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                    AC(IA).libXAC(1) = 0.0!
                Else ' 1 wheel, 2 gears.
                    AC(IA).libNTTrack = 2
                    ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                    AC(IA).libXAC(1) = -AC(IA).libTT * CSng(0.5)
                    AC(IA).libXAC(2) = -AC(IA).libXAC(1)
                End If
                AC(IA).libNEVPTS = 1
                ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
                AC(IA).libEVPTX(1) = 0.0! : AC(IA).libEVPTY(1) = 0.0!

            ElseIf AC(IA).libGear = "D" Then ' Dual-wheel, 2 gears.

            ElseIf AC(IA).libGear = "X" Then ' Dual-wheel, 2 gears.

            ElseIf AC(IA).libGear = "F" Or AC(IA).libGear = "H" Or AC(IA).libACName = "IL86" Then

                AC(IA).libNTires = 4
                ReDim AC(IA).libTX(AC(IA).libNTires)
                ReDim AC(IA).libTY(AC(IA).libNTires)

                If (InStr(AC(IA).libACName, "747", CompareMethod.Text) > 0) _
                Or (InStr(AC(IA).libACName, "380", CompareMethod.Text) > 0) Then
                    AC(IA).libTX(1) = -AC(IA).libTT * CSng(0.5) : AC(IA).libTY(1) = AC(IA).libB
                    AC(IA).libTX(2) = -AC(IA).libTT * CSng(0.5) : AC(IA).libTY(2) = 0.0!
                    AC(IA).libTX(3) = AC(IA).libTT * CSng(0.5) : AC(IA).libTY(3) = 0.0!
                    AC(IA).libTX(4) = AC(IA).libTT * CSng(0.5) : AC(IA).libTY(4) = AC(IA).libB
                Else
                    AC(IA).libTX(1) = -AC(IA).libTT * CSng(0.5) : AC(IA).libTY(1) = 0.0!
                    AC(IA).libTX(2) = AC(IA).libTT * CSng(0.5) : AC(IA).libTY(2) = 0.0!
                    AC(IA).libTX(3) = AC(IA).libTX(1) : AC(IA).libTY(3) = AC(IA).libB
                    AC(IA).libTX(4) = AC(IA).libTX(2) : AC(IA).libTY(4) = AC(IA).libB
                End If



                If Not AC(IA).libACName = "IL86" Then
                    AC(IA).libNTTrack = 4
                    ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                    AC(IA).libXAC(1) = -AC(IA).libTS * CSng(0.5) - AC(IA).libTT
                    AC(IA).libXAC(2) = -AC(IA).libTS * CSng(0.5)
                    AC(IA).libXAC(3) = -AC(IA).libXAC(2)
                    AC(IA).libXAC(4) = -AC(IA).libXAC(1)
                Else
                    AC(IA).libNTTrack = 6
                    ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                    AC(IA).libXAC(1) = CSng(-AC(IA).libTS * CSng(0.5) - AC(IA).libTT)
                    AC(IA).libXAC(2) = CSng(-AC(IA).libTS * CSng(0.5))
                    AC(IA).libXAC(3) = CSng(-AC(IA).libTT * CSng(0.5))
                    AC(IA).libXAC(4) = -AC(IA).libXAC(3)
                    AC(IA).libXAC(5) = -AC(IA).libXAC(2)
                    AC(IA).libXAC(6) = -AC(IA).libXAC(1)
                End If

                AC(IA).libNEVPTS = 8
                ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
                Temp = AC(IA).libTT * CSng(0.5)
                C = CSng(0.5604) * AC(IA).libTT - CSng(0.2637) * AC(IA).libB
                If C < 0.0! Then C = 0.0!
                M = -C / Temp
                AC(IA).libEVPTX(1) = Temp
                AC(IA).libEVPTY(1) = 0.0!
                AC(IA).libEVPTX(2) = Temp * CSng(0.8)
                AC(IA).libEVPTY(2) = C + M * AC(IA).libEVPTX(2)
                AC(IA).libEVPTX(3) = Temp * CSng(0.6)
                AC(IA).libEVPTY(3) = C + M * AC(IA).libEVPTX(3)
                AC(IA).libEVPTX(4) = Temp * CSng(0.4)
                AC(IA).libEVPTY(4) = C + M * AC(IA).libEVPTX(4)
                AC(IA).libEVPTX(5) = Temp * CSng(0.2)
                AC(IA).libEVPTY(5) = C + M * AC(IA).libEVPTX(5)
                AC(IA).libEVPTX(6) = 0.0!
                AC(IA).libEVPTY(6) = C
                AC(IA).libEVPTX(7) = 0.0!
                AC(IA).libEVPTY(7) = (AC(IA).libB * CSng(0.5) - C) * CSng(0.5) + C
                AC(IA).libEVPTX(8) = 0.0!
                AC(IA).libEVPTY(8) = AC(IA).libB * CSng(0.5)

            ElseIf AC(IA).libGear = "J" Then  ' 747 configuration.

                If AC(IA).libACName <> "B-747-16" Then
                    AC(IA).libNTires = 4
                    ReDim AC(IA).libTX(AC(IA).libNTires)
                    ReDim AC(IA).libTY(AC(IA).libNTires)

                    AC(IA).libTX(1) = CSng(-AC(IA).libTT * CSng(0.5)) : AC(IA).libTY(1) = 0.0!
                    AC(IA).libTX(2) = AC(IA).libTT * CSng(0.5) : AC(IA).libTY(2) = 0.0!
                    AC(IA).libTX(3) = AC(IA).libTX(1) : AC(IA).libTY(3) = AC(IA).libB
                    AC(IA).libTX(4) = AC(IA).libTX(2) : AC(IA).libTY(4) = AC(IA).libB
                End If
                AC(IA).libNTTrack = 8
                ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                AC(IA).libXAC(1) = -AC(IA).libTG * CSng(0.5) - 2.0! * AC(IA).libTT - AC(IA).libTS
                AC(IA).libXAC(2) = AC(IA).libXAC(1) + AC(IA).libTT
                AC(IA).libXAC(3) = AC(IA).libXAC(2) + AC(IA).libTS
                AC(IA).libXAC(4) = AC(IA).libXAC(3) + AC(IA).libTT
                AC(IA).libXAC(5) = -AC(IA).libXAC(4)
                AC(IA).libXAC(6) = -AC(IA).libXAC(3)
                AC(IA).libXAC(7) = -AC(IA).libXAC(2)
                AC(IA).libXAC(8) = -AC(IA).libXAC(1)
                If AC(IA).libACName <> "B-747-16" Then
                    AC(IA).libNEVPTS = 8
                    ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                    ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)

                    Temp = CSng(AC(IA).libTT * CSng(0.5))
                    C = CSng(0.5604 * AC(IA).libTT - 0.2637 * AC(IA).libB)
                    If C < 0.0! Then C = 0.0!
                    M = -C / Temp
                    AC(IA).libEVPTX(1) = Temp
                    AC(IA).libEVPTY(1) = 0.0!
                    AC(IA).libEVPTX(2) = CSng(Temp * 0.8)
                    AC(IA).libEVPTY(2) = C + M * AC(IA).libEVPTX(2)
                    AC(IA).libEVPTX(3) = CSng(Temp * 0.6)
                    AC(IA).libEVPTY(3) = C + M * AC(IA).libEVPTX(3)
                    AC(IA).libEVPTX(4) = CSng(Temp * 0.4)
                    AC(IA).libEVPTY(4) = C + M * AC(IA).libEVPTX(4)
                    AC(IA).libEVPTX(5) = CSng(Temp * 0.2)
                    AC(IA).libEVPTY(5) = C + M * AC(IA).libEVPTX(5)
                    AC(IA).libEVPTX(6) = 0.0!
                    AC(IA).libEVPTY(6) = C
                    AC(IA).libEVPTX(7) = 0.0!
                    AC(IA).libEVPTY(7) = CSng((AC(IA).libB * CSng(0.5) - C) * CSng(0.5) + C)
                    AC(IA).libEVPTX(8) = 0.0!
                    AC(IA).libEVPTY(8) = CSng(AC(IA).libB * CSng(0.5))
                End If
            ElseIf AC(IA).libACName = "A350-1000" Then  ' 777 configuration.
                AC(IA).libNTires = 6
                Temp = CSng(AC(IA).libTT * CSng(0.5)) : Temp1 = AC(IA).libB
                Dim Temp2 As Single
                Temp2 = 58.032 / 2

                ReDim AC(IA).libTX(AC(IA).libNTires)
                ReDim AC(IA).libTY(AC(IA).libNTires)

                AC(IA).libTX(1) = -Temp : AC(IA).libTY(1) = -Temp1
                AC(IA).libTX(2) = Temp : AC(IA).libTY(2) = -Temp1
                AC(IA).libTX(3) = -Temp2 : AC(IA).libTY(3) = 0.0!
                AC(IA).libTX(4) = Temp2 : AC(IA).libTY(4) = 0.0!
                AC(IA).libTX(5) = -Temp : AC(IA).libTY(5) = Temp1
                AC(IA).libTX(6) = Temp : AC(IA).libTY(6) = Temp1

                AC(IA).libNEVPTS = 6
                ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)

                Temp = AC(IA).libTT
                Temp2 = 58.032
                AC(IA).libEVPTX(1) = CSng(Temp2 * CSng(0.5)) : AC(IA).libEVPTY(1) = 0.0!
                AC(IA).libEVPTX(2) = CSng(Temp2 * 0.4) : AC(IA).libEVPTY(2) = 0.0!
                AC(IA).libEVPTX(3) = CSng(Temp2 * 0.3) : AC(IA).libEVPTY(3) = 0.0!
                AC(IA).libEVPTX(4) = CSng(Temp2 * 0.2) : AC(IA).libEVPTY(4) = 0.0!
                AC(IA).libEVPTX(5) = CSng(Temp2 * 0.1) : AC(IA).libEVPTY(5) = 0.0!
                AC(IA).libEVPTX(6) = 0.0! : AC(IA).libEVPTY(6) = 0.0!


                AC(IA).libNTTrack = 4
                ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                AC(IA).libXAC(1) = CSng(-AC(IA).libTS * CSng(0.5) - AC(IA).libTT)
                AC(IA).libXAC(2) = CSng(-AC(IA).libTS * CSng(0.5))
                AC(IA).libXAC(3) = -AC(IA).libXAC(2)
                AC(IA).libXAC(4) = -AC(IA).libXAC(1)

            ElseIf AC(IA).libGear = "N" Then  ' 777 configuration.

                AC(IA).libNTires = 6
                Temp = CSng(AC(IA).libTT * CSng(0.5)) : Temp1 = AC(IA).libB
                Dim Temp2 As Single
                Temp2 = 58.032 / 2
                If AC(IA).libACName <> "TU154B" Then
                    ReDim AC(IA).libTX(AC(IA).libNTires)
                    ReDim AC(IA).libTY(AC(IA).libNTires)

                    AC(IA).libTX(1) = -Temp : AC(IA).libTY(1) = -Temp1
                    AC(IA).libTX(2) = Temp : AC(IA).libTY(2) = -Temp1
                    AC(IA).libTX(3) = -Temp : AC(IA).libTY(3) = 0.0!
                    AC(IA).libTX(4) = Temp : AC(IA).libTY(4) = 0.0!
                    AC(IA).libTX(5) = -Temp : AC(IA).libTY(5) = Temp1
                    AC(IA).libTX(6) = Temp : AC(IA).libTY(6) = Temp1

                End If

                AC(IA).libNEVPTS = 6
                ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)

                Temp = AC(IA).libTT
                AC(IA).libEVPTX(1) = CSng(Temp * CSng(0.5)) : AC(IA).libEVPTY(1) = 0.0!
                AC(IA).libEVPTX(2) = CSng(Temp * 0.4) : AC(IA).libEVPTY(2) = 0.0!
                AC(IA).libEVPTX(3) = CSng(Temp * 0.3) : AC(IA).libEVPTY(3) = 0.0!
                AC(IA).libEVPTX(4) = CSng(Temp * 0.2) : AC(IA).libEVPTY(4) = 0.0!
                AC(IA).libEVPTX(5) = CSng(Temp * 0.1) : AC(IA).libEVPTY(5) = 0.0!
                AC(IA).libEVPTX(6) = 0.0! : AC(IA).libEVPTY(6) = 0.0!


                AC(IA).libNTTrack = 4
                ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                AC(IA).libXAC(1) = CSng(-AC(IA).libTS * CSng(0.5) - AC(IA).libTT)
                AC(IA).libXAC(2) = CSng(-AC(IA).libTS * CSng(0.5))
                AC(IA).libXAC(3) = -AC(IA).libXAC(2)
                AC(IA).libXAC(4) = -AC(IA).libXAC(1)


            ElseIf AC(IA).libGear = "X" Then  ' IL76T
                If AC(IA).libACName.Contains("UDA") Then
                    Try



                        'AC(IA).libNTTrack = 8
                        'ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                        'AC(IA).libXAC(1) = CSng(-AC(IA).libTS * CSng(0.5) - 2 * AC(IA).libTT - AC(IA).libTG)
                        'AC(IA).libXAC(2) = CSng(-AC(IA).libTS * CSng(0.5) - AC(IA).libTT - AC(IA).libTG)
                        'AC(IA).libXAC(3) = CSng(-AC(IA).libTS * CSng(0.5) - AC(IA).libTT)
                        'AC(IA).libXAC(4) = CSng(-AC(IA).libTS * CSng(0.5))
                        'AC(IA).libXAC(5) = -AC(IA).libXAC(4)
                        'AC(IA).libXAC(6) = -AC(IA).libXAC(3)
                        'AC(IA).libXAC(7) = -AC(IA).libXAC(2)
                        'AC(IA).libXAC(8) = -AC(IA).libXAC(1)
                        'AC(IA).libXAC = Nothing
                        'AC(IA).libNEVPTS = 11
                        'ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                        'ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)

                        'Temp = AC(IA).libTT * CSng(0.5)
                        'C = CSng(0.5604) * AC(IA).libTG - CSng(0.2637) * AC(IA).libB
                        'If C < 0.0! Then C = 0.0!
                        'M = -C / Temp
                        'AC(IA).libEVPTX(1) = Temp
                        'AC(IA).libEVPTY(1) = 0.0!
                        'AC(IA).libEVPTX(2) = Temp * CSng(0.8)
                        'AC(IA).libEVPTY(2) = C + M * AC(IA).libEVPTX(2)
                        'AC(IA).libEVPTX(3) = Temp * CSng(0.6)
                        'AC(IA).libEVPTY(3) = C + M * AC(IA).libEVPTX(3)
                        'AC(IA).libEVPTX(4) = Temp * CSng(0.4)
                        'AC(IA).libEVPTY(4) = C + M * AC(IA).libEVPTX(4)
                        'AC(IA).libEVPTX(5) = Temp * CSng(0.2)
                        'AC(IA).libEVPTY(5) = C + M * AC(IA).libEVPTX(5)
                        'AC(IA).libEVPTX(6) = 0.0!
                        'AC(IA).libEVPTY(6) = C
                        'AC(IA).libEVPTX(7) = 0.0!
                        'AC(IA).libEVPTY(7) = (AC(IA).libB * CSng(0.5) - C) * CSng(0.5) + C
                        'AC(IA).libEVPTX(8) = 0.0!
                        'AC(IA).libEVPTY(8) = AC(IA).libB * CSng(0.5)

                        'AC(IA).libEVPTX(9) = AC(IA).libTX(3)
                        'AC(IA).libEVPTY(9) = AC(IA).libTY(3)
                        'AC(IA).libEVPTX(10) = AC(IA).libTX(4)
                        'AC(IA).libEVPTY(10) = AC(IA).libTY(4)
                        'AC(IA).libEVPTX(11) = AC(IA).libTX(3) + AC(IA).libTT / 2
                        'AC(IA).libEVPTY(11) = AC(IA).libTY(3)
                    Catch ex As Exception

                    End Try

                Else
                    AC(IA).libNTTrack = 8
                    ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                    AC(IA).libXAC(1) = CSng(-AC(IA).libTS * CSng(0.5) - 2 * AC(IA).libTT - AC(IA).libTG)
                    AC(IA).libXAC(2) = CSng(-AC(IA).libTS * CSng(0.5) - AC(IA).libTT - AC(IA).libTG)
                    AC(IA).libXAC(3) = CSng(-AC(IA).libTS * CSng(0.5) - AC(IA).libTT)
                    AC(IA).libXAC(4) = CSng(-AC(IA).libTS * CSng(0.5))
                    AC(IA).libXAC(5) = -AC(IA).libXAC(4)
                    AC(IA).libXAC(6) = -AC(IA).libXAC(3)
                    AC(IA).libXAC(7) = -AC(IA).libXAC(2)
                    AC(IA).libXAC(8) = -AC(IA).libXAC(1)

                    AC(IA).libNEVPTS = 11
                    ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                    ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)

                    Temp = AC(IA).libTT * CSng(0.5)
                    C = CSng(0.5604) * AC(IA).libTG - CSng(0.2637) * AC(IA).libB
                    If C < 0.0! Then C = 0.0!
                    M = -C / Temp
                    AC(IA).libEVPTX(1) = Temp
                    AC(IA).libEVPTY(1) = 0.0!
                    AC(IA).libEVPTX(2) = Temp * CSng(0.8)
                    AC(IA).libEVPTY(2) = C + M * AC(IA).libEVPTX(2)
                    AC(IA).libEVPTX(3) = Temp * CSng(0.6)
                    AC(IA).libEVPTY(3) = C + M * AC(IA).libEVPTX(3)
                    AC(IA).libEVPTX(4) = Temp * CSng(0.4)
                    AC(IA).libEVPTY(4) = C + M * AC(IA).libEVPTX(4)
                    AC(IA).libEVPTX(5) = Temp * CSng(0.2)
                    AC(IA).libEVPTY(5) = C + M * AC(IA).libEVPTX(5)
                    AC(IA).libEVPTX(6) = 0.0!
                    AC(IA).libEVPTY(6) = C
                    AC(IA).libEVPTX(7) = 0.0!
                    AC(IA).libEVPTY(7) = (AC(IA).libB * CSng(0.5) - C) * CSng(0.5) + C
                    AC(IA).libEVPTX(8) = 0.0!
                    AC(IA).libEVPTY(8) = AC(IA).libB * CSng(0.5)

                    AC(IA).libEVPTX(9) = AC(IA).libTX(3)
                    AC(IA).libEVPTY(9) = AC(IA).libTY(3)
                    AC(IA).libEVPTX(10) = AC(IA).libTX(4)
                    AC(IA).libEVPTY(10) = AC(IA).libTY(4)
                    AC(IA).libEVPTX(11) = AC(IA).libTX(3) + AC(IA).libTT / 2
                    AC(IA).libEVPTY(11) = AC(IA).libTY(3)
                End If


            ElseIf AC(IA).libGear = "Z" Then  ' An-124, An-225

                AC(IA).libNTTrack = 4
                ReDim AC(IA).libXAC(AC(IA).libNTTrack)
                AC(IA).libXAC(1) = -AC(IA).libTS * CSng(0.5) - AC(IA).libTT
                AC(IA).libXAC(2) = -AC(IA).libTS * CSng(0.5)
                AC(IA).libXAC(3) = -AC(IA).libXAC(2)
                AC(IA).libXAC(4) = -AC(IA).libXAC(1)


                If AC(IA).libACName = "An-124" Then
                    AC(IA).libNEVPTS = 12
                    ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                    ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
                    AC(IA).libEVPTX(1) = AC(IA).libTT * CSng(0.5) : AC(IA).libEVPTY(1) = 0 * AC(IA).libB
                    AC(IA).libEVPTX(2) = AC(IA).libTT * CSng(0.375) : AC(IA).libEVPTY(2) = 0 * AC(IA).libB
                    AC(IA).libEVPTX(3) = AC(IA).libTT * CSng(0.25) : AC(IA).libEVPTY(3) = 0 * AC(IA).libB
                    AC(IA).libEVPTX(4) = 0.0! : AC(IA).libEVPTY(4) = 0 * AC(IA).libB

                    AC(IA).libEVPTX(5) = AC(IA).libTT * CSng(0.5) : AC(IA).libEVPTY(5) = 1 * AC(IA).libB
                    AC(IA).libEVPTX(6) = AC(IA).libTT * CSng(0.375) : AC(IA).libEVPTY(6) = 1 * AC(IA).libB
                    AC(IA).libEVPTX(7) = AC(IA).libTT * CSng(0.25) : AC(IA).libEVPTY(7) = 1 * AC(IA).libB
                    AC(IA).libEVPTX(8) = 0.0! : AC(IA).libEVPTY(8) = 1 * AC(IA).libB

                    AC(IA).libEVPTX(9) = AC(IA).libTT * CSng(0.5) : AC(IA).libEVPTY(9) = 2 * AC(IA).libB
                    AC(IA).libEVPTX(10) = AC(IA).libTT * CSng(0.375) : AC(IA).libEVPTY(10) = 2 * AC(IA).libB
                    AC(IA).libEVPTX(11) = AC(IA).libTT * CSng(0.25) : AC(IA).libEVPTY(11) = 2 * AC(IA).libB
                    AC(IA).libEVPTX(12) = 0.0! : AC(IA).libEVPTY(12) = 2 * AC(IA).libB

                End If

                If AC(IA).libACName = "An-225" Then
                    AC(IA).libNEVPTS = 16
                    ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                    ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
                    AC(IA).libEVPTX(1) = AC(IA).libTT * CSng(0.5) : AC(IA).libEVPTY(1) = 0 * AC(IA).libB
                    AC(IA).libEVPTX(2) = AC(IA).libTT * CSng(0.375) : AC(IA).libEVPTY(2) = 0 * AC(IA).libB
                    AC(IA).libEVPTX(3) = AC(IA).libTT * CSng(0.25) : AC(IA).libEVPTY(3) = 0 * AC(IA).libB
                    AC(IA).libEVPTX(4) = 0.0! : AC(IA).libEVPTY(4) = 0 * AC(IA).libB

                    AC(IA).libEVPTX(5) = AC(IA).libTT * CSng(0.5) : AC(IA).libEVPTY(5) = 1 * AC(IA).libB
                    AC(IA).libEVPTX(6) = AC(IA).libTT * CSng(0.375) : AC(IA).libEVPTY(6) = 1 * AC(IA).libB
                    AC(IA).libEVPTX(7) = AC(IA).libTT * CSng(0.25) : AC(IA).libEVPTY(7) = 1 * AC(IA).libB
                    AC(IA).libEVPTX(8) = 0.0! : AC(IA).libEVPTY(8) = 1 * AC(IA).libB

                    AC(IA).libEVPTX(9) = AC(IA).libTT * CSng(0.5) : AC(IA).libEVPTY(9) = 2 * AC(IA).libB
                    AC(IA).libEVPTX(10) = AC(IA).libTT * CSng(0.375) : AC(IA).libEVPTY(10) = 2 * AC(IA).libB
                    AC(IA).libEVPTX(11) = AC(IA).libTT * CSng(0.25) : AC(IA).libEVPTY(11) = 2 * AC(IA).libB
                    AC(IA).libEVPTX(12) = 0.0! : AC(IA).libEVPTY(12) = 2 * AC(IA).libB

                    AC(IA).libEVPTX(13) = AC(IA).libTT * CSng(0.5) : AC(IA).libEVPTY(13) = 3 * AC(IA).libB
                    AC(IA).libEVPTX(14) = AC(IA).libTT * CSng(0.375) : AC(IA).libEVPTY(14) = 3 * AC(IA).libB
                    AC(IA).libEVPTX(15) = AC(IA).libTT * CSng(0.25) : AC(IA).libEVPTY(15) = 3 * AC(IA).libB
                    AC(IA).libEVPTX(16) = 0.0! : AC(IA).libEVPTY(16) = 3 * AC(IA).libB

                End If


            End If

        Next IA
        libNAC = libNAC + NBelly



        '''''''''ikawa 05/05/03
        ''''''''FileOpen(5, "acAllFEDFAA.dat", OpenMode.Output, , , 1024)

        '''''''''PrintLine(5, "  lp" & "   Aircraft " & LPad(17, " libGL ") & " MGpcnt " & " libCP " _
        '''''''''& " Gear " & "IGear" & "  libTT" & "    libTS" & "    libTG" & "     libB" _
        '''''''''& "    NTires " & "NTTrack " & "NEVPTS")

        ''''''''For IA = 1 To libNAC 'UBound(libACName, 1)
        ''''''''    'PrintLine(5, LPad(4, CStr(IA)) & " " & AC(IA).libACName & Space(18 - Len(AC(IA).libACName)) _
        ''''''''    '  & LPad(9, CStr(AC(IA).libGL)) & LPad(8, CStr(AC(IA).libMGpcnt)) & LPad(7, CStr(Math.Round(AC(IA).libCP, 2))) _
        ''''''''    '  & LPad(6, CStr(AC(IA).libGear$)) & LPad(4, CStr(AC(IA).libIGear)) _
        ''''''''    '  & LPad(9, CStr(AC(IA).libTT)) & LPad(9, CStr(AC(IA).libTS)) _
        ''''''''    '  & LPad(9, CStr(AC(IA).libTG)) & LPad(9, CStr(AC(IA).libB)) & "   " _
        ''''''''    '  & LPad(7, CStr(AC(IA).libNTires)) _
        ''''''''    '  & LPad(7, CStr(AC(IA).libNTTrack)) _
        ''''''''    '  & LPad(7, CStr(AC(IA).libNEVPTS)))

        ''''''''    'PrintLine(5, "")
        ''''''''    Dim max1 As Short
        ''''''''    Dim SS As String

        ''''''''    max1 = Math.Max(Math.Max(AC(IA).libNTires, AC(IA).libNTTrack), AC(IA).libNEVPTS)

        ''''''''    For I = 1 To max1 'UBound(libTX, 2)
        ''''''''        PrintLine(5, LPadArray(12, "TX =", AC(IA).libTX, I) & LPadArray(14, "  TY =", AC(IA).libTY, I) _
        ''''''''        & LPadArray(19, "  EVPTX =", AC(IA).libEVPTX, I) & LPadArray(19, "  EVPTY =", AC(IA).libEVPTY, I) _
        ''''''''        & LPadArray(20, "   libXAC =", AC(IA).libXAC, I) & "   " & CStr(IA) & " " & AC(IA).libACName)
        ''''''''    Next I
        ''''''''    PrintLine(5, "")

        ''''''''Next IA
        ''''''''FileClose(5)

    End Sub

    Sub Check_H_type(ByRef AC() As AircraftCharacteristics, ByRef IA As Short,
                     ByVal I As Integer, ByRef NBelly As Short, ByVal libNAC As Integer,
                     ByRef cont1 As Boolean)

        Dim C As Single


        If AC(IA).libACName = "A380" Then

            NBelly = NBelly + 1S
            I = libNAC + NBelly
            ReDim Preserve AC(I)
            AC(I).libACName$ = AC(IA).libACName$ & BellyExt
            AC(I).libGL = AC(IA).libGL
            AC(I).libMGpcnt = 0.95 * 6 / 10 * 0.5
            AC(I).libCP = 217.6!

            AC(I).libGear = "Z" '"Z" 3
            AC(I).libIGear = 6
            AC(I).libTS = 73.1 * 2

            AC(I).libNTires = 6
            ReDim AC(I).libTX(AC(I).libNTires)
            ReDim AC(I).libTY(AC(I).libNTires)

            'AC(I).libTX(1) = -60.2 / 2 : AC(I).libTY(1) = 0.0# - 66.9
            'AC(I).libTX(2) = 60.2 / 2 : AC(I).libTY(2) = 0.0# - 66.9
            'AC(I).libTX(3) = -61 / 2 : AC(I).libTY(3) = 66.9 - 66.9
            'AC(I).libTX(4) = 61 / 2 : AC(I).libTY(4) = 66.9 - 66.9
            'AC(I).libTX(5) = -60.2 / 2 : AC(I).libTY(5) = 66.9 * 2 - 66.9
            'AC(I).libTX(6) = 60.2 / 2 : AC(I).libTY(6) = 66.9 * 2 - 66.9


            AC(I).libTX(1) = -60.2 / 2 : AC(I).libTY(1) = 0
            AC(I).libTX(2) = 60.2 / 2 : AC(I).libTY(2) = 0
            AC(I).libTX(3) = -61 / 2 : AC(I).libTY(3) = 66.9
            AC(I).libTX(4) = 61 / 2 : AC(I).libTY(4) = 66.9
            AC(I).libTX(5) = 60.2 / 2 : AC(I).libTY(5) = 66.9 * 2
            AC(I).libTX(6) = -60.2 / 2 : AC(I).libTY(6) = 66.9 * 2

            AC(I).libNEVPTS = 6
            ReDim AC(I).libEVPTX(AC(I).libNEVPTS)
            ReDim AC(I).libEVPTY(AC(I).libNEVPTS)

            'Temp = AC(IA).libTT
            Temp = 61
            AC(I).libEVPTX(1) = CSng(Temp * CSng(0.5)) : AC(I).libEVPTY(1) = 0.0! + 66.9
            AC(I).libEVPTX(2) = CSng(Temp * 0.4) : AC(I).libEVPTY(2) = 0.0! + 66.9
            AC(I).libEVPTX(3) = CSng(Temp * 0.3) : AC(I).libEVPTY(3) = 0.0! + 66.9
            AC(I).libEVPTX(4) = CSng(Temp * 0.2) : AC(I).libEVPTY(4) = 0.0! + 66.9
            AC(I).libEVPTX(5) = CSng(Temp * 0.1) : AC(I).libEVPTY(5) = 0.0! + 66.9
            AC(I).libEVPTX(6) = 0.0! : AC(I).libEVPTY(6) = 0.0! + 66.9


            ''Continue For
            ''GoTo DefineBody
            'cont1 = True


        ElseIf AC(IA).libACName = "A380e" Then

            NBelly = NBelly + 1S
            I = libNAC + NBelly
            ReDim Preserve AC(I)
            AC(I).libACName$ = AC(IA).libACName$ & BellyExt
            AC(I).libGL = AC(IA).libGL
            AC(I).libMGpcnt = 0.95 * 6 / 10 * 0.5
            AC(I).libCP = 217.6!

            AC(I).libGear = "Z" '"Z" 3
            AC(I).libIGear = 6
            AC(I).libTS = 73.1 * 2

            AC(I).libNTires = 6
            ReDim AC(I).libTX(AC(I).libNTires)
            ReDim AC(I).libTY(AC(I).libNTires)

            AC(I).libTX(1) = -60.2 / 2 : AC(I).libTY(1) = 0.0#
            AC(I).libTX(2) = 60.2 / 2 : AC(I).libTY(2) = 0.0#
            AC(I).libTX(3) = -61 / 2 : AC(I).libTY(3) = 66.9
            AC(I).libTX(4) = 61 / 2 : AC(I).libTY(4) = 66.9
            AC(I).libTX(5) = -60.2 / 2 : AC(I).libTY(5) = 66.9 * 2
            AC(I).libTX(6) = 60.2 / 2 : AC(I).libTY(6) = 66.9 * 2


            'AC(I).libTX(1) = -60.2 / 2 : AC(I).libTY(1) = 66.9 * 2 - 66.9
            'AC(I).libTX(2) = -61 / 2 : AC(I).libTY(2) = 66.9 - 66.9
            'AC(I).libTX(3) = -60.2 / 2 : AC(I).libTY(3) = 0.0# - 66.9
            'AC(I).libTX(4) = 60.2 / 2 : AC(I).libTY(4) = 0.0# - 66.9
            'AC(I).libTX(5) = 61 / 2 : AC(I).libTY(5) = 66.9 - 66.9
            'AC(I).libTX(6) = 60.2 / 2 : AC(I).libTY(6) = 66.9 * 2 - 66.9

            AC(I).libNEVPTS = 6
            ReDim AC(I).libEVPTX(AC(I).libNEVPTS)
            ReDim AC(I).libEVPTY(AC(I).libNEVPTS)

            'Temp = AC(IA).libTT

            Temp = 61
            AC(I).libEVPTX(1) = CSng(Temp * CSng(0.5)) : AC(I).libEVPTY(1) = 0.0! + 66.9
            AC(I).libEVPTX(2) = CSng(Temp * 0.4) : AC(I).libEVPTY(2) = 0.0! + 66.9
            AC(I).libEVPTX(3) = CSng(Temp * 0.3) : AC(I).libEVPTY(3) = 0.0! + 66.9
            AC(I).libEVPTX(4) = CSng(Temp * 0.2) : AC(I).libEVPTY(4) = 0.0! + 66.9
            AC(I).libEVPTX(5) = CSng(Temp * 0.1) : AC(I).libEVPTY(5) = 0.0! + 66.9
            AC(I).libEVPTX(6) = 0.0! : AC(I).libEVPTY(6) = 0.0! + 66.9

            'Continue For
            cont1 = True

        ElseIf AC(IA).libACName = "B747-100 SF" _
         Or AC(IA).libACName = "B747-200B Combi Mixed" _
         Or AC(IA).libACName = "B747-300 Combi Mixed" _
         Or AC(IA).libACName = "B747-400" _
         Or AC(IA).libACName = "B747-400ER" _
         Or AC(IA).libACName = "B747-8" _
         Or AC(IA).libACName = "B747-8F" _
         Or AC(IA).libACName = "B747-SP" Then


            NBelly = NBelly + 1S
            I = libNAC + NBelly
            ReDim Preserve AC(I)
            AC(I).libACName$ = AC(IA).libACName$ & BellyExt
            AC(I).libGL = AC(IA).libGL
            AC(I).libMGpcnt = 0.95 * 4 / 8 * 0.5



            If AC(IA).libACName = "B747-100 SF" Then
                AC(I).libCP = 232.0!
            ElseIf AC(IA).libACName = "B747-200B Combi Mixed" Then
                AC(I).libCP = 190.0!
            ElseIf AC(IA).libACName = "B747-300 Combi Mixed" Then
                AC(I).libCP = 190.0!
            ElseIf AC(IA).libACName = "B747-400" Then
                AC(I).libCP = 200.0!
            ElseIf AC(IA).libACName = "B747-400ER" Then
                AC(I).libCP = 230.0!
            ElseIf AC(IA).libACName = "B747-8" Then
                AC(I).libCP = 221.0!
            ElseIf AC(IA).libACName = "B747-8F" Then
                AC(I).libCP = 221.0!
            ElseIf AC(IA).libACName = "B747-SP" Then
                AC(I).libCP = 205.0!
            End If


            AC(I).libGear = "F" '"Z" 3
            AC(I).libIGear = 3
            AC(I).libTS = 53 * 2
            AC(I).libTT = 44 : AC(I).libB = 58


            If AC(IA).libACName = "B747-8n" Or AC(IA).libACName = "B747-8Fn" Then
                AC(I).libTT = 46.8 : AC(I).libB = 56.5
            End If

            '=====================================================================

            AC(I).libNTires = 4
            ReDim AC(I).libTX(AC(I).libNTires)
            ReDim AC(I).libTY(AC(I).libNTires)

            'AC(I).libTX(1) = -(AC(I).libTS / 2 + AC(I).libTT)
            'AC(I).libTX(2) = AC(I).libTX(1)
            'AC(I).libTX(3) = -(AC(I).libTS / 2)
            'AC(I).libTX(4) = AC(I).libTX(3)

            AC(I).libTX(1) = -(AC(I).libTT / 2)
            AC(I).libTX(2) = AC(I).libTX(1)
            AC(I).libTX(3) = (AC(I).libTT / 2)
            AC(I).libTX(4) = AC(I).libTX(3)



            AC(I).libTY(1) = AC(I).libB
            AC(I).libTY(2) = 0.0#
            AC(I).libTY(3) = 0.0#
            AC(I).libTY(4) = AC(I).libB

            '=====================================================================

            AC(I).libNEVPTS = 8  ' Do the wing dual-tandem first. Center of gear to wheel 3.
            ReDim AC(I).libEVPTX(AC(I).libNEVPTS)
            ReDim AC(I).libEVPTY(AC(I).libNEVPTS)
            Temp = AC(I).libTT * 0.5!
            C = 0.5604! * AC(I).libTT - 0.2637! * AC(I).libB
            If C < 0.0! Then C = 0.0!

            AC(I).libEVPTX(1) = AC(I).libTX(4)    ' Start at wheel 8 (righear).
            AC(I).libEVPTY(1) = AC(I).libTY(4)
            AC(I).libEVPTX(2) = AC(I).libEVPTX(1) - Temp * 0.2!
            AC(I).libEVPTY(2) = AC(I).libEVPTY(1) - C * 0.2!
            AC(I).libEVPTX(3) = AC(I).libEVPTX(1) - Temp * 0.4!
            AC(I).libEVPTY(3) = AC(I).libEVPTY(1) - C * 0.4!
            AC(I).libEVPTX(4) = AC(I).libEVPTX(1) - Temp * 0.6!
            AC(I).libEVPTY(4) = AC(I).libEVPTY(1) - C * 0.6!
            AC(I).libEVPTX(5) = AC(I).libEVPTX(1) - Temp * 0.8!
            AC(I).libEVPTY(5) = AC(I).libEVPTY(1) - C * 0.8!
            AC(I).libEVPTX(6) = AC(I).libEVPTX(1) - Temp
            AC(I).libEVPTY(6) = AC(I).libEVPTY(1) - C
            AC(I).libEVPTX(7) = AC(I).libEVPTX(6)
            AC(I).libEVPTY(7) = AC(I).libEVPTY(6) - (AC(I).libB * 0.5! - C) * 0.5!
            AC(I).libEVPTX(8) = AC(I).libEVPTX(6)
            AC(I).libEVPTY(8) = AC(I).libEVPTY(6) - (AC(I).libB * 0.5! - C) 'Center of gear.


            cont1 = True


        End If





    End Sub



    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

