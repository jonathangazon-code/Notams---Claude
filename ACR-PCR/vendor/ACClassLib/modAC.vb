Option Explicit On 
Option Strict On

Imports System.IO

Module modAC
    Public Const Err_DeviceUnavailable As Integer = 68
    Public Const Err_DiskNotReady As Integer = 71, Err_FileAlreadyExists As Integer = 58
    Public Const Err_TooManyFiles As Integer = 67, Err_RenameAcrossDisks As Integer = 74
    Public Const Err_Path_FileAccessError As Integer = 75, Err_DeviceIO As Integer = 57
    Public Const Err_DiskFull As Integer = 61, Err_BadFileName As Integer = 64
    Public Const Err_BadFileNameOrNumber As Integer = 52, Err_FileNotFound As Integer = 53
    Public Const Err_PathDoesNotExist As Integer = 76, Err_BadFileMode As Integer = 54
    Public Const Err_FileAlreadyOpen As Integer = 55, Err_InputPastEndOfFile As Integer = 62
    Public Const MB_EXCLAIM As Integer = 48, MB_STOP As Integer = 16

    Public Const kPaTopsi As Double = 0.1450377438
    Public Const cmToin As Double = 0.3937008
    Public Const kgTolb As Double = 2.2046225


    Public Const ExternalAircraftFileName As String = "LEDFAAacLibrary"
    Public ACDATPath As String

    Public ComputeAircraftCDF, AircraftLifeCDFvsDesignLife As Boolean
    Public WriteAircraftLibrary, WriteAircraftLibraryWithTracksandEvals As Boolean

    Sub GeneralAviation(ByRef AC() As clsAC.AircraftCharacteristics, ByRef IA As Short)

        ' General aviation aircraft. Called in Sub IinitAClib. GFH 06-23-04.


        AC(IA).libACName = "Single Wheel 2"
        AC(IA).libGL = 2000.0! : AC(IA).libMGpcnt = 1
        AC(IA).libMGpcntPCN = 1.0! 'PCN Calculations 
        AC(IA).libCP = 30.0!
        AC(IA).libGear = "A"
        AC(IA).libIGear = 1
        AC(IA).libTT = 0 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Single Wheel 5"
        AC(IA).libGL = 5000.0! : AC(IA).libMGpcnt = 1
        AC(IA).libMGpcntPCN = 1.0! 'PCN Calculations 
        AC(IA).libCP = 45.0!
        AC(IA).libGear = "A"
        AC(IA).libIGear = 1
        AC(IA).libTT = 0 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Single Wheel 10"
        AC(IA).libGL = 10000.0! : AC(IA).libMGpcnt = 1
        AC(IA).libMGpcntPCN = 1.0! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "A"
        AC(IA).libIGear = 1
        AC(IA).libTT = 0 : AC(IA).libTS = 0
        IA += 1S

        '----------------------------------------------------------------
        AC(IA).libACName = "S-3"
        AC(IA).libGL = 3000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S


        AC(IA).libACName = "S-5"
        AC(IA).libGL = 5000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "S-10"
        AC(IA).libGL = 10000.0! : AC(IA).libMGpcnt = 0.475 '1.41.0003
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S



        '----------------------------------------------------------------

        AC(IA).libACName = "S-12.5"
        AC(IA).libGL = 12500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "S-15"
        AC(IA).libGL = 15000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "S-20"
        AC(IA).libGL = 20000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 75.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "S-25"
        AC(IA).libGL = 25000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 100.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S



        AC(IA).libACName = "S-30 HTP"
        AC(IA).libGL = 30000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 125.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S


        AC(IA).libACName = "S-35 HTP"
        AC(IA).libGL = 35000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 125.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S


        AC(IA).libACName = "S-40 HTP"
        AC(IA).libGL = 40000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 125.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 150.0 : AC(IA).libTS = 0
        IA += 1S


        '----------------------------------------------------------------
        '----------------------------------------------------------------

        AC(IA).libACName = "D-15"
        AC(IA).libGL = 15000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 55.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 11 : AC(IA).libTS = 139
        IA += 1S

        AC(IA).libACName = "D-20"
        AC(IA).libGL = 20000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 65.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 12 : AC(IA).libTS = 138
        IA += 1S

        AC(IA).libACName = "D-25"
        AC(IA).libGL = 25000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 75.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 15 : AC(IA).libTS = 135
        IA += 1S

        AC(IA).libACName = "D-30"
        AC(IA).libGL = 30000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 85.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 15 : AC(IA).libTS = 135
        IA += 1S

        AC(IA).libACName = "D-35"
        AC(IA).libGL = 35000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 90.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 15 : AC(IA).libTS = 135
        IA += 1S

        AC(IA).libACName = "D-40"
        AC(IA).libGL = 40000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 90.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 15 : AC(IA).libTS = 135
        IA += 1S



        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        AC(IA).libACName = "Aztec-D"
        AC(IA).libGL = 5200.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 46.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 136 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Baron-E-55"
        AC(IA).libGL = 5424.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 56.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 96 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "BeechJet-400"
        AC(IA).libGL = 15500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 90.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 112 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "BeechJet-400A"
        AC(IA).libGL = 16300.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 105.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 112 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Bonanza-F-33A"
        AC(IA).libGL = 3412.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 40.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 115 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Canadair-CL-215"
        AC(IA).libGL = 33000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 177.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 207 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Centurion-210"
        AC(IA).libGL = 4100.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 102 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Challenger-CL-604"
        AC(IA).libGL = 48200.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 145.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 17.4 : AC(IA).libTS = 107.6
        IA += 1S

        AC(IA).libACName = "Chancellor-414"
        AC(IA).libGL = 6200.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 62.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 175.5 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Chk.Arrow-PA-28-200"
        AC(IA).libGL = 2500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 126 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Chk.Six-PA-32"
        AC(IA).libGL = 3400.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 127 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Citation-525"
        AC(IA).libGL = 10500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 98.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 151 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Citation-550B"
        AC(IA).libGL = 15000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 130.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 157 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Citation-V"
        AC(IA).libGL = 16500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 130.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 151 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Citation-VI/VII"
        AC(IA).libGL = 23200.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 168.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 9.75 : AC(IA).libTS = 103.26
        IA += 1S

        AC(IA).libACName = "Citation-X"
        AC(IA).libGL = 36000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 189.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 12.25 : AC(IA).libTS = 114.76
        IA += 1S

        AC(IA).libACName = "Conquest-441"
        AC(IA).libGL = 9925.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 95.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 168 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "DC-3"
        AC(IA).libGL = 26900.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 45.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 222 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Falcon-50"
        AC(IA).libGL = 38800.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 208.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 14 : AC(IA).libTS = 143
        IA += 1S

        AC(IA).libACName = "Falcon-900"
        AC(IA).libGL = 45500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 145.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 14 : AC(IA).libTS = 161
        IA += 1S

        AC(IA).libACName = "Falcon-2000"
        AC(IA).libGL = 35000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 197.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 14 : AC(IA).libTS = 160
        IA += 1S

        AC(IA).libACName = "Fokker-F-28-1000"
        AC(IA).libGL = 66500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 96.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 22.8 : AC(IA).libTS = 175.6
        IA += 1S

        AC(IA).libACName = "Fokker-F-28-2000"
        AC(IA).libGL = 65000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 97.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 22.8 : AC(IA).libTS = 175.6
        IA += 1S

        AC(IA).libACName = "Fokker-F-28-4000"
        AC(IA).libGL = 73000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 105.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 22.8 : AC(IA).libTS = 175.6
        IA += 1S

        AC(IA).libACName = "GrnCaravan-CE-208B"
        AC(IA).libGL = 8750.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 75.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 164 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Gulfstream-G-II"
        AC(IA).libGL = 66000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 160.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 15.75 : AC(IA).libTS = 148.26
        IA += 1S

        AC(IA).libACName = "Gulfstream-G-III"
        AC(IA).libGL = 70200.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 175.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 15.75 : AC(IA).libTS = 148.26
        IA += 1S

        AC(IA).libACName = "Gulfstream-G-IV"
        AC(IA).libGL = 75000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 185.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 15.75 : AC(IA).libTS = 148.26
        IA += 1S

        AC(IA).libACName = "Gulfstream-G-V"
        AC(IA).libGL = 90900.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 188.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 18.5 : AC(IA).libTS = 153.5
        IA += 1S

        AC(IA).libACName = "Hawker-800"
        AC(IA).libGL = 27520.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 135.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 13.3 : AC(IA).libTS = 96.7
        IA += 1S

        AC(IA).libACName = "Hawker-800XP"
        AC(IA).libGL = 28120.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 135.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 13.3 : AC(IA).libTS = 96.7
        IA += 1S

        AC(IA).libACName = "KingAir-B-100"
        AC(IA).libGL = 11500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 52.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 11.58 : AC(IA).libTS = 180
        IA += 1S

        AC(IA).libACName = "KingAir-C-90"
        AC(IA).libGL = 9710.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 58.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 153 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Learjet-35A/65A"
        AC(IA).libGL = 18000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 171.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 12 : AC(IA).libTS = 87
        IA += 1S

        AC(IA).libACName = "Learjet-55"
        AC(IA).libGL = 21500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 201.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 12 : AC(IA).libTS = 87
        IA += 1S

        AC(IA).libACName = "Malibu-PA-46-350P"
        AC(IA).libGL = 4118.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 55.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 146 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Navajo-C"
        AC(IA).libGL = 6536.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 66.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 165 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "RegionalJet-200"
        AC(IA).libGL = 47450.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 177.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 11.56 : AC(IA).libTS = 113.4
        IA += 1S

        AC(IA).libACName = "RegionalJet-700"
        AC(IA).libGL = 72500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 183.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 21.6 : AC(IA).libTS = 140.4
        IA += 1S

        AC(IA).libACName = "Saab 340B" 'ik 2017.01.12
        AC(IA).libGL = 29000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 55.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 14.5! : AC(IA).libTS = 249.5!
        IA += CShort(1)

        AC(IA).libACName = "Sabreliner-40"
        AC(IA).libGL = 19035.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 185.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 86.52 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Sabreliner-60"
        AC(IA).libGL = 20372.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 214.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 86.52 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Sabreliner-65"
        AC(IA).libGL = 24000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 159.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 86.52 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Sabreliner-80"
        AC(IA).libGL = 23500.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 185.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 11.5 : AC(IA).libTS = 77
        IA += 1S

        AC(IA).libACName = "Sarat.PA-32R-301"
        AC(IA).libGL = 3616.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 38.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 133 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Seneca-II"
        AC(IA).libGL = 4570.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 55.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 133 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Shorts-330-200"
        AC(IA).libGL = 22900.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 79.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 167 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Shorts-360"
        AC(IA).libGL = 27200.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 78.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 167 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Skyhawk-172"
        AC(IA).libGL = 2558.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 86 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Skylane-1-82"
        AC(IA).libGL = 3110.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 50.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 96 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "Stationair-206"
        AC(IA).libGL = 3612.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 52.0!
        AC(IA).libGear = "B"
        AC(IA).libIGear = 2
        AC(IA).libTT = 98 : AC(IA).libTS = 0
        IA += 1S

        AC(IA).libACName = "SuperKingAir-300"
        AC(IA).libGL = 14100.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 92.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 11.58 : AC(IA).libTS = 194.42
        IA += 1S

        AC(IA).libACName = "SuperKingAir-350"
        AC(IA).libGL = 15100.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 92.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 11.58 : AC(IA).libTS = 194.42
        IA += 1S

        AC(IA).libACName = "SuperKingAir-B200"
        AC(IA).libGL = 12590.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libMGpcntPCN = 0.475! 'PCN Calculations 
        AC(IA).libCP = 98.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 11.58 : AC(IA).libTS = 194.42
        IA += 1S

    End Sub


    Sub ReadExternalFile(ByRef AC() As clsAC.AircraftCharacteristics, ByRef IA As Short)

        Dim J, I, IErr As Short
        Dim S, SS As String
        Dim FileName, FileFormat As String
        Dim EFNo As Short
        On Error GoTo RExternalFileError


        Dim NL, NL2 As String, Ret As Long, MaxLibAC As Integer
        NL = Environment.NewLine
        NL2 = NL & NL
        MaxLibAC = 140


        FileName = ACDATPath & ExternalAircraftFileName & ".Ext"
        EFNo = CShort(FreeFile())
        FileOpen(EFNo, FileName, OpenMode.Input, , , 1024)

        If EOF(EFNo) Then
            S = "File " & FileName & " is empty or does not exist." & NL2
            S = S & "No action will be taken now."
            Exit Sub
        End If

        FileFormat = LineInput(EFNo)

        Do
            S = LineInput(EFNo)
            '   Next two lines deleted and variables set True in frmStartup.form.load. GFH 07-06-04.
            'If Trim(LCase(S)) = "compute aircraft cdf" Then ComputeAircraftCDF = True
            'If Trim(LCase(S)) = "aircraft life cdf based on design life" Then AircraftLifeCDFvsDesignLife = True
            If Trim(LCase(S)) = "write aircraft library base data to file" Then WriteAircraftLibrary = True
            If Trim(LCase(S)) = "write aircraft library data to file with tracks and evaluation points" Then
                WriteAircraftLibraryWithTracksandEvals = True
            End If
            If S = "Start Data" Then Exit Do
            If EOF(EFNo) Then
                S = "The Start Data marker does not exist in " & FileName & "." & NL2
                S = S & "Aircraft data cannot be read."
                Exit Sub
            End If
        Loop

        Do
            If IA + 1 > MaxLibAC Then
                ReDim Preserve AC(IA + 1)
                'SS = Format(MaxLibAC, "0")
                'S = "The maximum number of aircraft allowed is " & SS & "." & vbCrLf
                'S = S & "The external library file exceeds this value." & vbCrLf
                'S = S & "No more aircraft will be loaded."
                'Ret = CShort(MsgBoxDQ(S, 0, "Too Many Aircraft"))
                'System.Diagnostics.Debug.WriteLine(MaxLibAC & IA)
                'Exit Do
            End If
            If EOF(EFNo) Then Exit Do
            IA += 1S
            AC(IA).libACName = LineInput(EFNo)
            Input(EFNo, AC(IA).libGL)
            Input(EFNo, AC(IA).libMGpcnt) ' input in percent.
            AC(IA).libMGpcnt = AC(IA).libMGpcnt / 100
            Input(EFNo, AC(IA).libCP)
            Input(EFNo, AC(IA).libGear)
            Input(EFNo, AC(IA).libIGear)
            Input(EFNo, AC(IA).libTT)
            Input(EFNo, AC(IA).libTS)
            Input(EFNo, AC(IA).libTG)
            Input(EFNo, AC(IA).libB)

            Input(EFNo, AC(IA).libNTires)
            ReDim AC(IA).libTX(AC(IA).libNTires)
            ReDim AC(IA).libTY(AC(IA).libNTires)
            For I = 1 To AC(IA).libNTires
                Input(EFNo, AC(IA).libTX(I))
                Input(EFNo, AC(IA).libTY(I))
            Next I

            Input(EFNo, AC(IA).libNTTrack)
            ReDim AC(IA).libXAC(AC(IA).libNTTrack)
            For I = 1 To AC(IA).libNTTrack
                Input(EFNo, AC(IA).libXAC(I))
            Next I

            Input(EFNo, AC(IA).libNEVPTS)
            ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
            ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
            For I = 1 To AC(IA).libNEVPTS
                Input(EFNo, AC(IA).libEVPTX(I))
                Input(EFNo, AC(IA).libEVPTY(I))
            Next I

        Loop
        FileClose(EFNo)


        Exit Sub

RExternalFileError:
        IErr = CShort(Err.Number)
        If IErr = 62 Then
            IA = IA - 1S
            Exit Sub
        End If
        Ret = FileErrors(IErr, FileName)
        If Ret = 0 Then Resume
        If Ret = 3 Then
            S = "Error # " & Format(IErr, "0") & " occurred reading" & vbCrLf
            S = S & "file " & FileName & "." & NL2
            S = S & "It may have been caused by a corrupted file or by" & vbCrLf
            S = S & "a file with incorrect format. Please check the file." & NL2
            S = S & "If the error occurred during startup, try loading" & vbCrLf
            S = S & "another job or create a new job. Repair or delete" & vbCrLf
            S = S & "the bad job file as soon as possible."
            Ret = MsgBoxDQ(S, 0, "File Error")
        End If
        FileClose(EFNo)
        Exit Sub

    End Sub

    Sub Add_NAPTF_CC6(ByRef IA As Short, ByRef AC() As clsAC.AircraftCharacteristics)

        AC(IA).libACName = "CC6 45k" '55,000 lbs per wheel
        AC(IA).libGL = 45000 * 8 / 0.95! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 250.0!
        AC(IA).libGear = "F"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)


        AC(IA).libACName = "CC6 52k" '55,000 lbs per wheel
        AC(IA).libGL = 52000 * 8 / 0.95! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 250.0!
        AC(IA).libGear = "F"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)

        'AC(IA).libACName = "CC6 70k100" '55,000 lbs per wheel
        'AC(IA).libGL = 70000 * 8.0! : AC(IA).libMGpcnt = 0.5
        'AC(IA).libCP = 250.0!
        'AC(IA).libGear = "F"
        'AC(IA).libIGear = 3
        'AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        'AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        'IA += CShort(1)


        AC(IA).libACName = "CC6 70k" '55,000 lbs per wheel
        AC(IA).libGL = 70000 * 8 / 0.95! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 250.0!
        AC(IA).libGear = "F"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)


        AC(IA).libACName = "CC8 20k" '20,000 lbs per wheel
        AC(IA).libGL = 20000 * 8 / 0.95! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 220.0!
        AC(IA).libGear = "F"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)

        AC(IA).libACName = "CC8 20k D" '20,000 lbs per wheel
        AC(IA).libGL = 20000 * 4 / 0.95! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 220.0!
        AC(IA).libGear = "D"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 0.0!
        IA += CShort(1)


        AC(IA).libACName = "CC8 20k 3D" '20,000 lbs per wheel
        AC(IA).libGL = 20000 * 12 / 0.95! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 220.0!
        AC(IA).libGear = "N"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)


    End Sub


    Sub Add_NAPTF_Conf(ByRef IA As Short, ByRef AC() As clsAC.AircraftCharacteristics)

        AC(IA).libACName = "CC3 2D" '55,000 lbs per wheel
        AC(IA).libGL = 463158.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 207.5!
        AC(IA).libGear = "F"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)

        AC(IA).libACName = "CC3 3D" '47,500 lbs per wheel
        AC(IA).libGL = 600000.0! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 207.5!
        AC(IA).libGear = "N"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)


        AC(IA).libACName = "3D_55lbs_243"
        AC(IA).libGL = 55000 * 12 / 0.95 : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 243
        AC(IA).libGear = "N"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)


        AC(IA).libACName = "3D_55lbs_220"
        AC(IA).libGL = 55000 * 12 / 0.95 : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 220
        AC(IA).libGear = "N"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 400.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)


        AC(IA).libACName = "CC5-6-50lbs"
        AC(IA).libGL = 50000 * 6 / 0.475! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 243
        AC(IA).libGear = "N"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 440.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)


        AC(IA).libACName = "CC5-6-58lbs"
        AC(IA).libGL = 58000 * 6 / 0.475! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 243
        AC(IA).libGear = "N"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 440.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)

        AC(IA).libACName = "CC5-6-65lbs"
        AC(IA).libGL = 65000 * 6 / 0.475! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 243
        AC(IA).libGear = "N"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 440.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)

        AC(IA).libACName = "CC5-6-70lbs"
        AC(IA).libGL = 70000 * 6 / 0.475! : AC(IA).libMGpcnt = 0.475
        AC(IA).libCP = 243
        AC(IA).libGear = "N"
        AC(IA).libIGear = 3
        AC(IA).libTT = 54.0! : AC(IA).libTS = 440.0!
        AC(IA).libTG = 0.0! : AC(IA).libB = 57.0!
        IA += CShort(1)



    End Sub


    Sub CC510Wheels(ByRef IA As Short, ByRef AC() As clsAC.AircraftCharacteristics)

        ' Sets parameters for Airbus multiple gear aircraft.
        ' A380-800, A380-800F, and A340-600, started 02-17-03 by GFH.
        ' IA is incremented before the call. Therefore, last statement is to increment IA.

        Dim Temp As Single, C As Single
        Dim IEVPTX As Short
        Dim IX, IY, NX, NY As Integer
        Dim StartX, StartY As Single
        Dim DeltaX, DeltaY As Single




        '-------------50,000-----------------
        '10Wheel Coordinates.xlsx

        AC(IA).libACName = "CC5-10-50lbs"
        AC(IA).libGL = 50000 * 20 / 0.95 : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 243.0!
        AC(IA).libGear = "WFBN"  ' Wing = dual-tandem, Body = dual-tridem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 53.1!      ' Wing dual spacing.
        AC(IA).libTS = 84.95!
        AC(IA).libTG = 207.2
        AC(IA).libB = 66.9!       ' Tandem spacing, same for wing and body.
        AC(IA).libNTires = 20

        ReDim AC(IA).libTX(AC(IA).libNTires)
        ReDim AC(IA).libTY(AC(IA).libNTires)

        AC(IA).libTX(1) = -268 : AC(IA).libTY(1) = 57
        AC(IA).libTX(2) = -268 : AC(IA).libTY(2) = 0
        AC(IA).libTX(3) = -214 : AC(IA).libTY(3) = 0
        AC(IA).libTX(4) = -214 : AC(IA).libTY(4) = 57
        AC(IA).libTX(5) = -154 : AC(IA).libTY(5) = 114
        AC(IA).libTX(6) = -154 : AC(IA).libTY(6) = 57
        AC(IA).libTX(7) = -154 : AC(IA).libTY(7) = 0
        AC(IA).libTX(8) = -100 : AC(IA).libTY(8) = 0
        AC(IA).libTX(9) = -100 : AC(IA).libTY(9) = 57
        AC(IA).libTX(10) = -100 : AC(IA).libTY(10) = 114
        AC(IA).libTX(11) = 100 : AC(IA).libTY(11) = 114
        AC(IA).libTX(12) = 100 : AC(IA).libTY(12) = 57
        AC(IA).libTX(13) = 100 : AC(IA).libTY(13) = 0
        AC(IA).libTX(14) = 154 : AC(IA).libTY(14) = 0
        AC(IA).libTX(15) = 154 : AC(IA).libTY(15) = 57
        AC(IA).libTX(16) = 154 : AC(IA).libTY(16) = 114
        AC(IA).libTX(17) = 214 : AC(IA).libTY(17) = 57
        AC(IA).libTX(18) = 214 : AC(IA).libTY(18) = 0
        AC(IA).libTX(19) = 268 : AC(IA).libTY(19) = 0
        AC(IA).libTX(20) = 268 : AC(IA).libTY(20) = 57



        AC(IA).libNTTrack = 8
        ReDim AC(IA).libXAC(AC(IA).libNTTrack)
        AC(IA).libXAC(1) = AC(IA).libTX(1)
        AC(IA).libXAC(2) = AC(IA).libTX(3)
        AC(IA).libXAC(3) = AC(IA).libTX(5)
        AC(IA).libXAC(4) = AC(IA).libTX(8)
        AC(IA).libXAC(5) = AC(IA).libTX(13)
        AC(IA).libXAC(6) = AC(IA).libTX(16)
        AC(IA).libXAC(7) = AC(IA).libTX(18)
        AC(IA).libXAC(8) = AC(IA).libTX(20)
        AC(IA).libNEVPTS = 8  ' Do the standard dual-tandem first. Center of gear to wheel 3.
        ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
        Temp = AC(IA).libTT * 0.5!
        C = 0.5604! * AC(IA).libTT - 0.2637! * AC(IA).libB ' TT and B are set for the wing gear.
        If C < 0.0! Then C = 0.0!
        AC(IA).libEVPTX(1) = AC(IA).libTX(3)    ' Start at wheel 3 (right rear on the left wing gear).
        AC(IA).libEVPTY(1) = AC(IA).libTY(3)
        AC(IA).libEVPTX(2) = AC(IA).libEVPTX(1) - Temp * 0.2!
        AC(IA).libEVPTY(2) = AC(IA).libEVPTY(1) + C * 0.2!
        AC(IA).libEVPTX(3) = AC(IA).libEVPTX(1) - Temp * 0.4!
        AC(IA).libEVPTY(3) = AC(IA).libEVPTY(1) + C * 0.4!
        AC(IA).libEVPTX(4) = AC(IA).libEVPTX(1) - Temp * 0.6!
        AC(IA).libEVPTY(4) = AC(IA).libEVPTY(1) + C * 0.6!
        AC(IA).libEVPTX(5) = AC(IA).libEVPTX(1) - Temp * 0.8!
        AC(IA).libEVPTY(5) = AC(IA).libEVPTY(1) + C * 0.8!
        AC(IA).libEVPTX(6) = AC(IA).libEVPTX(1) - Temp
        AC(IA).libEVPTY(6) = AC(IA).libEVPTY(1) + C
        AC(IA).libEVPTX(7) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(7) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C) * 0.5!
        AC(IA).libEVPTX(8) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(8) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C)   ' Center of gear.
        AC(IA).libNEVPTS = AC(IA).libNEVPTS + 8S  ' Do the standard tridem but with 8 points
        ReDim Preserve AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim Preserve AC(IA).libEVPTY(AC(IA).libNEVPTS)
        ' to be compatible with DT multiple gear.
        Temp = AC(IA).libTX(9) - AC(IA).libTX(6)   ' Center dual spacing slightly larger than TT.
        Temp = Temp / 2 / 7  ' For 8 points from the center to the inboard center wheel.
        AC(IA).libEVPTX(9) = AC(IA).libTX(9) : AC(IA).libEVPTY(9) = AC(IA).libTY(9)
        AC(IA).libEVPTX(10) = AC(IA).libEVPTX(9) - Temp : AC(IA).libEVPTY(10) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(11) = AC(IA).libEVPTX(10) - Temp : AC(IA).libEVPTY(11) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(12) = AC(IA).libEVPTX(11) - Temp : AC(IA).libEVPTY(12) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(13) = AC(IA).libEVPTX(12) - Temp : AC(IA).libEVPTY(13) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(14) = AC(IA).libEVPTX(13) - Temp : AC(IA).libEVPTY(14) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(15) = AC(IA).libEVPTX(14) - Temp : AC(IA).libEVPTY(15) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(16) = AC(IA).libEVPTX(15) - Temp : AC(IA).libEVPTY(16) = AC(IA).libEVPTY(9)
        IEVPTX = 16

        For IX = 1 To -16
            'Debug.Print IX; Format(libEVPTX( IX), "000.00"); "  "; Format(libEVPTY( IX), "000.00")
        Next IX

        ' Now put in a full rectangular grid from wheel 1 (left front wheel in left wing gear)
        ' to the origin (CL of fuselage on line through rear wheels of the body gear).
        NX = 4  ' Total number of lateral points = NX + 1. Force an even number of points
        If (NX + 1) Mod 2 <> 0 Then NX = NX + 1 ' laterally so that grid can be split evenly.
        NY = 4  ' Total number of longitudinal points = NY + 1.
        DeltaX = AC(IA).libTX(1) / NX
        DeltaY = AC(IA).libTY(5) / NY
        StartX = -DeltaX
        StartY = 0
        For IX = 1 To NX + 1
            IEVPTX = IEVPTX + 1S
            ReDim Preserve AC(IA).libEVPTY(IEVPTX)
            AC(IA).libEVPTY(IEVPTX) = StartY
            StartX = StartX + DeltaX
            ReDim Preserve AC(IA).libEVPTX(IEVPTX)
            AC(IA).libEVPTX(IEVPTX) = StartX
            Temp = AC(IA).libEVPTX(IEVPTX)
            '    Debug.Print IEVPTX; IY; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            For IY = 1 To NY
                IEVPTX = IEVPTX + 1S
                ReDim Preserve AC(IA).libEVPTY(IEVPTX)
                ReDim Preserve AC(IA).libEVPTX(IEVPTX)
                AC(IA).libEVPTY(IEVPTX) = AC(IA).libEVPTY(IEVPTX - 1) + DeltaY
                AC(IA).libEVPTX(IEVPTX) = StartX
                '      Debug.Print IEVPTX; IY; IX; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            Next IY
        Next IX
        AC(IA).libNEVPTS = IEVPTX

        Dim I As Integer
        For I = 22 To 26
            AC(IA).libEVPTX(I) = AC(IA).libTX(10) / 2
            AC(IA).libEVPTX(I + 5) = AC(IA).libTX(10)
            AC(IA).libEVPTX(I + 10) = AC(IA).libTX(5)
            AC(IA).libEVPTX(I + 15) = AC(IA).libTX(3)
            AC(IA).libEVPTX(I + 20) = AC(IA).libTX(1)
        Next I

        IA += 1S




        '-------------58,000-----------------

        AC(IA).libACName = "CC5-10-58lbs"
        AC(IA).libGL = 58000 * 20 / 0.95 : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 243.0!
        AC(IA).libGear = "WFBN"  ' Wing = dual-tandem, Body = dual-tridem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 53.1!      ' Wing dual spacing.
        AC(IA).libTS = 84.95!
        AC(IA).libTG = 207.2
        AC(IA).libB = 66.9!       ' Tandem spacing, same for wing and body.
        AC(IA).libNTires = 20

        ReDim AC(IA).libTX(AC(IA).libNTires)
        ReDim AC(IA).libTY(AC(IA).libNTires)

        AC(IA).libTX(1) = -268 : AC(IA).libTY(1) = 57
        AC(IA).libTX(2) = -268 : AC(IA).libTY(2) = 0
        AC(IA).libTX(3) = -214 : AC(IA).libTY(3) = 0
        AC(IA).libTX(4) = -214 : AC(IA).libTY(4) = 57
        AC(IA).libTX(5) = -154 : AC(IA).libTY(5) = 114
        AC(IA).libTX(6) = -154 : AC(IA).libTY(6) = 57
        AC(IA).libTX(7) = -154 : AC(IA).libTY(7) = 0
        AC(IA).libTX(8) = -100 : AC(IA).libTY(8) = 0
        AC(IA).libTX(9) = -100 : AC(IA).libTY(9) = 57
        AC(IA).libTX(10) = -100 : AC(IA).libTY(10) = 114
        AC(IA).libTX(11) = 100 : AC(IA).libTY(11) = 114
        AC(IA).libTX(12) = 100 : AC(IA).libTY(12) = 57
        AC(IA).libTX(13) = 100 : AC(IA).libTY(13) = 0
        AC(IA).libTX(14) = 154 : AC(IA).libTY(14) = 0
        AC(IA).libTX(15) = 154 : AC(IA).libTY(15) = 57
        AC(IA).libTX(16) = 154 : AC(IA).libTY(16) = 114
        AC(IA).libTX(17) = 214 : AC(IA).libTY(17) = 57
        AC(IA).libTX(18) = 214 : AC(IA).libTY(18) = 0
        AC(IA).libTX(19) = 268 : AC(IA).libTY(19) = 0
        AC(IA).libTX(20) = 268 : AC(IA).libTY(20) = 57



        AC(IA).libNTTrack = 8
        ReDim AC(IA).libXAC(AC(IA).libNTTrack)
        AC(IA).libXAC(1) = AC(IA).libTX(1)
        AC(IA).libXAC(2) = AC(IA).libTX(3)
        AC(IA).libXAC(3) = AC(IA).libTX(5)
        AC(IA).libXAC(4) = AC(IA).libTX(8)
        AC(IA).libXAC(5) = AC(IA).libTX(13)
        AC(IA).libXAC(6) = AC(IA).libTX(16)
        AC(IA).libXAC(7) = AC(IA).libTX(18)
        AC(IA).libXAC(8) = AC(IA).libTX(20)
        AC(IA).libNEVPTS = 8  ' Do the standard dual-tandem first. Center of gear to wheel 3.
        ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
        Temp = AC(IA).libTT * 0.5!
        C = 0.5604! * AC(IA).libTT - 0.2637! * AC(IA).libB ' TT and B are set for the wing gear.
        If C < 0.0! Then C = 0.0!
        AC(IA).libEVPTX(1) = AC(IA).libTX(3)    ' Start at wheel 3 (right rear on the left wing gear).
        AC(IA).libEVPTY(1) = AC(IA).libTY(3)
        AC(IA).libEVPTX(2) = AC(IA).libEVPTX(1) - Temp * 0.2!
        AC(IA).libEVPTY(2) = AC(IA).libEVPTY(1) + C * 0.2!
        AC(IA).libEVPTX(3) = AC(IA).libEVPTX(1) - Temp * 0.4!
        AC(IA).libEVPTY(3) = AC(IA).libEVPTY(1) + C * 0.4!
        AC(IA).libEVPTX(4) = AC(IA).libEVPTX(1) - Temp * 0.6!
        AC(IA).libEVPTY(4) = AC(IA).libEVPTY(1) + C * 0.6!
        AC(IA).libEVPTX(5) = AC(IA).libEVPTX(1) - Temp * 0.8!
        AC(IA).libEVPTY(5) = AC(IA).libEVPTY(1) + C * 0.8!
        AC(IA).libEVPTX(6) = AC(IA).libEVPTX(1) - Temp
        AC(IA).libEVPTY(6) = AC(IA).libEVPTY(1) + C
        AC(IA).libEVPTX(7) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(7) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C) * 0.5!
        AC(IA).libEVPTX(8) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(8) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C)   ' Center of gear.
        AC(IA).libNEVPTS = AC(IA).libNEVPTS + 8S  ' Do the standard tridem but with 8 points
        ReDim Preserve AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim Preserve AC(IA).libEVPTY(AC(IA).libNEVPTS)
        ' to be compatible with DT multiple gear.
        Temp = AC(IA).libTX(9) - AC(IA).libTX(6)   ' Center dual spacing slightly larger than TT.
        Temp = Temp / 2 / 7  ' For 8 points from the center to the inboard center wheel.
        AC(IA).libEVPTX(9) = AC(IA).libTX(9) : AC(IA).libEVPTY(9) = AC(IA).libTY(9)
        AC(IA).libEVPTX(10) = AC(IA).libEVPTX(9) - Temp : AC(IA).libEVPTY(10) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(11) = AC(IA).libEVPTX(10) - Temp : AC(IA).libEVPTY(11) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(12) = AC(IA).libEVPTX(11) - Temp : AC(IA).libEVPTY(12) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(13) = AC(IA).libEVPTX(12) - Temp : AC(IA).libEVPTY(13) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(14) = AC(IA).libEVPTX(13) - Temp : AC(IA).libEVPTY(14) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(15) = AC(IA).libEVPTX(14) - Temp : AC(IA).libEVPTY(15) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(16) = AC(IA).libEVPTX(15) - Temp : AC(IA).libEVPTY(16) = AC(IA).libEVPTY(9)
        IEVPTX = 16

        For IX = 1 To -16
            'Debug.Print IX; Format(libEVPTX( IX), "000.00"); "  "; Format(libEVPTY( IX), "000.00")
        Next IX

        ' Now put in a full rectangular grid from wheel 1 (left front wheel in left wing gear)
        ' to the origin (CL of fuselage on line through rear wheels of the body gear).
        NX = 4  ' Total number of lateral points = NX + 1. Force an even number of points
        If (NX + 1) Mod 2 <> 0 Then NX = NX + 1 ' laterally so that grid can be split evenly.
        NY = 4  ' Total number of longitudinal points = NY + 1.
        DeltaX = AC(IA).libTX(1) / NX
        DeltaY = AC(IA).libTY(5) / NY
        StartX = -DeltaX
        StartY = 0
        For IX = 1 To NX + 1
            IEVPTX = IEVPTX + 1S
            ReDim Preserve AC(IA).libEVPTY(IEVPTX)
            AC(IA).libEVPTY(IEVPTX) = StartY
            StartX = StartX + DeltaX
            ReDim Preserve AC(IA).libEVPTX(IEVPTX)
            AC(IA).libEVPTX(IEVPTX) = StartX
            Temp = AC(IA).libEVPTX(IEVPTX)
            '    Debug.Print IEVPTX; IY; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            For IY = 1 To NY
                IEVPTX = IEVPTX + 1S
                ReDim Preserve AC(IA).libEVPTY(IEVPTX)
                ReDim Preserve AC(IA).libEVPTX(IEVPTX)
                AC(IA).libEVPTY(IEVPTX) = AC(IA).libEVPTY(IEVPTX - 1) + DeltaY
                AC(IA).libEVPTX(IEVPTX) = StartX
                '      Debug.Print IEVPTX; IY; IX; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            Next IY
        Next IX
        AC(IA).libNEVPTS = IEVPTX


        For I = 22 To 26
            AC(IA).libEVPTX(I) = AC(IA).libTX(10) / 2
            AC(IA).libEVPTX(I + 5) = AC(IA).libTX(10)
            AC(IA).libEVPTX(I + 10) = AC(IA).libTX(5)
            AC(IA).libEVPTX(I + 15) = AC(IA).libTX(3)
            AC(IA).libEVPTX(I + 20) = AC(IA).libTX(1)
        Next I

        IA += 1S




        '-------------65,000-----------------

        AC(IA).libACName = "CC5-10-65lbs"
        AC(IA).libGL = 65000 * 20 / 0.95 : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 243.0!
        AC(IA).libGear = "WFBN"  ' Wing = dual-tandem, Body = dual-tridem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 53.1!      ' Wing dual spacing.
        AC(IA).libTS = 84.95!
        AC(IA).libTG = 207.2
        AC(IA).libB = 66.9!       ' Tandem spacing, same for wing and body.
        AC(IA).libNTires = 20

        ReDim AC(IA).libTX(AC(IA).libNTires)
        ReDim AC(IA).libTY(AC(IA).libNTires)

        AC(IA).libTX(1) = -268 : AC(IA).libTY(1) = 57
        AC(IA).libTX(2) = -268 : AC(IA).libTY(2) = 0
        AC(IA).libTX(3) = -214 : AC(IA).libTY(3) = 0
        AC(IA).libTX(4) = -214 : AC(IA).libTY(4) = 57
        AC(IA).libTX(5) = -154 : AC(IA).libTY(5) = 114
        AC(IA).libTX(6) = -154 : AC(IA).libTY(6) = 57
        AC(IA).libTX(7) = -154 : AC(IA).libTY(7) = 0
        AC(IA).libTX(8) = -100 : AC(IA).libTY(8) = 0
        AC(IA).libTX(9) = -100 : AC(IA).libTY(9) = 57
        AC(IA).libTX(10) = -100 : AC(IA).libTY(10) = 114
        AC(IA).libTX(11) = 100 : AC(IA).libTY(11) = 114
        AC(IA).libTX(12) = 100 : AC(IA).libTY(12) = 57
        AC(IA).libTX(13) = 100 : AC(IA).libTY(13) = 0
        AC(IA).libTX(14) = 154 : AC(IA).libTY(14) = 0
        AC(IA).libTX(15) = 154 : AC(IA).libTY(15) = 57
        AC(IA).libTX(16) = 154 : AC(IA).libTY(16) = 114
        AC(IA).libTX(17) = 214 : AC(IA).libTY(17) = 57
        AC(IA).libTX(18) = 214 : AC(IA).libTY(18) = 0
        AC(IA).libTX(19) = 268 : AC(IA).libTY(19) = 0
        AC(IA).libTX(20) = 268 : AC(IA).libTY(20) = 57



        AC(IA).libNTTrack = 8
        ReDim AC(IA).libXAC(AC(IA).libNTTrack)
        AC(IA).libXAC(1) = AC(IA).libTX(1)
        AC(IA).libXAC(2) = AC(IA).libTX(3)
        AC(IA).libXAC(3) = AC(IA).libTX(5)
        AC(IA).libXAC(4) = AC(IA).libTX(8)
        AC(IA).libXAC(5) = AC(IA).libTX(13)
        AC(IA).libXAC(6) = AC(IA).libTX(16)
        AC(IA).libXAC(7) = AC(IA).libTX(18)
        AC(IA).libXAC(8) = AC(IA).libTX(20)
        AC(IA).libNEVPTS = 8  ' Do the standard dual-tandem first. Center of gear to wheel 3.
        ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
        Temp = AC(IA).libTT * 0.5!
        C = 0.5604! * AC(IA).libTT - 0.2637! * AC(IA).libB ' TT and B are set for the wing gear.
        If C < 0.0! Then C = 0.0!
        AC(IA).libEVPTX(1) = AC(IA).libTX(3)    ' Start at wheel 3 (right rear on the left wing gear).
        AC(IA).libEVPTY(1) = AC(IA).libTY(3)
        AC(IA).libEVPTX(2) = AC(IA).libEVPTX(1) - Temp * 0.2!
        AC(IA).libEVPTY(2) = AC(IA).libEVPTY(1) + C * 0.2!
        AC(IA).libEVPTX(3) = AC(IA).libEVPTX(1) - Temp * 0.4!
        AC(IA).libEVPTY(3) = AC(IA).libEVPTY(1) + C * 0.4!
        AC(IA).libEVPTX(4) = AC(IA).libEVPTX(1) - Temp * 0.6!
        AC(IA).libEVPTY(4) = AC(IA).libEVPTY(1) + C * 0.6!
        AC(IA).libEVPTX(5) = AC(IA).libEVPTX(1) - Temp * 0.8!
        AC(IA).libEVPTY(5) = AC(IA).libEVPTY(1) + C * 0.8!
        AC(IA).libEVPTX(6) = AC(IA).libEVPTX(1) - Temp
        AC(IA).libEVPTY(6) = AC(IA).libEVPTY(1) + C
        AC(IA).libEVPTX(7) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(7) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C) * 0.5!
        AC(IA).libEVPTX(8) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(8) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C)   ' Center of gear.
        AC(IA).libNEVPTS = AC(IA).libNEVPTS + 8S  ' Do the standard tridem but with 8 points
        ReDim Preserve AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim Preserve AC(IA).libEVPTY(AC(IA).libNEVPTS)
        ' to be compatible with DT multiple gear.
        Temp = AC(IA).libTX(9) - AC(IA).libTX(6)   ' Center dual spacing slightly larger than TT.
        Temp = Temp / 2 / 7  ' For 8 points from the center to the inboard center wheel.
        AC(IA).libEVPTX(9) = AC(IA).libTX(9) : AC(IA).libEVPTY(9) = AC(IA).libTY(9)
        AC(IA).libEVPTX(10) = AC(IA).libEVPTX(9) - Temp : AC(IA).libEVPTY(10) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(11) = AC(IA).libEVPTX(10) - Temp : AC(IA).libEVPTY(11) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(12) = AC(IA).libEVPTX(11) - Temp : AC(IA).libEVPTY(12) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(13) = AC(IA).libEVPTX(12) - Temp : AC(IA).libEVPTY(13) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(14) = AC(IA).libEVPTX(13) - Temp : AC(IA).libEVPTY(14) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(15) = AC(IA).libEVPTX(14) - Temp : AC(IA).libEVPTY(15) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(16) = AC(IA).libEVPTX(15) - Temp : AC(IA).libEVPTY(16) = AC(IA).libEVPTY(9)
        IEVPTX = 16

        For IX = 1 To -16
            'Debug.Print IX; Format(libEVPTX( IX), "000.00"); "  "; Format(libEVPTY( IX), "000.00")
        Next IX

        ' Now put in a full rectangular grid from wheel 1 (left front wheel in left wing gear)
        ' to the origin (CL of fuselage on line through rear wheels of the body gear).
        NX = 4  ' Total number of lateral points = NX + 1. Force an even number of points
        If (NX + 1) Mod 2 <> 0 Then NX = NX + 1 ' laterally so that grid can be split evenly.
        NY = 4  ' Total number of longitudinal points = NY + 1.
        DeltaX = AC(IA).libTX(1) / NX
        DeltaY = AC(IA).libTY(5) / NY
        StartX = -DeltaX
        StartY = 0
        For IX = 1 To NX + 1
            IEVPTX = IEVPTX + 1S
            ReDim Preserve AC(IA).libEVPTY(IEVPTX)
            AC(IA).libEVPTY(IEVPTX) = StartY
            StartX = StartX + DeltaX
            ReDim Preserve AC(IA).libEVPTX(IEVPTX)
            AC(IA).libEVPTX(IEVPTX) = StartX
            Temp = AC(IA).libEVPTX(IEVPTX)
            '    Debug.Print IEVPTX; IY; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            For IY = 1 To NY
                IEVPTX = IEVPTX + 1S
                ReDim Preserve AC(IA).libEVPTY(IEVPTX)
                ReDim Preserve AC(IA).libEVPTX(IEVPTX)
                AC(IA).libEVPTY(IEVPTX) = AC(IA).libEVPTY(IEVPTX - 1) + DeltaY
                AC(IA).libEVPTX(IEVPTX) = StartX
                '      Debug.Print IEVPTX; IY; IX; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            Next IY
        Next IX
        AC(IA).libNEVPTS = IEVPTX


        For I = 22 To 26
            AC(IA).libEVPTX(I) = AC(IA).libTX(10) / 2
            AC(IA).libEVPTX(I + 5) = AC(IA).libTX(10)
            AC(IA).libEVPTX(I + 10) = AC(IA).libTX(5)
            AC(IA).libEVPTX(I + 15) = AC(IA).libTX(3)
            AC(IA).libEVPTX(I + 20) = AC(IA).libTX(1)
        Next I

        IA += 1S




        '-------------70,000-----------------


        AC(IA).libACName = "CC5-10-70lbs"
        AC(IA).libGL = 70000 * 20 / 0.95 : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 243.0!
        AC(IA).libGear = "WFBN"  ' Wing = dual-tandem, Body = dual-tridem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 53.1!      ' Wing dual spacing.
        AC(IA).libTS = 84.95!
        AC(IA).libTG = 207.2
        AC(IA).libB = 66.9!       ' Tandem spacing, same for wing and body.
        AC(IA).libNTires = 20

        ReDim AC(IA).libTX(AC(IA).libNTires)
        ReDim AC(IA).libTY(AC(IA).libNTires)

        AC(IA).libTX(1) = -268 : AC(IA).libTY(1) = 57
        AC(IA).libTX(2) = -268 : AC(IA).libTY(2) = 0
        AC(IA).libTX(3) = -214 : AC(IA).libTY(3) = 0
        AC(IA).libTX(4) = -214 : AC(IA).libTY(4) = 57
        AC(IA).libTX(5) = -154 : AC(IA).libTY(5) = 114
        AC(IA).libTX(6) = -154 : AC(IA).libTY(6) = 57
        AC(IA).libTX(7) = -154 : AC(IA).libTY(7) = 0
        AC(IA).libTX(8) = -100 : AC(IA).libTY(8) = 0
        AC(IA).libTX(9) = -100 : AC(IA).libTY(9) = 57
        AC(IA).libTX(10) = -100 : AC(IA).libTY(10) = 114
        AC(IA).libTX(11) = 100 : AC(IA).libTY(11) = 114
        AC(IA).libTX(12) = 100 : AC(IA).libTY(12) = 57
        AC(IA).libTX(13) = 100 : AC(IA).libTY(13) = 0
        AC(IA).libTX(14) = 154 : AC(IA).libTY(14) = 0
        AC(IA).libTX(15) = 154 : AC(IA).libTY(15) = 57
        AC(IA).libTX(16) = 154 : AC(IA).libTY(16) = 114
        AC(IA).libTX(17) = 214 : AC(IA).libTY(17) = 57
        AC(IA).libTX(18) = 214 : AC(IA).libTY(18) = 0
        AC(IA).libTX(19) = 268 : AC(IA).libTY(19) = 0
        AC(IA).libTX(20) = 268 : AC(IA).libTY(20) = 57



        AC(IA).libNTTrack = 8
        ReDim AC(IA).libXAC(AC(IA).libNTTrack)
        AC(IA).libXAC(1) = AC(IA).libTX(1)
        AC(IA).libXAC(2) = AC(IA).libTX(3)
        AC(IA).libXAC(3) = AC(IA).libTX(5)
        AC(IA).libXAC(4) = AC(IA).libTX(8)
        AC(IA).libXAC(5) = AC(IA).libTX(13)
        AC(IA).libXAC(6) = AC(IA).libTX(16)
        AC(IA).libXAC(7) = AC(IA).libTX(18)
        AC(IA).libXAC(8) = AC(IA).libTX(20)
        AC(IA).libNEVPTS = 8  ' Do the standard dual-tandem first. Center of gear to wheel 3.
        ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
        Temp = AC(IA).libTT * 0.5!
        C = 0.5604! * AC(IA).libTT - 0.2637! * AC(IA).libB ' TT and B are set for the wing gear.
        If C < 0.0! Then C = 0.0!
        AC(IA).libEVPTX(1) = AC(IA).libTX(3)    ' Start at wheel 3 (right rear on the left wing gear).
        AC(IA).libEVPTY(1) = AC(IA).libTY(3)
        AC(IA).libEVPTX(2) = AC(IA).libEVPTX(1) - Temp * 0.2!
        AC(IA).libEVPTY(2) = AC(IA).libEVPTY(1) + C * 0.2!
        AC(IA).libEVPTX(3) = AC(IA).libEVPTX(1) - Temp * 0.4!
        AC(IA).libEVPTY(3) = AC(IA).libEVPTY(1) + C * 0.4!
        AC(IA).libEVPTX(4) = AC(IA).libEVPTX(1) - Temp * 0.6!
        AC(IA).libEVPTY(4) = AC(IA).libEVPTY(1) + C * 0.6!
        AC(IA).libEVPTX(5) = AC(IA).libEVPTX(1) - Temp * 0.8!
        AC(IA).libEVPTY(5) = AC(IA).libEVPTY(1) + C * 0.8!
        AC(IA).libEVPTX(6) = AC(IA).libEVPTX(1) - Temp
        AC(IA).libEVPTY(6) = AC(IA).libEVPTY(1) + C
        AC(IA).libEVPTX(7) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(7) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C) * 0.5!
        AC(IA).libEVPTX(8) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(8) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C)   ' Center of gear.
        AC(IA).libNEVPTS = AC(IA).libNEVPTS + 8S  ' Do the standard tridem but with 8 points
        ReDim Preserve AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim Preserve AC(IA).libEVPTY(AC(IA).libNEVPTS)
        ' to be compatible with DT multiple gear.
        Temp = AC(IA).libTX(9) - AC(IA).libTX(6)   ' Center dual spacing slightly larger than TT.
        Temp = Temp / 2 / 7  ' For 8 points from the center to the inboard center wheel.
        AC(IA).libEVPTX(9) = AC(IA).libTX(9) : AC(IA).libEVPTY(9) = AC(IA).libTY(9)
        AC(IA).libEVPTX(10) = AC(IA).libEVPTX(9) - Temp : AC(IA).libEVPTY(10) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(11) = AC(IA).libEVPTX(10) - Temp : AC(IA).libEVPTY(11) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(12) = AC(IA).libEVPTX(11) - Temp : AC(IA).libEVPTY(12) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(13) = AC(IA).libEVPTX(12) - Temp : AC(IA).libEVPTY(13) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(14) = AC(IA).libEVPTX(13) - Temp : AC(IA).libEVPTY(14) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(15) = AC(IA).libEVPTX(14) - Temp : AC(IA).libEVPTY(15) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(16) = AC(IA).libEVPTX(15) - Temp : AC(IA).libEVPTY(16) = AC(IA).libEVPTY(9)
        IEVPTX = 16

        For IX = 1 To -16
            'Debug.Print IX; Format(libEVPTX( IX), "000.00"); "  "; Format(libEVPTY( IX), "000.00")
        Next IX

        ' Now put in a full rectangular grid from wheel 1 (left front wheel in left wing gear)
        ' to the origin (CL of fuselage on line through rear wheels of the body gear).
        NX = 4  ' Total number of lateral points = NX + 1. Force an even number of points
        If (NX + 1) Mod 2 <> 0 Then NX = NX + 1 ' laterally so that grid can be split evenly.
        NY = 4  ' Total number of longitudinal points = NY + 1.
        DeltaX = AC(IA).libTX(1) / NX
        DeltaY = AC(IA).libTY(5) / NY
        StartX = -DeltaX
        StartY = 0
        For IX = 1 To NX + 1
            IEVPTX = IEVPTX + 1S
            ReDim Preserve AC(IA).libEVPTY(IEVPTX)
            AC(IA).libEVPTY(IEVPTX) = StartY
            StartX = StartX + DeltaX
            ReDim Preserve AC(IA).libEVPTX(IEVPTX)
            AC(IA).libEVPTX(IEVPTX) = StartX
            Temp = AC(IA).libEVPTX(IEVPTX)
            '    Debug.Print IEVPTX; IY; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            For IY = 1 To NY
                IEVPTX = IEVPTX + 1S
                ReDim Preserve AC(IA).libEVPTY(IEVPTX)
                ReDim Preserve AC(IA).libEVPTX(IEVPTX)
                AC(IA).libEVPTY(IEVPTX) = AC(IA).libEVPTY(IEVPTX - 1) + DeltaY
                AC(IA).libEVPTX(IEVPTX) = StartX
                '      Debug.Print IEVPTX; IY; IX; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            Next IY
        Next IX
        AC(IA).libNEVPTS = IEVPTX


        For I = 22 To 26
            AC(IA).libEVPTX(I) = AC(IA).libTX(10) / 2
            AC(IA).libEVPTX(I + 5) = AC(IA).libTX(10)
            AC(IA).libEVPTX(I + 10) = AC(IA).libTX(5)
            AC(IA).libEVPTX(I + 15) = AC(IA).libTX(3)
            AC(IA).libEVPTX(I + 20) = AC(IA).libTX(1)
        Next I

        'IA += 1S





    End Sub


    Sub AirbusMultiGear(ByRef IA As Short, ByRef AC() As clsAC.AircraftCharacteristics)

        ' Sets parameters for Airbus multiple gear aircraft.
        ' A380-800, A380-800F, and A340-600, started 02-17-03 by GFH.
        ' IA is incremented before the call. Therefore, last statement is to increment IA.

        Dim Temp As Single, C As Single
        Dim IEVPTX As Short
        Dim IX, IY, NX, NY As Integer
        Dim StartX, StartY As Single
        Dim DeltaX, DeltaY As Single



        'modified Izydor Kawa 2012.09.25
        AC(IA).libACName = "A380"
        AC(IA).libGL = 562000 * kgTolb : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 217.6!
        AC(IA).libGear = "WFBN"  ' Wing = dual-tandem, Body = dual-tridem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 53.1!      ' Wing dual spacing.
        AC(IA).libTS = 84.95!
        AC(IA).libTG = 207.2
        AC(IA).libB = 66.9!       ' Tandem spacing, same for wing and body.
        AC(IA).libNTires = 20

        ReDim AC(IA).libTX(AC(IA).libNTires)
        ReDim AC(IA).libTY(AC(IA).libNTires)

        AC(IA).libTX(1) = -271.75 : AC(IA).libTY(1) = 229.35
        AC(IA).libTX(2) = -271.75 : AC(IA).libTY(2) = 162.45
        AC(IA).libTX(3) = -218.65 : AC(IA).libTY(3) = 162.45
        AC(IA).libTX(4) = -218.65 : AC(IA).libTY(4) = 229.35
        AC(IA).libTX(5) = -133.7 : AC(IA).libTY(5) = 133.8
        AC(IA).libTX(6) = -134.1 : AC(IA).libTY(6) = 66.9
        AC(IA).libTX(7) = -133.7 : AC(IA).libTY(7) = 0.0#
        AC(IA).libTX(8) = -73.5 : AC(IA).libTY(8) = 0.0#
        AC(IA).libTX(9) = -73.1 : AC(IA).libTY(9) = 66.9
        AC(IA).libTX(10) = -73.5 : AC(IA).libTY(10) = 133.8
        AC(IA).libTX(11) = 73.5 : AC(IA).libTY(11) = 133.8
        AC(IA).libTX(12) = 73.1 : AC(IA).libTY(12) = 66.9
        AC(IA).libTX(13) = 73.5 : AC(IA).libTY(13) = 0.0#
        AC(IA).libTX(14) = 133.7 : AC(IA).libTY(14) = 0.0#
        AC(IA).libTX(15) = 134.1 : AC(IA).libTY(15) = 66.9
        AC(IA).libTX(16) = 133.7 : AC(IA).libTY(16) = 133.8
        AC(IA).libTX(17) = 218.65 : AC(IA).libTY(17) = 229.35
        AC(IA).libTX(18) = 218.65 : AC(IA).libTY(18) = 162.45
        AC(IA).libTX(19) = 271.75 : AC(IA).libTY(19) = 162.45
        AC(IA).libTX(20) = 271.75 : AC(IA).libTY(20) = 229.35
        AC(IA).libNTTrack = 8
        ReDim AC(IA).libXAC(AC(IA).libNTTrack)
        AC(IA).libXAC(1) = -271.75
        AC(IA).libXAC(2) = -218.65
        AC(IA).libXAC(3) = -133.7
        AC(IA).libXAC(4) = -73.5
        AC(IA).libXAC(5) = 73.5
        AC(IA).libXAC(6) = 133.7
        AC(IA).libXAC(7) = 218.65
        AC(IA).libXAC(8) = 271.75
        AC(IA).libNEVPTS = 8  ' Do the standard dual-tandem first. Center of gear to wheel 3.
        ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
        Temp = AC(IA).libTT * 0.5!
        C = 0.5604! * AC(IA).libTT - 0.2637! * AC(IA).libB ' TT and B are set for the wing gear.
        If C < 0.0! Then C = 0.0!
        AC(IA).libEVPTX(1) = AC(IA).libTX(3)    ' Start at wheel 3 (right rear on the left wing gear).
        AC(IA).libEVPTY(1) = AC(IA).libTY(3)
        AC(IA).libEVPTX(2) = AC(IA).libEVPTX(1) - Temp * 0.2!
        AC(IA).libEVPTY(2) = AC(IA).libEVPTY(1) + C * 0.2!
        AC(IA).libEVPTX(3) = AC(IA).libEVPTX(1) - Temp * 0.4!
        AC(IA).libEVPTY(3) = AC(IA).libEVPTY(1) + C * 0.4!
        AC(IA).libEVPTX(4) = AC(IA).libEVPTX(1) - Temp * 0.6!
        AC(IA).libEVPTY(4) = AC(IA).libEVPTY(1) + C * 0.6!
        AC(IA).libEVPTX(5) = AC(IA).libEVPTX(1) - Temp * 0.8!
        AC(IA).libEVPTY(5) = AC(IA).libEVPTY(1) + C * 0.8!
        AC(IA).libEVPTX(6) = AC(IA).libEVPTX(1) - Temp
        AC(IA).libEVPTY(6) = AC(IA).libEVPTY(1) + C
        AC(IA).libEVPTX(7) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(7) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C) * 0.5!
        AC(IA).libEVPTX(8) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(8) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C)   ' Center of gear.
        AC(IA).libNEVPTS = AC(IA).libNEVPTS + 8S  ' Do the standard tridem but with 8 points
        ReDim Preserve AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim Preserve AC(IA).libEVPTY(AC(IA).libNEVPTS)
        ' to be compatible with DT multiple gear.
        Temp = AC(IA).libTX(9) - AC(IA).libTX(6)   ' Center dual spacing slightly larger than TT.
        Temp = Temp / 2 / 7  ' For 8 points from the center to the inboard center wheel.
        AC(IA).libEVPTX(9) = AC(IA).libTX(9) : AC(IA).libEVPTY(9) = AC(IA).libTY(9)
        AC(IA).libEVPTX(10) = AC(IA).libEVPTX(9) - Temp : AC(IA).libEVPTY(10) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(11) = AC(IA).libEVPTX(10) - Temp : AC(IA).libEVPTY(11) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(12) = AC(IA).libEVPTX(11) - Temp : AC(IA).libEVPTY(12) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(13) = AC(IA).libEVPTX(12) - Temp : AC(IA).libEVPTY(13) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(14) = AC(IA).libEVPTX(13) - Temp : AC(IA).libEVPTY(14) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(15) = AC(IA).libEVPTX(14) - Temp : AC(IA).libEVPTY(15) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(16) = AC(IA).libEVPTX(15) - Temp : AC(IA).libEVPTY(16) = AC(IA).libEVPTY(9)
        IEVPTX = 16

        For IX = 1 To -16
            'Debug.Print IX; Format(libEVPTX( IX), "000.00"); "  "; Format(libEVPTY( IX), "000.00")
        Next IX

        ' Now put in a full rectangular grid from wheel 1 (left front wheel in left wing gear)
        ' to the origin (CL of fuselage on line through rear wheels of the body gear).
        NX = 4  ' Total number of lateral points = NX + 1. Force an even number of points
        If (NX + 1) Mod 2 <> 0 Then NX = NX + 1 ' laterally so that grid can be split evenly.
        NY = 4  ' Total number of longitudinal points = NY + 1.
        DeltaX = AC(IA).libTX(1) / NX
        DeltaY = AC(IA).libTY(1) / NY
        StartX = -DeltaX
        StartY = 0
        For IX = 1 To NX + 1
            IEVPTX = IEVPTX + 1S
            ReDim Preserve AC(IA).libEVPTY(IEVPTX)
            AC(IA).libEVPTY(IEVPTX) = StartY
            StartX = StartX + DeltaX
            ReDim Preserve AC(IA).libEVPTX(IEVPTX)
            AC(IA).libEVPTX(IEVPTX) = StartX
            Temp = AC(IA).libEVPTX(IEVPTX)
            '    Debug.Print IEVPTX; IY; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            For IY = 1 To NY
                IEVPTX = IEVPTX + 1S
                ReDim Preserve AC(IA).libEVPTY(IEVPTX)
                ReDim Preserve AC(IA).libEVPTX(IEVPTX)
                AC(IA).libEVPTY(IEVPTX) = AC(IA).libEVPTY(IEVPTX - 1) + DeltaY
                AC(IA).libEVPTX(IEVPTX) = StartX
                '      Debug.Print IEVPTX; IY; IX; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            Next IY
        Next IX
        AC(IA).libNEVPTS = IEVPTX

        IA += 1S

        'modified Izydor Kawa 2012.09.25
        AC(IA).libACName$ = "A380e"
        AC(IA).libGL = 575000 * kgTolb : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 217.6!
        AC(IA).libGear$ = "WFBN"  ' Wing = dual-tandem, Body = dual-tridem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 53.1!      ' Wing dual spacing.
        AC(IA).libTS = 84.95!
        AC(IA).libTG = 207.2
        AC(IA).libB = 70.9!       ' Tandem spacing, wing. Body is 68.9.
        AC(IA).libNTires = 20

        ReDim AC(IA).libTX(AC(IA).libNTires)
        ReDim AC(IA).libTY(AC(IA).libNTires)


        AC(IA).libTX(1) = -271.75 : AC(IA).libTY(1) = 231.35
        AC(IA).libTX(2) = -271.75 : AC(IA).libTY(2) = 162.45
        AC(IA).libTX(3) = -218.65 : AC(IA).libTY(3) = 162.45
        AC(IA).libTX(4) = -218.65 : AC(IA).libTY(4) = 231.35
        AC(IA).libTX(5) = -133.7 : AC(IA).libTY(5) = 135.8
        AC(IA).libTX(6) = -134.1 : AC(IA).libTY(6) = 68.9
        AC(IA).libTX(7) = -133.7 : AC(IA).libTY(7) = 0.0#
        AC(IA).libTX(8) = -73.5 : AC(IA).libTY(8) = 0.0#
        AC(IA).libTX(9) = -73.1 : AC(IA).libTY(9) = 68.9
        AC(IA).libTX(10) = -73.5 : AC(IA).libTY(10) = 135.8
        AC(IA).libTX(11) = 73.5 : AC(IA).libTY(11) = 135.8
        AC(IA).libTX(12) = 73.1 : AC(IA).libTY(12) = 68.9
        AC(IA).libTX(13) = 73.5 : AC(IA).libTY(13) = 0.0#
        AC(IA).libTX(14) = 133.7 : AC(IA).libTY(14) = 0.0#
        AC(IA).libTX(15) = 134.1 : AC(IA).libTY(15) = 68.9
        AC(IA).libTX(16) = 133.7 : AC(IA).libTY(16) = 135.8
        AC(IA).libTX(17) = 218.65 : AC(IA).libTY(17) = 231.35
        AC(IA).libTX(18) = 218.65 : AC(IA).libTY(18) = 162.45
        AC(IA).libTX(19) = 271.75 : AC(IA).libTY(19) = 162.45
        AC(IA).libTX(20) = 271.75 : AC(IA).libTY(20) = 231.35
        AC(IA).libNTTrack = 8
        ReDim AC(IA).libXAC(AC(IA).libNTTrack)
        AC(IA).libXAC(1) = -271.75
        AC(IA).libXAC(2) = -218.65
        AC(IA).libXAC(3) = -133.7
        AC(IA).libXAC(4) = -73.5
        AC(IA).libXAC(5) = 73.5
        AC(IA).libXAC(6) = 133.7
        AC(IA).libXAC(7) = 218.65
        AC(IA).libXAC(8) = 271.75
        AC(IA).libNEVPTS = 8  ' Do the standard dual-tandem first. Center of gear to wheel 3.
        ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
        Temp = AC(IA).libTT * 0.5!
        C = 0.5604! * AC(IA).libTT - 0.2637! * AC(IA).libB ' TT and B are set for the wing gear.
        If C < 0.0! Then C = 0.0!
        AC(IA).libEVPTX(1) = AC(IA).libTX(3)    ' Start at wheel 3 (right rear on the left wing gear).
        AC(IA).libEVPTY(1) = AC(IA).libTY(3)
        AC(IA).libEVPTX(2) = AC(IA).libEVPTX(1) - Temp * 0.2!
        AC(IA).libEVPTY(2) = AC(IA).libEVPTY(1) + C * 0.2!
        AC(IA).libEVPTX(3) = AC(IA).libEVPTX(1) - Temp * 0.4!
        AC(IA).libEVPTY(3) = AC(IA).libEVPTY(1) + C * 0.4!
        AC(IA).libEVPTX(4) = AC(IA).libEVPTX(1) - Temp * 0.6!
        AC(IA).libEVPTY(4) = AC(IA).libEVPTY(1) + C * 0.6!
        AC(IA).libEVPTX(5) = AC(IA).libEVPTX(1) - Temp * 0.8!
        AC(IA).libEVPTY(5) = AC(IA).libEVPTY(1) + C * 0.8!
        AC(IA).libEVPTX(6) = AC(IA).libEVPTX(1) - Temp
        AC(IA).libEVPTY(6) = AC(IA).libEVPTY(1) + C
        AC(IA).libEVPTX(7) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(7) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C) * 0.5!
        AC(IA).libEVPTX(8) = AC(IA).libEVPTX(6)
        AC(IA).libEVPTY(8) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C)   ' Center of gear.
        AC(IA).libNEVPTS = AC(IA).libNEVPTS + 8S  ' Do the standard tridem but with 8 points
        ReDim Preserve AC(IA).libEVPTX(AC(IA).libNEVPTS)
        ReDim Preserve AC(IA).libEVPTY(AC(IA).libNEVPTS)
        ' to be compatible with DT multiple gear.
        Temp = AC(IA).libTX(9) - AC(IA).libTX(6)   ' Center dual spacing slightly larger than TT.
        Temp = Temp / 2 / 7  ' For 8 points from the center to the inboard center wheel.
        AC(IA).libEVPTX(9) = AC(IA).libTX(9) : AC(IA).libEVPTY(9) = AC(IA).libTY(9)
        AC(IA).libEVPTX(10) = AC(IA).libEVPTX(9) - Temp : AC(IA).libEVPTY(10) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(11) = AC(IA).libEVPTX(10) - Temp : AC(IA).libEVPTY(11) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(12) = AC(IA).libEVPTX(11) - Temp : AC(IA).libEVPTY(12) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(13) = AC(IA).libEVPTX(12) - Temp : AC(IA).libEVPTY(13) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(14) = AC(IA).libEVPTX(13) - Temp : AC(IA).libEVPTY(14) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(15) = AC(IA).libEVPTX(14) - Temp : AC(IA).libEVPTY(15) = AC(IA).libEVPTY(9)
        AC(IA).libEVPTX(16) = AC(IA).libEVPTX(15) - Temp : AC(IA).libEVPTY(16) = AC(IA).libEVPTY(9)
        IEVPTX = 16

        For IX = 1 To -16
            'Debug.Print IX; Format(libEVPTX( IX), "000.00"); "  "; Format(libEVPTY( IX), "000.00")
        Next IX

        ' Now put in a full rectangular grid from wheel 1 (left front wheel in left wing gear)
        ' to the origin (CL of fuselage on line through rear wheels of the body gear).
        NX = 4  ' Total number of lateral points = NX + 1. Force an even number of points
        If (NX + 1) Mod 2 <> 0 Then NX = NX + 1 ' laterally so that grid can be split evenly.
        NY = 4  ' Total number of longitudinal points = NY + 1.
        DeltaX = AC(IA).libTX(1) / NX
        DeltaY = AC(IA).libTY(1) / NY
        StartX = -DeltaX
        StartY = 0
        For IX = 1 To NX + 1
            IEVPTX = IEVPTX + 1S
            ReDim Preserve AC(IA).libEVPTY(IEVPTX)
            AC(IA).libEVPTY(IEVPTX) = StartY
            StartX = StartX + DeltaX
            ReDim Preserve AC(IA).libEVPTX(IEVPTX)
            AC(IA).libEVPTX(IEVPTX) = StartX
            Temp = AC(IA).libEVPTX(IEVPTX)
            '    Debug.Print IEVPTX; IY; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            For IY = 1 To NY
                IEVPTX = IEVPTX + 1S
                ReDim Preserve AC(IA).libEVPTY(IEVPTX)
                ReDim Preserve AC(IA).libEVPTX(IEVPTX)
                AC(IA).libEVPTY(IEVPTX) = AC(IA).libEVPTY(IEVPTX - 1) + DeltaY
                AC(IA).libEVPTX(IEVPTX) = StartX
                '      Debug.Print IEVPTX; IY; IX; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            Next IY
        Next IX
        AC(IA).libNEVPTS = IEVPTX

        IA += 1S


        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        'AC(IA).libACName$ = "NAPTF 10 wheels"
        'AC(IA).libGL = 1305125.0! : AC(IA).libMGpcnt = 0.95
        'AC(IA).libCP = 197.0!
        'AC(IA).libGear$ = "WFBN"  ' Wing = dual-tandem, Body = dual-tridem.
        'AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        'AC(IA).libTT = 54.0!      ' Wing dual spacing.
        'AC(IA).libTS = 200.0!
        'AC(IA).libTG = 0
        'AC(IA).libB = 57.0!       ' Tandem spacing, wing. Body is 68.9.
        'AC(IA).libNTires = 20

        'ReDim AC(IA).libTX(AC(IA).libNTires)
        'ReDim AC(IA).libTY(AC(IA).libNTires)

        'Dim tand_01, dual_01, ts_01, dist_01 As Single
        'tand_01 = 57 : dual_01 = 54 : ts_01 = AC(IA).libTS / 2 : dist_01 = 144


        'AC(IA).libTX(1) = -(ts_01 + dist_01 + dual_01) : AC(IA).libTY(1) = tand_01
        'AC(IA).libTX(2) = -(ts_01 + dist_01 + dual_01) : AC(IA).libTY(2) = 0
        'AC(IA).libTX(3) = -(ts_01 + dist_01) : AC(IA).libTY(3) = 0
        'AC(IA).libTX(4) = -(ts_01 + dist_01) : AC(IA).libTY(4) = tand_01
        'AC(IA).libTX(5) = -(ts_01 + dual_01) : AC(IA).libTY(5) = tand_01 * 2
        'AC(IA).libTX(6) = -(ts_01 + dual_01) : AC(IA).libTY(6) = tand_01
        'AC(IA).libTX(7) = -(ts_01 + dual_01) : AC(IA).libTY(7) = 0.0#
        'AC(IA).libTX(8) = -ts_01 : AC(IA).libTY(8) = 0.0#
        'AC(IA).libTX(9) = -ts_01 : AC(IA).libTY(9) = tand_01
        'AC(IA).libTX(10) = -ts_01 : AC(IA).libTY(10) = tand_01 * 2

        'AC(IA).libTX(11) = ts_01 : AC(IA).libTY(11) = tand_01 * 2
        'AC(IA).libTX(12) = ts_01 : AC(IA).libTY(12) = tand_01
        'AC(IA).libTX(13) = ts_01 : AC(IA).libTY(13) = 0.0#
        'AC(IA).libTX(14) = ts_01 + dual_01 : AC(IA).libTY(14) = 0.0#
        'AC(IA).libTX(15) = ts_01 + dual_01 : AC(IA).libTY(15) = tand_01
        'AC(IA).libTX(16) = ts_01 + dual_01 : AC(IA).libTY(16) = tand_01 * 2
        'AC(IA).libTX(17) = ts_01 + dist_01 : AC(IA).libTY(17) = tand_01
        'AC(IA).libTX(18) = ts_01 + dist_01 : AC(IA).libTY(18) = 0
        'AC(IA).libTX(19) = ts_01 + dist_01 + dual_01 : AC(IA).libTY(19) = 0
        'AC(IA).libTX(20) = ts_01 + dist_01 + dual_01 : AC(IA).libTY(20) = tand_01

        'AC(IA).libNTTrack = 8
        'ReDim AC(IA).libXAC(AC(IA).libNTTrack)
        'AC(IA).libXAC(1) = -(ts_01 + dist_01 + dual_01)
        'AC(IA).libXAC(2) = -(ts_01 + dist_01)
        'AC(IA).libXAC(3) = -(ts_01 + dual_01)
        'AC(IA).libXAC(4) = -ts_01
        'AC(IA).libXAC(5) = ts_01
        'AC(IA).libXAC(6) = ts_01 + dual_01
        'AC(IA).libXAC(7) = ts_01 + dist_01
        'AC(IA).libXAC(8) = ts_01 + dist_01 + dual_01
        'AC(IA).libNEVPTS = 8  ' Do the standard dual-tandem first. Center of gear to wheel 3.
        'ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
        'ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
        'Temp = AC(IA).libTT * 0.5!
        'C = 0.5604! * AC(IA).libTT - 0.2637! * AC(IA).libB ' TT and B are set for the wing gear.
        'If C < 0.0! Then C = 0.0!
        'AC(IA).libEVPTX(1) = AC(IA).libTX(3)    ' Start at wheel 3 (right rear on the left wing gear).
        'AC(IA).libEVPTY(1) = AC(IA).libTY(3)
        'AC(IA).libEVPTX(2) = AC(IA).libEVPTX(1) - Temp * 0.2!
        'AC(IA).libEVPTY(2) = AC(IA).libEVPTY(1) + C * 0.2!
        'AC(IA).libEVPTX(3) = AC(IA).libEVPTX(1) - Temp * 0.4!
        'AC(IA).libEVPTY(3) = AC(IA).libEVPTY(1) + C * 0.4!
        'AC(IA).libEVPTX(4) = AC(IA).libEVPTX(1) - Temp * 0.6!
        'AC(IA).libEVPTY(4) = AC(IA).libEVPTY(1) + C * 0.6!
        'AC(IA).libEVPTX(5) = AC(IA).libEVPTX(1) - Temp * 0.8!
        'AC(IA).libEVPTY(5) = AC(IA).libEVPTY(1) + C * 0.8!
        'AC(IA).libEVPTX(6) = AC(IA).libEVPTX(1) - Temp
        'AC(IA).libEVPTY(6) = AC(IA).libEVPTY(1) + C
        'AC(IA).libEVPTX(7) = AC(IA).libEVPTX(6)
        'AC(IA).libEVPTY(7) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C) * 0.5!
        'AC(IA).libEVPTX(8) = AC(IA).libEVPTX(6)
        'AC(IA).libEVPTY(8) = AC(IA).libEVPTY(6) + (AC(IA).libB * 0.5! - C)   ' Center of gear.
        'AC(IA).libNEVPTS = AC(IA).libNEVPTS + 8S  ' Do the standard tridem but with 8 points
        'ReDim Preserve AC(IA).libEVPTX(AC(IA).libNEVPTS)
        'ReDim Preserve AC(IA).libEVPTY(AC(IA).libNEVPTS)
        '' to be compatible with DT multiple gear.
        'Temp = AC(IA).libTX(9) - AC(IA).libTX(6)   ' Center dual spacing slightly larger than TT.
        'Temp = Temp / 2 / 7  ' For 8 points from the center to the inboard center wheel.
        'AC(IA).libEVPTX(9) = AC(IA).libTX(9) : AC(IA).libEVPTY(9) = AC(IA).libTY(9)
        'AC(IA).libEVPTX(10) = AC(IA).libEVPTX(9) - Temp : AC(IA).libEVPTY(10) = AC(IA).libEVPTY(9)
        'AC(IA).libEVPTX(11) = AC(IA).libEVPTX(10) - Temp : AC(IA).libEVPTY(11) = AC(IA).libEVPTY(9)
        'AC(IA).libEVPTX(12) = AC(IA).libEVPTX(11) - Temp : AC(IA).libEVPTY(12) = AC(IA).libEVPTY(9)
        'AC(IA).libEVPTX(13) = AC(IA).libEVPTX(12) - Temp : AC(IA).libEVPTY(13) = AC(IA).libEVPTY(9)
        'AC(IA).libEVPTX(14) = AC(IA).libEVPTX(13) - Temp : AC(IA).libEVPTY(14) = AC(IA).libEVPTY(9)
        'AC(IA).libEVPTX(15) = AC(IA).libEVPTX(14) - Temp : AC(IA).libEVPTY(15) = AC(IA).libEVPTY(9)
        'AC(IA).libEVPTX(16) = AC(IA).libEVPTX(15) - Temp : AC(IA).libEVPTY(16) = AC(IA).libEVPTY(9)
        'IEVPTX = 16

        'For IX = 1 To -16
        '    'Debug.Print IX; Format(libEVPTX( IX), "000.00"); "  "; Format(libEVPTY( IX), "000.00")
        'Next IX

        '' Now put in a full rectangular grid from wheel 1 (left front wheel in left wing gear)
        '' to the origin (CL of fuselage on line through rear wheels of the body gear).
        'NX = 4  ' Total number of lateral points = NX + 1. Force an even number of points
        'If (NX + 1) Mod 2 <> 0 Then NX = NX + 1 ' laterally so that grid can be split evenly.
        'NY = 4  ' Total number of longitudinal points = NY + 1.
        'DeltaX = AC(IA).libTX(1) / NX
        'DeltaY = AC(IA).libTY(1) / NY

        'DeltaX = AC(IA).libTX(1) / NX
        'DeltaY = AC(IA).libTY(5) / NY

        'StartX = -DeltaX
        'StartY = 0
        'For IX = 1 To NX + 1
        '    IEVPTX = IEVPTX + 1S
        '    ReDim Preserve AC(IA).libEVPTY(IEVPTX)
        '    AC(IA).libEVPTY(IEVPTX) = StartY
        '    StartX = StartX + DeltaX
        '    ReDim Preserve AC(IA).libEVPTX(IEVPTX)
        '    AC(IA).libEVPTX(IEVPTX) = StartX
        '    Temp = AC(IA).libEVPTX(IEVPTX)
        '    '    Debug.Print IEVPTX; IY; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
        '    For IY = 1 To NY
        '        IEVPTX = IEVPTX + 1S
        '        ReDim Preserve AC(IA).libEVPTY(IEVPTX)
        '        ReDim Preserve AC(IA).libEVPTX(IEVPTX)
        '        AC(IA).libEVPTY(IEVPTX) = AC(IA).libEVPTY(IEVPTX - 1) + DeltaY
        '        AC(IA).libEVPTX(IEVPTX) = StartX
        '        '      Debug.Print IEVPTX; IY; IX; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
        '    Next IY
        'Next IX
        'AC(IA).libNEVPTS = IEVPTX

        'IA += 1S


        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


    End Sub


    Sub BoeingMultiGear(ByRef IA As Short, ByRef AC() As clsAC.AircraftCharacteristics)

        ' Sets parameters for Boeing multiple gear aircraft.
        ' B-747 started 02-17-03 by GFH.
        ' IA is incremented before the call. Therefore, last statement is to increment IA.

        AC(IA).libACName = "B747-100 SF"
        AC(IA).libGL = 738000.0! : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 232.0!
        AC(IA).libGear$ = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 44.0!      ' Wing and body dual spacing.
        AC(IA).libTS = 97.0!
        AC(IA).libTG = 106.0#
        AC(IA).libB = 58.0!     ' Tandem spacing, same for wing and body.
        Call B747Data(AC, IA)
        IA += 1S
        '_________________________________________________________________________________

        AC(IA).libACName = "B747-200B Combi Mixed"
        AC(IA).libGL = 836000.0! : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 190.0!
        AC(IA).libGear$ = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 44.0!      ' Wing and body dual spacing.
        AC(IA).libTS = 98.0!
        AC(IA).libTG = 106.0#
        AC(IA).libB = 58.0!     ' Tandem spacing, same for wing and body.
        Call B747Data(AC, IA)
        IA += 1S

        '_________________________________________________________________________________
        AC(IA).libACName$ = "B747-300 Combi Mixed"
        AC(IA).libGL = 836000.0! : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 190.0!
        AC(IA).libGear$ = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 44.0!      ' Wing and body dual spacing.
        AC(IA).libTS = 98.0!
        AC(IA).libTG = 106.0#
        AC(IA).libB = 58.0!     ' Tandem spacing, same for wing and body.
        Call B747Data(AC, IA)
        IA += 1S


        'modified Izydor Kawa 2012.09.25
        AC(IA).libACName$ = "B747-400"
        AC(IA).libGL = 877000.0! : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 200.0!
        AC(IA).libGear$ = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 44.0!      ' Wing and body dual spacing.
        AC(IA).libTS = 98.0!
        AC(IA).libTG = 106.0#
        AC(IA).libB = 58.0!     ' Tandem spacing, same for wing and body.
        Call B747Data(AC, IA)
        IA += 1S

        'modified Izydor Kawa 2012.09.25
        AC(IA).libACName$ = "B747-400ER" ' GFH 06-18-04.
        AC(IA).libGL = 913000.0! : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 230 ' 230.0!
        AC(IA).libGear = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 44.0!      ' Wing and body dual spacing.
        AC(IA).libTS = 98.0!
        AC(IA).libTG = 106.0#
        AC(IA).libB = 58.0!     ' Tandem spacing, same for wing and body.
        Call B747Data(AC, IA)
        IA += 1S

        '(6)
        'modified Izydor Kawa 2012.09.25
        AC(IA).libACName$ = "B747-8" ' GFH 06-18-04.
        AC(IA).libGL = 990000.0! : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 221.0!
        AC(IA).libGear = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 46.8!      ' Wing and body dual spacing.
        AC(IA).libTS = 99.5!
        AC(IA).libTG = 106.0#
        AC(IA).libB = 56.5!     ' Tandem spacing, same for wing and body.
        Call B747Data(AC, IA)
        IA += 1S

        '(7)
        'modified Izydor Kawa 2012.09.25
        AC(IA).libACName$ = "B747-8F" ' GFH 06-18-04.
        AC(IA).libGL = 990000.0! : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 221.0!
        AC(IA).libGear = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 46.8!      ' Wing and body dual spacing.
        AC(IA).libTS = 99.5!
        AC(IA).libTG = 106.0#
        AC(IA).libB = 56.5!     ' Tandem spacing, same for wing and body.
        Call B747Data(AC, IA)
        IA += 1S

        '(8)
        'modified Izydor Kawa 2012.09.25
        AC(IA).libACName$ = "B747-SP"
        AC(IA).libGL = 703000.0! : AC(IA).libMGpcnt = 0.95
        AC(IA).libCP = 205.0!
        AC(IA).libGear$ = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        AC(IA).libTT = 44.0!     ' Wing and body dual spacing.
        AC(IA).libTS = 99.5
        AC(IA).libTG = 106.0
        AC(IA).libB = 58.0!       ' Tandem spacing, same for wing and body.
        Call B747Data(AC, IA)
        IA += 1S

        'AC(IA).libACName$ = "B747-8 Intercontinental (Preliminary)"
        'AC(IA).libGL = 978000.0! : AC(IA).libMGpcnt = 0.95
        'AC(IA).libCP = 221.0!
        'AC(IA).libGear$ = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        'AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        'AC(IA).libTT = 46.8!     ' Wing and body dual spacing.
        'AC(IA).libTS = 99.5
        'AC(IA).libTG = 106.0
        'AC(IA).libB = 56.5!       ' Tandem spacing, same for wing and body.
        'Call B747Data(AC, IA)
        'IA += 1S

        'AC(IA).libACName$ = "B747-8"
        'AC(IA).libGL = 990000.0! : AC(IA).libMGpcnt = 0.95
        'AC(IA).libCP = 221.0!
        'AC(IA).libGear$ = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        'AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        'AC(IA).libTT = 46.8!     ' Wing and body dual spacing.
        'AC(IA).libTS = 99.5
        'AC(IA).libTG = 106.0
        'AC(IA).libB = 56.5!       ' Tandem spacing, same for wing and body.
        'Call B747Data(AC, IA)
        'IA += 1S

        'AC(IA).libACName$ = "B747-9 (Preliminary)"
        'AC(IA).libGL = 555000.0! : AC(IA).libMGpcnt = 0.95
        'AC(IA).libCP = 224.0!
        'AC(IA).libGear$ = "WFBF"  ' Wing = dual-tandem, Body = dual-tandem.
        'AC(IA).libIGear = 4       ' Military double-dual tandem, eg B747.
        'AC(IA).libTT = 60.0!     ' Wing and body dual spacing.
        'AC(IA).libTS = 326.0!
        'AC(IA).libTG = 106.0
        'AC(IA).libB = 59.5!       ' Tandem spacing, same for wing and body.
        'Call B747Data(AC, IA)
        'IA += 1S



    End Sub






    Function LPad(ByRef N As Integer, ByRef SS As String) As String
        ' Adds leading spaces to variant string SS to make it N characters long.
        ' Used to format output to a file. #### characters in a Format function
        ' do not force spaces like QuickBasic.
        ' Typically, SS = Format(XX, "0.00")
        Dim ITemp As Integer

        If Not IsDBNull(SS) Then
            ITemp = Len(SS)
        Else
            ITemp = 0
            SS = ""
        End If

        If N - ITemp < 1 Then N = ITemp + 1
        LPad = Space(N - ITemp) & SS
    End Function




    Function LPadArray(ByRef N As Integer, ByVal PreText As String, ByRef SS() As Single, ByVal Index As Short) As String
        Dim ITemp As Integer

        If Index <= UBound(SS, 1) Then
            ITemp = Len(CStr(SS(Index)))
            LPadArray = CStr(SS(Index))
        Else
            ITemp = 0
            LPadArray = ""
        End If

        If Index <= UBound(SS, 1) Then
            If N - ITemp < 1 Then N = ITemp + 1
            LPadArray = PreText & Space(N - ITemp - Len(PreText)) & LPadArray
        Else
            LPadArray = Space(N)
        End If
    End Function


    Function FileErrors(ByRef errVal As Short, ByRef FileName As String) As Short
        ' Return Value  Meaning     Return Value    Meaning
        ' 0             Resume      2               Unrecoverable error
        ' 1             Resume Next 3               Unrecognized error
        Dim MsgType As Short
        Dim Response As Short
        Dim MSG As String
        MsgType = MB_EXCLAIM
        Select Case errVal
            Case Err_DeviceUnavailable ' Error #68
                MSG = "The disk appears to be unavailable."
                MsgType = MB_EXCLAIM + 5
            Case Err_DiskNotReady ' Error #71
                MSG = "The disk is not ready."
            Case Err_DeviceIO
                MSG = "The disk is full."
            Case Err_BadFileName, Err_BadFileNameOrNumber ' Errors #64 & 52
                MSG = FileName & " path or file name is illegal."
            Case Err_PathDoesNotExist ' Error #76
                MSG = "The path for " & FileName & " doesn't exist."
            Case Err_BadFileMode ' Error #54
                MSG = "Can't open " & FileName & " for that type of access."
            Case Err_FileAlreadyOpen ' Error #55
                MSG = FileName & " is already open."
            Case Err_InputPastEndOfFile ' Error #62
                MSG = FileName & " has a nonstandard end-of-file marker,"
                MSG = MSG & "or an attempt was made to read beyond "
                MSG = MSG & "the end-of-file marker."
            Case Err_FileNotFound
                MSG = "Cannot find " & FileName & "."
            Case Else
                FileErrors = 3
                MSG = "Unknown file error"
                Exit Function
        End Select
        Response = CShort(MsgBoxDQ(MSG, MsgType, "File Error"))
        Select Case Response
            Case 4 ' Retry button.
                FileErrors = 0
            Case 5 ' Ignore button.
                FileErrors = 1
            Case 1, 2, 3 ' Ok and Cancel buttons.
                FileErrors = 2
            Case Else
                FileErrors = 3
        End Select
    End Function

    Function MsgBoxDQ(ByRef Caption As String, ByRef I As Short, ByRef Title As String) As Long

        Dim DemoState As Boolean
        Dim NL, NL2 As String
        NL = Environment.NewLine
        NL2 = NL & NL

        DemoState = False

        If DemoState Then
            If I = 0 Then
                Caption = Caption & NL2 & "Click OK to continue."
            Else
                Caption = Caption & NL2 & "Click the appropriate button to continue."
            End If
        End If
        MsgBoxDQ = MsgBox(Caption, CType(I, Microsoft.VisualBasic.MsgBoxStyle), Title)

        'End Function  Must be the last statement, so not here.

        ' See below for useful data input routines.
        ' Sub StripCommas (SS$)
        ' Removes thousands commas from a number string.
        ' Use replaced by GetInputSingle
        '  Dim I As Integer, NSS As Integer, IEnd As Integer
        '  Dim Pos As Integer, LSS$
        '  Const Comma$ = ","
        '  SS$ = Trim$(SS$)
        '  If InStr(1, SS$, Comma$) <> 0 Then
        '    LSS$ = ""
        '    NSS = Len(SS$)
        '    IEnd = NSS \ 4
        '    For I = 1 To IEnd
        '      Pos = InStr(1, SS$, Comma$)
        '      If Pos <> 0 Then LSS$ = LSS$ + Left$(SS$, Pos - 1)
        '      NSS = NSS - Pos
        '      SS$ = Right$(SS$, NSS)
        '    Next I
        '    SS$ = LSS$ + SS$
        '  End If
        ' End Sub
        ' Sub NumberKeys (KeyAscii As Integer)
        ' Not used but could be handy. Use replaced by GetInputSingle.
        '   Msg = "KeyAscii = " & Format(I, "0")
        '   MsgBox Msg
        '   If KeyAscii <> 8 Then         ' Backspace key.
        '     If KeyAscii < 48 Or KeyAscii > 57 Then  ' Number keys.
        '       KeyAscii = 0
        '       Beep
        '     End If
        '   End If
        ' End Sub

    End Function



    Private Sub B747Data(ByRef AC1() As clsAC.AircraftCharacteristics, ByVal I1 As Integer)

        Dim Temp, C As Single
        Dim DeltaX, DeltaY As Single
        Dim StartX, StartY As Single
        Dim IX, IY, NX, NY As Integer
        Dim IEVPTX As Short

        AC1(I1).libNTires = 16
        ReDim AC1(I1).libTX(AC1(I1).libNTires)
        ReDim AC1(I1).libTY(AC1(I1).libNTires)

        AC1(I1).libTX(1) = -(AC1(I1).libTG / 2 + 2 * AC1(I1).libTT + AC1(I1).libTS)
        AC1(I1).libTX(2) = AC1(I1).libTX(1)
        AC1(I1).libTX(3) = -(AC1(I1).libTG / 2 + AC1(I1).libTT + AC1(I1).libTS)
        AC1(I1).libTX(4) = AC1(I1).libTX(3)
        AC1(I1).libTX(5) = -(AC1(I1).libTG / 2 + AC1(I1).libTT)
        AC1(I1).libTX(6) = AC1(I1).libTX(5)
        AC1(I1).libTX(7) = -AC1(I1).libTG / 2
        AC1(I1).libTX(8) = AC1(I1).libTX(7)
        AC1(I1).libTX(9) = AC1(I1).libTG / 2
        AC1(I1).libTX(10) = AC1(I1).libTX(9)
        AC1(I1).libTX(11) = AC1(I1).libTG / 2 + AC1(I1).libTT
        AC1(I1).libTX(12) = AC1(I1).libTX(11)
        AC1(I1).libTX(13) = AC1(I1).libTG / 2 + AC1(I1).libTT + AC1(I1).libTS
        AC1(I1).libTX(14) = AC1(I1).libTX(13)
        AC1(I1).libTX(15) = AC1(I1).libTG / 2 + 2 * AC1(I1).libTT + AC1(I1).libTS
        AC1(I1).libTX(16) = AC1(I1).libTX(15)

  
        If AC1(I1).libACName = "B747-SP" Then
            AC1(I1).libTY(1) = 171.0#
        Else
            AC1(I1).libTY(1) = 179.0#
        End If

        AC1(I1).libTY(2) = AC1(I1).libTY(1) - AC1(I1).libB
        AC1(I1).libTY(3) = AC1(I1).libTY(1) - AC1(I1).libB
        AC1(I1).libTY(4) = AC1(I1).libTY(1)
        AC1(I1).libTY(5) = AC1(I1).libB
        AC1(I1).libTY(6) = 0.0#
        AC1(I1).libTY(7) = 0.0#
        AC1(I1).libTY(8) = AC1(I1).libB
        AC1(I1).libTY(9) = AC1(I1).libB
        AC1(I1).libTY(10) = 0.0#
        AC1(I1).libTY(11) = 0.0#
        AC1(I1).libTY(12) = AC1(I1).libB
        AC1(I1).libTY(13) = AC1(I1).libTY(1)
        AC1(I1).libTY(14) = AC1(I1).libTY(1) - AC1(I1).libB
        AC1(I1).libTY(15) = AC1(I1).libTY(1) - AC1(I1).libB
        AC1(I1).libTY(16) = AC1(I1).libTY(1)

        AC1(I1).libNTTrack = 8
        ReDim AC1(I1).libXAC(AC1(I1).libNTTrack)
        AC1(I1).libXAC(1) = AC1(I1).libTX(1) : AC1(I1).libXAC(5) = AC1(I1).libTX(9)
        AC1(I1).libXAC(2) = AC1(I1).libTX(3) : AC1(I1).libXAC(6) = AC1(I1).libTX(11)
        AC1(I1).libXAC(3) = AC1(I1).libTX(5) : AC1(I1).libXAC(7) = AC1(I1).libTX(13)
        AC1(I1).libXAC(4) = AC1(I1).libTX(7) : AC1(I1).libXAC(8) = AC1(I1).libTX(15)

        AC1(I1).libNEVPTS = 8  ' Do the wing dual-tandem first. Center of gear to wheel 3.
        ReDim AC1(I1).libEVPTX(AC1(I1).libNEVPTS)
        ReDim AC1(I1).libEVPTY(AC1(I1).libNEVPTS)
        Temp = AC1(I1).libTT * 0.5!
        C = 0.5604! * AC1(I1).libTT - 0.2637! * AC1(I1).libB
        If C < 0.0! Then C = 0.0!
        AC1(I1).libEVPTX(1) = AC1(I1).libTX(3)    ' Start at wheel 3 (right rear on left wing gear).
        AC1(I1).libEVPTY(1) = AC1(I1).libTY(3)
        AC1(I1).libEVPTX(2) = AC1(I1).libEVPTX(1) - Temp * 0.2!
        AC1(I1).libEVPTY(2) = AC1(I1).libEVPTY(1) + C * 0.2!
        AC1(I1).libEVPTX(3) = AC1(I1).libEVPTX(1) - Temp * 0.4!
        AC1(I1).libEVPTY(3) = AC1(I1).libEVPTY(1) + C * 0.4!
        AC1(I1).libEVPTX(4) = AC1(I1).libEVPTX(1) - Temp * 0.6!
        AC1(I1).libEVPTY(4) = AC1(I1).libEVPTY(1) + C * 0.6!
        AC1(I1).libEVPTX(5) = AC1(I1).libEVPTX(1) - Temp * 0.8!
        AC1(I1).libEVPTY(5) = AC1(I1).libEVPTY(1) + C * 0.8!
        AC1(I1).libEVPTX(6) = AC1(I1).libEVPTX(1) - Temp
        AC1(I1).libEVPTY(6) = AC1(I1).libEVPTY(1) + C
        AC1(I1).libEVPTX(7) = AC1(I1).libEVPTX(6)
        AC1(I1).libEVPTY(7) = AC1(I1).libEVPTY(6) + (AC1(I1).libB * 0.5! - C) * 0.5!
        AC1(I1).libEVPTX(8) = AC1(I1).libEVPTX(6)
        AC1(I1).libEVPTY(8) = AC1(I1).libEVPTY(6) + (AC1(I1).libB * 0.5! - C)   ' Center of gear.
        AC1(I1).libNEVPTS = AC1(I1).libNEVPTS + 8S  ' Do the body dual tandem.
        ReDim Preserve AC1(I1).libEVPTX(AC1(I1).libNEVPTS)
        ReDim Preserve AC1(I1).libEVPTY(AC1(I1).libNEVPTS)
        Temp = AC1(I1).libTT * 0.5!
        C = 0.5604! * AC1(I1).libTT - 0.2637! * AC1(I1).libB
        If C < 0.0! Then C = 0.0!
        AC1(I1).libEVPTX(9) = AC1(I1).libTX(8)    ' Start at wheel 8 (right front on left body gear).
        AC1(I1).libEVPTY(9) = AC1(I1).libTY(8)
        AC1(I1).libEVPTX(10) = AC1(I1).libEVPTX(9) - Temp * 0.2!
        AC1(I1).libEVPTY(10) = AC1(I1).libEVPTY(9) - C * 0.2!
        AC1(I1).libEVPTX(11) = AC1(I1).libEVPTX(9) - Temp * 0.4!
        AC1(I1).libEVPTY(11) = AC1(I1).libEVPTY(9) - C * 0.4!
        AC1(I1).libEVPTX(12) = AC1(I1).libEVPTX(9) - Temp * 0.6!
        AC1(I1).libEVPTY(12) = AC1(I1).libEVPTY(9) - C * 0.6!
        AC1(I1).libEVPTX(13) = AC1(I1).libEVPTX(9) - Temp * 0.8!
        AC1(I1).libEVPTY(13) = AC1(I1).libEVPTY(9) - C * 0.8!
        AC1(I1).libEVPTX(14) = AC1(I1).libEVPTX(9) - Temp
        AC1(I1).libEVPTY(14) = AC1(I1).libEVPTY(9) - C
        AC1(I1).libEVPTX(15) = AC1(I1).libEVPTX(14)
        AC1(I1).libEVPTY(15) = AC1(I1).libEVPTY(14) - (AC1(I1).libB * 0.5! - C) * 0.5!
        AC1(I1).libEVPTX(16) = AC1(I1).libEVPTX(14)
        AC1(I1).libEVPTY(16) = AC1(I1).libEVPTY(14) - (AC1(I1).libB * 0.5! - C)   ' Center of gear.
        IEVPTX = 16
        For IX = 1 To -16
            'Debug.Print IX; Format(AC(IA).libEVPTX( IX), "000.00"); "  "; Format(AC(IA).libEVPTY( IX), "000.00")
        Next IX
        ' Now put in a full rectangular grid from wheel 1 (left front wheel in left wing gear)
        ' to the origin (CL of fuselage on line through rear wheels of the body gear).
        NX = 4  ' Total number of lateral points = NX + 1. Force an even number of points
        If (NX + 1) Mod 2 <> 0 Then NX = NX + 1 ' laterally so that grid can be split evenly.
        NY = 4  ' Total number of longitudinal points = NY + 1.
        DeltaX = AC1(I1).libTX(1) / NX
        DeltaY = AC1(I1).libTY(1) / NY
        StartX = -DeltaX
        StartY = 0
        For IX = 1 To NX + 1
            IEVPTX = IEVPTX + 1S
            ReDim Preserve AC1(I1).libEVPTY(IEVPTX)
            AC1(I1).libEVPTY(IEVPTX) = StartY
            StartX = StartX + DeltaX
            ReDim Preserve AC1(I1).libEVPTX(IEVPTX)
            AC1(I1).libEVPTX(IEVPTX) = StartX
            Temp = AC1(I1).libEVPTX(IEVPTX)
            '    Debug.Print IEVPTX; IY; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            For IY = 1 To NY
                IEVPTX = IEVPTX + 1S
                ReDim Preserve AC1(I1).libEVPTY(IEVPTX)
                ReDim Preserve AC1(I1).libEVPTX(IEVPTX)
                AC1(I1).libEVPTY(IEVPTX) = AC1(I1).libEVPTY(IEVPTX - 1) + DeltaY
                AC1(I1).libEVPTX(IEVPTX) = StartX
                '      Debug.Print IEVPTX; IY; IX; Format(AC(IA).libEVPTX( IEVPTX), "000.00"); "  "; Format(AC(IA).libEVPTY( IEVPTX), "000.00")
            Next IY
        Next IX
        AC1(I1).libNEVPTS = IEVPTX


    End Sub
End Module
