Option Explicit On
Option Strict On


Imports System.Xml
Imports System.Threading

Public Structure UnitsConversion
    Dim Metric As Boolean
    Dim inch As Double
    Dim inchName As String
    Dim inchFormat As String
    Dim pounds As Double
    Dim poundsName As String
    Dim poundsFormat As String
    Dim psi As Double
    Dim psiName As String
    Dim psiFormat As String
    Dim psiMPa As Double
    Dim psiMPaName As String
    Dim psiMPaFormat As String
    Dim pci As Double
    Dim pciName As String
    Dim pciFormat As String
    'ikawa March 25, 2008
    Dim inch2 As Double
    Dim inch2Name As String
    Dim inch2Format As String
End Structure



Module Module1

    Public UnitsOut As UnitsConversion





    Sub ReadExternalXMLFile(ByRef AC() As clsAC.AircraftCharacteristics, ByRef IA As Short)


        Try


        Dim reader As XmlTextReader = New XmlTextReader("FAAairplaneLibrary.xml")
        Dim Element As String, iWheel, iEval As Integer
        Dim iNG As Integer

        Element = "" 'i = 207  '181, 195,  207

        Dim ConvertToDot As Boolean

        If Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator.ToString = "." Then
            ConvertToDot = True
        Else
            ConvertToDot = False
        End If


        Do While (reader.Read())
            Select Case reader.NodeType
                Case XmlNodeType.Element 'Display beginning of element. (like GrossWt)
                    'Debug.Write("<" + reader.Name)
                    Element = reader.Name

                    If reader.HasAttributes Then 'If attributes exist
                        While reader.MoveToNextAttribute()
                            'Display attribute name and value.
                            'Debug.Write(" {0}='{1}' " & reader.Name & "   " & reader.Value)
                        End While
                    End If

                    'Debug.WriteLine(">")
                Case XmlNodeType.Text 'Display the text in each element. (like 25,000)
                    If Element = "No" Then
                        'IA += 1 : iWheel = 1 : iEval = 1
                        'ReDim Preserve AC(IA)
                    ElseIf Element = "Name" Then
                        IA += 1S : iWheel = 1 : iEval = 1
                        ReDim Preserve AC(IA)

                        AC(IA).libACName = reader.Value
                        iNG = 0
                    ElseIf Element = "GrossWt" Then

                        If ConvertToDot Then
                            AC(IA).libGL = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libGL = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libGL = CSng(reader.Value)

                    ElseIf Element = "MGpcnt" Then

                        If ConvertToDot Then
                            AC(IA).libMGpcnt = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libMGpcnt = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libMGpcnt = CSng(reader.Value)



                    ElseIf Element = "MGpcntPCN" Then

                        If ConvertToDot Then
                            AC(IA).libMGpcntPCN = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libMGpcntPCN = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libMGpcnt = CSng(reader.Value)




                    ElseIf Element = "CP" Then

                        If ConvertToDot Then
                            AC(IA).libCP = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libCP = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libCP = CSng(reader.Value)

                    ElseIf Element = "Gear" Then
                        AC(IA).libGear = reader.Value
                    ElseIf Element = "IGear" Then

                        If ConvertToDot Then
                            AC(IA).libIGear = CShort(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libIGear = CShort(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libIGear = CShort(reader.Value)

                        If AC(IA).libGear = "X" Or AC(IA).libGear = "WFBN" Or AC(IA).libGear = "WFBF" Then
                            IA = IA
                        End If

                    ElseIf Element = "TT" Then

                        If ConvertToDot Then
                            AC(IA).libTT = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTT = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libTT = CSng(reader.Value)

                    ElseIf Element = "TS" Then

                        If ConvertToDot Then
                            AC(IA).libTS = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTS = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libTS = CSng(reader.Value)

                    ElseIf Element = "TG" Then

                        If ConvertToDot Then
                            AC(IA).libTG = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTG = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libTG = CSng(reader.Value)

                    ElseIf Element = "B" Then

                        If ConvertToDot Then
                            AC(IA).libB = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libB = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libB = CSng(reader.Value)

                    ElseIf Element = "TV" Then

                        If ConvertToDot Then
                            AC(IA).libTV = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTV = CSng(reader.Value.Replace(".", ","))
                        End If

                    ElseIf Element = "NTires" Then

                        iNG = iNG + 1

                        If iNG = 2 Then
                            iWheel = iWheel + AC(IA).libNTires1
                        End If


                        If iNG = 1 Then
                            If ConvertToDot Then
                                AC(IA).libNTires1 = CShort(reader.Value.Replace(",", "."))
                            Else
                                AC(IA).libNTires1 = CShort(reader.Value.Replace(".", ","))
                            End If
                        Else
                            If ConvertToDot Then
                                AC(IA).libNTires2 = CShort(reader.Value.Replace(",", "."))
                            Else
                                AC(IA).libNTires2 = CShort(reader.Value.Replace(".", ","))
                            End If
                        End If

                        If iNG = 1 Then
                            If AC(IA).libGear = "X" Or AC(IA).libGear = "WFBN" Or AC(IA).libGear = "WFBF" Then
                                AC(IA).libNTires = CShort(AC(IA).libNTires1 * 2)
                                ReDim AC(IA).libTX(AC(IA).libNTires1 * 2)
                                ReDim AC(IA).libTY(AC(IA).libNTires1 * 2)
                            Else
                                AC(IA).libNTires = AC(IA).libNTires1
                                ReDim AC(IA).libTX(AC(IA).libNTires1)
                                ReDim AC(IA).libTY(AC(IA).libNTires1)
                            End If

                        Else
                            If AC(IA).libGear = "X" Or AC(IA).libGear = "WFBN" Or AC(IA).libGear = "WFBF" Then
                                AC(IA).libNTires = AC(IA).libNTires + CShort(AC(IA).libNTires2 * 2)
                                ReDim Preserve AC(IA).libTX(AC(IA).libNTires)
                                ReDim Preserve AC(IA).libTY(AC(IA).libNTires)
                            Else

                            End If
                        End If

                    ElseIf Element = "TX" Then

                        If AC(IA).libTX Is Nothing Then
                            MsgBox("Number of Wheels is Missing!", MsgBoxStyle.OkOnly, "Aircraft: " & AC(IA).libACName)
                        Else

                        End If

                        If ConvertToDot Then
                            AC(IA).libTX(iWheel) = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTX(iWheel) = CSng(reader.Value.Replace(".", ","))
                        End If

                        If AC(IA).libGear = "X" Or AC(IA).libGear = "WFBN" Or AC(IA).libGear = "WFBF" Then
                            'AC(IA).libTX(iWheel + CInt(AC(IA).libWheelsInGroup(iNG) / 2)) = -AC(IA).libTX(iWheel)
                            If iNG = 1 Then
                                AC(IA).libTX(iWheel + CInt(AC(IA).libNTires1)) = -AC(IA).libTX(iWheel)
                            Else
                                AC(IA).libTX(iWheel + CInt(AC(IA).libNTires2)) = -AC(IA).libTX(iWheel)
                            End If

                        End If


                    ElseIf Element = "TY" Then

                        If AC(IA).libTY Is Nothing Then
                            MsgBox("Number of Wheels is Missing!", MsgBoxStyle.OkOnly, "Message")
                        Else

                        End If


                        If ConvertToDot Then
                            AC(IA).libTY(iWheel) = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTY(iWheel) = CSng(reader.Value.Replace(".", ","))
                        End If


                        If AC(IA).libGear = "X" Or AC(IA).libGear = "WFBN" Or AC(IA).libGear = "WFBF" Then
                            'AC(IA).libTY(iWheel + CInt(AC(IA).libWheelsInGroup(iNG) / 2)) = AC(IA).libTY(iWheel)
                            'AC(IA).libTY(iWheel + CInt(AC(IA).libNTires1 * 2)) = AC(IA).libTY(iWheel)

                            If iNG = 1 Then
                                AC(IA).libTY(iWheel + CInt(AC(IA).libNTires1)) = AC(IA).libTY(iWheel)
                            Else
                                AC(IA).libTY(iWheel + CInt(AC(IA).libNTires2)) = AC(IA).libTY(iWheel)
                            End If

                        End If


                        iWheel += 1
                    ElseIf Element = "NEVPTS" Then

                        If iNG = 1 Then
                            If ConvertToDot Then
                                AC(IA).libNEVPTS1 = CShort(reader.Value.Replace(",", "."))
                            Else
                                AC(IA).libNEVPTS1 = CShort(reader.Value.Replace(".", ","))
                            End If
                        Else
                            If ConvertToDot Then
                                AC(IA).libNEVPTS2 = CShort(reader.Value.Replace(",", "."))
                            Else
                                AC(IA).libNEVPTS2 = CShort(reader.Value.Replace(".", ","))
                            End If
                        End If



                        If iNG = 1 Then
                            ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS1)
                            ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS1)
                            AC(IA).libNEVPTS = AC(IA).libNEVPTS1
                        Else
                            ReDim Preserve AC(IA).libEVPTX(AC(IA).libNEVPTS1 + AC(IA).libNEVPTS2)
                            ReDim Preserve AC(IA).libEVPTY(AC(IA).libNEVPTS1 + AC(IA).libNEVPTS2)
                            AC(IA).libNEVPTS = AC(IA).libNEVPTS1 + AC(IA).libNEVPTS2
                        End If

                        'ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                        'ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)


                    ElseIf Element = "EVPTX" Then

                        If ConvertToDot Then
                            AC(IA).libEVPTX(iEval) = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libEVPTX(iEval) = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libEVPTX(iEval) = CSng(reader.Value)

                    ElseIf Element = "EVPTY" Then

                        If ConvertToDot Then
                            AC(IA).libEVPTY(iEval) = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libEVPTY(iEval) = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libEVPTY(iEval) = CSng(reader.Value)
                        iEval += 1

                    ElseIf Element = "NWheelGroups" Then

                        If ConvertToDot Then
                            AC(IA).libNGroups = CShort(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libNGroups = CShort(reader.Value.Replace(".", ","))
                        End If

                        'ReDim AC(IA).libWheelsInGroup(AC(IA).libNGroups)
                        'ReDim AC(IA).libEvalPointsInGroup(AC(IA).libNGroups)

                    ElseIf Element = "WheelsInGroup" Then

                        iNG = iNG + 1

                        If ConvertToDot Then
                            AC(IA).libWheelsInGroup(iNG) = CShort(CShort(reader.Value.Replace(",", ".")) * 2)
                        Else
                            AC(IA).libWheelsInGroup(iNG) = CShort(CShort(reader.Value.Replace(".", ",")) * 2)
                        End If

                        If iNG = 1 Then
                            ReDim AC(IA).libTX(AC(IA).libWheelsInGroup(iNG))
                            AC(IA).libNTires = AC(IA).libWheelsInGroup(iNG)

                        Else
                            ReDim Preserve AC(IA).libTX(AC(IA).libTX.Length + AC(IA).libWheelsInGroup(iNG) - 1)
                            AC(IA).libNTires = AC(IA).libNTires + AC(IA).libWheelsInGroup(iNG)
                        End If

                        If iNG = 1 Then
                            ReDim AC(IA).libTY(AC(IA).libWheelsInGroup(iNG))
                        Else
                            ReDim Preserve AC(IA).libTY(AC(IA).libTY.Length + AC(IA).libWheelsInGroup(iNG) - 1)
                        End If


                        If iNG = 2 Then
                            iWheel = iWheel + CInt(AC(IA).libWheelsInGroup(1) / 2)
                        End If

                    ElseIf Element = "EvaluationPointsInGroup" Then

                        If ConvertToDot Then
                            AC(IA).libEvalPointsInGroup(iNG) = CShort(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libEvalPointsInGroup(iNG) = CShort(reader.Value.Replace(".", ","))
                        End If

                        If iNG = 1 Then
                            ReDim AC(IA).libEVPTX(AC(IA).libEvalPointsInGroup(iNG))
                            AC(IA).libNEVPTS = AC(IA).libEvalPointsInGroup(iNG)
                        Else
                            ReDim Preserve AC(IA).libEVPTX(AC(IA).libEVPTX.Length + AC(IA).libEvalPointsInGroup(iNG) - 1)
                            AC(IA).libNEVPTS = AC(IA).libNEVPTS + AC(IA).libEvalPointsInGroup(iNG)
                        End If


                        If iNG = 1 Then
                            ReDim AC(IA).libEVPTY(AC(IA).libEvalPointsInGroup(iNG))
                        Else
                            ReDim Preserve AC(IA).libEVPTY(AC(IA).libEVPTY.Length + AC(IA).libEvalPointsInGroup(iNG) - 1)
                        End If



                    End If

                    'Debug.WriteLine(reader.Value)

                Case XmlNodeType.EndElement 'Display end of element.
                    'Debug.Write("</" + reader.Name) : Debug.WriteLine(">")
            End Select

        Loop

        'Debug.ReadLine()

        Catch ex As Exception

            Dim txt As String
            txt = ex.Message
            txt = txt + Environment.NewLine + Environment.NewLine
            txt = txt + ex.StackTrace
            txt = txt + Environment.NewLine + Environment.NewLine
            MsgBox(txt)

        End Try

    End Sub




    Sub ReadExternalXMLFile_old(ByRef AC() As clsAC.AircraftCharacteristics, ByRef IA As Short)
        Dim reader As XmlTextReader = New XmlTextReader("FAAairplaneLibrary.xml")
        Dim Element As String, iWheel, iEval As Integer
        Element = "" 'i = 207  '181, 195,  207

        Dim ConvertToDot As Boolean

        If Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator.ToString = "." Then
            ConvertToDot = True
        Else
            ConvertToDot = False
        End If


        Do While (reader.Read())
            Select Case reader.NodeType
                Case XmlNodeType.Element 'Display beginning of element. (like GrossWt)
                    'Debug.Write("<" + reader.Name)
                    Element = reader.Name

                    If reader.HasAttributes Then 'If attributes exist
                        While reader.MoveToNextAttribute()
                            'Display attribute name and value.
                            'Debug.Write(" {0}='{1}' " & reader.Name & "   " & reader.Value)
                        End While
                    End If

                    Debug.WriteLine(">")
                Case XmlNodeType.Text 'Display the text in each element. (like 25,000)
                    If Element = "No" Then
                        'IA += 1 : iWheel = 1 : iEval = 1
                        'ReDim Preserve AC(IA)
                    ElseIf Element = "Name" Then
                        IA += 1S : iWheel = 1 : iEval = 1
                        ReDim Preserve AC(IA)

                        AC(IA).libACName = reader.Value
                    ElseIf Element = "GrossWt" Then

                        If ConvertToDot Then
                            AC(IA).libGL = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libGL = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libGL = CSng(reader.Value)

                    ElseIf Element = "MGpcnt" Then

                        If ConvertToDot Then
                            AC(IA).libMGpcnt = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libMGpcnt = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libMGpcnt = CSng(reader.Value)

                    ElseIf Element = "CP" Then

                        If ConvertToDot Then
                            AC(IA).libCP = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libCP = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libCP = CSng(reader.Value)

                    ElseIf Element = "Gear" Then
                        AC(IA).libGear = reader.Value
                    ElseIf Element = "IGear" Then

                        If ConvertToDot Then
                            AC(IA).libIGear = CShort(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libIGear = CShort(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libIGear = CShort(reader.Value)

                    ElseIf Element = "TT" Then

                        If ConvertToDot Then
                            AC(IA).libTT = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTT = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libTT = CSng(reader.Value)

                    ElseIf Element = "TS" Then

                        If ConvertToDot Then
                            AC(IA).libTS = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTS = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libTS = CSng(reader.Value)

                    ElseIf Element = "TG" Then

                        If ConvertToDot Then
                            AC(IA).libTG = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTG = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libTG = CSng(reader.Value)

                    ElseIf Element = "B" Then

                        If ConvertToDot Then
                            AC(IA).libB = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libB = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libB = CSng(reader.Value)

                    ElseIf Element = "NTires" Then

                        If ConvertToDot Then
                            AC(IA).libNTires = CShort(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libNTires = CShort(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libNTires = CShort(reader.Value)

                        ReDim AC(IA).libTX(AC(IA).libNTires)
                        ReDim AC(IA).libTY(AC(IA).libNTires)
                    ElseIf Element = "TX" Then

                        If ConvertToDot Then
                            AC(IA).libTX(iWheel) = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTX(iWheel) = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libTX(iWheel) = CSng(reader.Value)

                    ElseIf Element = "TY" Then

                        If ConvertToDot Then
                            AC(IA).libTY(iWheel) = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libTY(iWheel) = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libTY(iWheel) = CSng(reader.Value)

                        iWheel += 1
                    ElseIf Element = "NEVPTS" Then

                        If ConvertToDot Then
                            AC(IA).libNEVPTS = CShort(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libNEVPTS = CShort(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libNEVPTS = CShort(reader.Value)

                        ReDim AC(IA).libEVPTX(AC(IA).libNEVPTS)
                        ReDim AC(IA).libEVPTY(AC(IA).libNEVPTS)
                    ElseIf Element = "EVPTX" Then

                        If ConvertToDot Then
                            AC(IA).libEVPTX(iEval) = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libEVPTX(iEval) = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libEVPTX(iEval) = CSng(reader.Value)

                    ElseIf Element = "EVPTY" Then

                        If ConvertToDot Then
                            AC(IA).libEVPTY(iEval) = CSng(reader.Value.Replace(",", "."))
                        Else
                            AC(IA).libEVPTY(iEval) = CSng(reader.Value.Replace(".", ","))
                        End If
                        'AC(IA).libEVPTY(iEval) = CSng(reader.Value)
                        iEval += 1
                    End If

                    'Debug.WriteLine(reader.Value)

                Case XmlNodeType.EndElement 'Display end of element.
                    'Debug.Write("</" + reader.Name) : Debug.WriteLine(">")
            End Select

        Loop

        'Debug.ReadLine()

    End Sub



    Sub WriteExternalXMLFile_old(ByRef AC() As clsAC.AircraftCharacteristics, ByVal iStart As Integer, ByRef iEnd As Short)

        Try


            With UnitsOut
                '     Use these to input/output from English to English.
                .Metric = False
                .inch = 1
                .inchName = "in"
                .inchFormat = "#,##0.00" ' "#,##0.00" ' GFH 3/4/03. Change for more precision in output.

                .inch2 = 1
                .inch2Name = "in^2"
                .inch2Format = "#,##0.00" ' "#,##0.00" 

                .pounds = 1
                .poundsName = "lbs"
                .poundsFormat = "#,###,##0"
                .psi = 1
                .psiName = "psi"
                .psiFormat = "#,##0"
                .psiMPa = 1
                .psiMPaName = "psi"
                .psiMPaFormat = "#,##0"
                .pci = 1
                .pciName = "pci"
                .pciFormat = "#,##0.0"
            End With



            Dim xmlFile As String

            xmlFile = "FAAairplaneLibrary.xml"
            Dim xw As New Xml.XmlTextWriter(xmlFile, System.Text.Encoding.UTF8)

            xw.WriteStartDocument()
            xw.WriteStartElement("FAAairplaneLibrary")

            'For i As Int16 = 182 To 195
            For i As Int16 = CShort(iStart + 1) To iEnd
                xw.WriteStartElement("AirplaneInfo")
                xw.WriteElementString("Name", AC(i).libACName)
                xw.WriteElementString("GrossWt", CStr(Format(AC(i).libGL * UnitsOut.pounds, "#0")))

                xw.WriteElementString("MGpcnt", Format(AC(i).libMGpcnt, "#0.00").ToString)
                xw.WriteElementString("CP", Format(AC(i).libCP * UnitsOut.pci, "#0.0").ToString)
                xw.WriteElementString("Gear", AC(i).libGear)
                xw.WriteElementString("IGear", AC(i).libIGear.ToString)

                xw.WriteElementString("TT", Format(AC(i).libTT * UnitsOut.inch, "#0.00").ToString)
                xw.WriteElementString("TS", Format(AC(i).libTS * UnitsOut.inch, "#0.00").ToString)
                xw.WriteElementString("TG", Format(AC(i).libTG * UnitsOut.inch, "#0.00").ToString)
                xw.WriteElementString("B", Format(AC(i).libB * UnitsOut.inch, "#0.00").ToString)

                xw.WriteElementString("NTires", AC(i).libNTires.ToString)

                xw.WriteStartElement("Wheel_Coordinates")
                For i1 As Int16 = 1 To CShort(AC(i).libNTires)
                    xw.WriteElementString("TX", Format(AC(i).libTX(i1), "#0.00").ToString)
                    xw.WriteElementString("TY", Format(AC(i).libTY(i1), "#0.00").ToString)
                Next
                xw.WriteEndElement()

                xw.WriteElementString("NEVPTS", AC(i).libNEVPTS.ToString)

                xw.WriteStartElement("Evaluation_Points")
                For i1 As Int16 = 1 To CShort(AC(i).libNEVPTS)
                    xw.WriteElementString("EVPTX", Format(AC(i).libEVPTX(i1), "#0.00").ToString)
                    xw.WriteElementString("EVPTY", Format(AC(i).libEVPTY(i1), "#0.00").ToString)
                Next

                xw.WriteEndElement()
                xw.WriteEndElement()

            Next

            xw.WriteEndElement()
            xw.WriteEndDocument()
            xw.Close()


        Catch ex As Exception

            Dim txt As String
            txt = ex.Message
            txt = txt + Environment.NewLine + Environment.NewLine
            txt = txt + ex.StackTrace
            txt = txt + Environment.NewLine + Environment.NewLine
            MsgBox(txt)

        End Try


    End Sub





    Sub WriteExternalXMLFile(ByRef AC() As clsAC.AircraftCharacteristics, ByVal iStart As Integer, ByRef iEnd As Short)

        Try


            With UnitsOut
                '     Use these to input/output from English to English.
                .Metric = False
                .inch = 1
                .inchName = "in"
                .inchFormat = "#,##0.00" ' "#,##0.00" ' GFH 3/4/03. Change for more precision in output.

                .inch2 = 1
                .inch2Name = "in^2"
                .inch2Format = "#,##0.00" ' "#,##0.00" 

                .pounds = 1
                .poundsName = "lbs"
                .poundsFormat = "#,###,##0"
                .psi = 1
                .psiName = "psi"
                .psiFormat = "#,##0"
                .psiMPa = 1
                .psiMPaName = "psi"
                .psiMPaFormat = "#,##0"
                .pci = 1
                .pciName = "pci"
                .pciFormat = "#,##0.0"
            End With



            Dim xmlFile As String

            xmlFile = "FAAairplaneLibrary.xml"
            Dim xw As New Xml.XmlTextWriter(xmlFile, System.Text.Encoding.UTF8)

            xw.WriteStartDocument()
            xw.WriteStartElement("FAAairplaneLibrary")

            'For i As Int16 = 182 To 195
            For i As Int16 = CShort(iStart + 1) To iEnd
                xw.WriteStartElement("AirplaneInfo")
                xw.WriteElementString("Name", AC(i).libACName)
                xw.WriteElementString("GrossWt", CStr(Format(AC(i).libGL * UnitsOut.pounds, "#0")))

                xw.WriteElementString("MGpcnt", Format(AC(i).libMGpcnt, "#0.00").ToString)
                xw.WriteElementString("MGpcntPCN", Format(AC(i).libMGpcntPCN, "#0.00").ToString) 'newPCN
                xw.WriteElementString("CP", Format(AC(i).libCP * UnitsOut.pci, "#0.0").ToString)
                xw.WriteElementString("Gear", AC(i).libGear)
                xw.WriteElementString("IGear", AC(i).libIGear.ToString)

                xw.WriteElementString("TT", Format(AC(i).libTT * UnitsOut.inch, "#0.00").ToString)
                xw.WriteElementString("TS", Format(AC(i).libTS * UnitsOut.inch, "#0.00").ToString)
                xw.WriteElementString("TG", Format(AC(i).libTG * UnitsOut.inch, "#0.00").ToString)
                xw.WriteElementString("B", Format(AC(i).libB * UnitsOut.inch, "#0.00").ToString)

                xw.WriteElementString("NTires", AC(i).libNTires.ToString)

                xw.WriteStartElement("Wheel_Coordinates")
                For i1 As Int16 = 1 To CShort(AC(i).libNTires)
                    xw.WriteElementString("TX", Format(AC(i).libTX(i1), "#0.00").ToString)
                    xw.WriteElementString("TY", Format(AC(i).libTY(i1), "#0.00").ToString)
                Next
                xw.WriteEndElement()

                xw.WriteElementString("NEVPTS", AC(i).libNEVPTS.ToString)

                xw.WriteStartElement("Evaluation_Points")
                For i1 As Int16 = 1 To CShort(AC(i).libNEVPTS)
                    xw.WriteElementString("EVPTX", Format(AC(i).libEVPTX(i1), "#0.00").ToString)
                    xw.WriteElementString("EVPTY", Format(AC(i).libEVPTY(i1), "#0.00").ToString)
                Next

                xw.WriteEndElement()
                xw.WriteEndElement()

            Next

            xw.WriteEndElement()
            xw.WriteEndDocument()
            xw.Close()


        Catch ex As Exception

            Dim txt As String
            txt = ex.Message
            txt = txt + Environment.NewLine + Environment.NewLine
            txt = txt + ex.StackTrace
            txt = txt + Environment.NewLine + Environment.NewLine
            MsgBox(txt)

        End Try


    End Sub







End Module
