Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces


''' <summary>
''' US Customary Measurement System
''' </summary>
<DataContract>
Public Class UsCustomary
    Implements IMeasurmentSystem
    ''' <summary>
    ''' Display name of the Measurement system
    ''' </summary>
    ''' <returns></returns>
    <DataMember>
    Public Property Name As String Implements IMeasurmentSystem.Name

    ''' <summary>
    ''' This is the default constructor for the UsCustomary class.
    ''' It also initializes the name to US Customary.
    ''' </summary>
    Sub New()
        Name = "US Customary"
    End Sub

    ''' <summary>
    ''' Selects the US Customary value (because it that is what this Concrete class does)
    ''' </summary>
    ''' <param name="us">US Customary string value</param>
    ''' <param name="si">SI string value</param>
    ''' <returns>ALWAYS return US Customary</returns>
    Public Function GetValue(us As String, si As String) As String Implements IMeasurmentSystem.GetValue
        'This function avoids IF statements where choices are made.
        Return us
    End Function

    ''' <summary>
    ''' Selects the US Customary numeric value.
    ''' </summary>
    ''' <param name="us">US Customary numeric value</param>
    ''' <param name="si">SI numeric value</param>
    ''' <returns>ALWAYS US Customary</returns>
    Public Function GetValue(us As Double, si As Double) As Double Implements IMeasurmentSystem.GetValue
        'This function avoids IF statements for Measurement Systems
        Return us
    End Function

    Public Sub SetValue(value As Double, convert As Double, ByRef us As Double, ByRef si As Double) Implements IMeasurmentSystem.SetValue
        us = value
        si = us * convert
    End Sub

    ''' <summary>
    ''' Sets the temperature depending on the measurement system.
    ''' Converts to metric if needed.
    ''' </summary>
    ''' <param name="us">US Customary numeric value</param>
    ''' <param name="si">SI numeric value</param>
    ''' <param name= "value"> The numeric value passed in </param>
    Public Sub SetTemperatureValue(value As Double, ByRef us As Double, ByRef si As Double) Implements IMeasurmentSystem.SetTemperatureValue
        us = value
        si = (value - 32) / 1.8
    End Sub
End Class
