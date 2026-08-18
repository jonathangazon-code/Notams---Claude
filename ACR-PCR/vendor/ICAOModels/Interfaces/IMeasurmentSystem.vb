Namespace Interfaces
    Public Interface IMeasurmentSystem
        ''' <summary>
        ''' The display name of the measurement system
        ''' </summary>
        ''' <returns>Name of the Unit of measurement</returns>
        Property Name As String

        ''' <summary>
        ''' For setting values in US and SI (non-temperature).
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="convert"></param>
        Sub SetValue(value As Double, convert As Double, ByRef us As Double, ByRef si As Double)

        ''' <summary>
        ''' Used for setting temperatures (not a straight conversion)
        ''' </summary>
        ''' <param name="value"></param>
        Sub SetTemperatureValue(value As Double, ByRef us As Double, ByRef si As Double)

        ''' <summary>
        ''' Selects between US and SI strings
        ''' </summary>
        ''' <param name="us">US version of string</param>
        ''' <param name="si">SI version of string</param>
        ''' <returns>Returns string for Measurement System</returns>
        Function GetValue(us As String, si As String) As String

        ''' <summary>
        ''' Selects between US and SI numeric values
        ''' </summary>
        ''' <param name="us">US value</param>
        ''' <param name="si">SI value</param>
        ''' <returns>Numeric value for Measurement system</returns>
        Function GetValue(us As Double, si As Double) As Double
    End Interface
End Namespace