Namespace Interfaces
    Public Interface IValidationRange
        ''' <summary>
        ''' Check the input value using the input measurement system
        ''' </summary>
        ''' <param name="value">Value to check for validity</param>
        ''' <param name="measurementSystem">Measuremnt system value is from</param>
        ''' <returns>IValidate with validation method</returns>
        Function Validate(value As Double, form As String, control As String, label As String, measurementSystem As IMeasurmentSystem) As IValidation
        ''' <summary>
        ''' This variable validation range is required for the analysis
        ''' </summary>
        ''' <returns></returns>
        Property Required As Boolean
    End Interface
End Namespace