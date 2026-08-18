Imports ICAOModels.Interfaces

Public Class CoordinateValidationRange
    Implements IValidationRange

    ''' <summary>
    ''' X direction validation
    ''' </summary>
    Public Property ValidationX As IValidationRange
    ''' <summary>
    ''' Y direction validation
    ''' </summary>
    Public Property ValidationY As IValidationRange
    ''' <summary>
    ''' This validation ranged must be run
    ''' </summary>
    Public Property Required As Boolean Implements IValidationRange.Required

    ''' <summary>
    ''' Constructor to made to set the x and y coordinates to the validation instances.
    ''' </summary>
    ''' <param name="validationX">This is the x value.</param>
    ''' <param name="validationY">This is the Y value.</param>
    ''' <param name="required">This is a boolean variable made to represent if the validation range is required or not.</param>
    Sub New(validationX As IValidationRange, validationY As IValidationRange, required As Boolean)
        Me.ValidationX = validationX
        Me.ValidationY = validationY
        Me.Required = required
    End Sub

    Public Function Validate(value As Double, form As String, control As String, label As String, measurementSystem As IMeasurmentSystem) As IValidation Implements IValidationRange.Validate
        Return Nothing
    End Function
End Class

