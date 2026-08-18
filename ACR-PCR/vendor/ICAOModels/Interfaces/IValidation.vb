Namespace Interfaces
    Public Interface IValidation
        ''' <summary>
        ''' Form that validation message appears on
        ''' </summary>
        ''' <returns></returns>
        Property Form As String
        ''' <summary>
        ''' Label in form that message is appended to
        ''' </summary>
        ''' <returns></returns>
        Property Label As String
        ''' <summary>
        ''' The control that gets focus on specific error
        ''' </summary>
        ''' <returns></returns>
        Property Control As String
        ''' <summary>
        ''' Validation message that gets displayed
        ''' </summary>
        ''' <returns></returns>
        Property Message As String
    End Interface
End Namespace