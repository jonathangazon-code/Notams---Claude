Namespace Interfaces
    Public Interface IModelAction
        Function Validate(measurementSystem As IMeasurmentSystem) As List(Of IValidation)
    End Interface
End Namespace