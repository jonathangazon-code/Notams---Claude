Imports System.Runtime.Serialization
Imports ICAOModels.Interfaces
<DataContract>
<KnownType(GetType(LengthCoordinates))>
<KnownType(GetType(Length))>
Public Class LengthCoordinates
    Implements ICoordinates

    <DataMember>
    Public Property X As IDimensionalProperty Implements ICoordinates.X
    <DataMember>
    Public Property Y As IDimensionalProperty Implements ICoordinates.Y

    ''' <summary>
    ''' This is the default constructor of the LengthCoordinates class.
    ''' </summary>
    Sub New()

    End Sub

    ''' <summary>
    ''' This is an alternate constructor for the LengthCoordinates class.
    ''' It uses two IDimensionalProperty objects as parameters.
    ''' </summary>
    ''' <param name="x">This is one of the IDimensionalProperty objects meant to represent the X coordinate.</param>
    ''' <param name="y">This is the second IDimensionalProperty object meant to represent the Y coordinate.</param>
    Sub New(x As IDimensionalProperty, y As IDimensionalProperty)
        Me.X = x
        Me.Y = y
    End Sub

    Public Overrides Function ToString() As String
        Return "(" + X.GetValue(New UsCustomary()).ToString() + "," + Y.GetValue(New UsCustomary()).ToString() + ")"
    End Function
End Class
