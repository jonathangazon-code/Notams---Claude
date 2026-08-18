Namespace Interfaces
    Public Interface IDesignOptions
        Property MeasurementSystem As IMeasurmentSystem
        Property CdfTolerance As Double
        Property LifeTolerance As Double
        Property CalculateHmaCdf As Boolean
        Property AlternateSubgrade As Boolean
        Property AutomaticFlexibleBaseDesign As Boolean
        Property SectionParameterN As Double
        Property AllowPartiallyBonded As Boolean
        Property PCAConversion As Boolean
        Property Outfile As Boolean
        Property CrackPropogation As IDimensionalProperty
        Property ComputeCompaction As Boolean
        Property ThickPccOverlay As Boolean
        Property ACROptions As Boolean
    End Interface
End Namespace