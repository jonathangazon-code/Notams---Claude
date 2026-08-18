Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization.Json
Imports System.Security.Cryptography
Imports System.Text
Imports ICAOModels.Interfaces

<DataContract>
<KnownType(GetType(AirplaneInfo))>
Public Class SignedAircraftLibrary

    <DataMember>
    Public Property Airplanes As List(Of IAirplaneInfo)
    <DataMember>
    Public Property Exponent As Byte()
    <DataMember>
    Public Property Modulus As Byte()
    <DataMember>
    Public Property SignedHash As Byte()

    ''' <summary>
    ''' This is the default constructor for the SignedAircraftLibrary class.
    ''' </summary>
    Public Sub New()

    End Sub

    Public Function GetHash() As Byte()
        Dim array = New Byte() {}
        If Airplanes Is Nothing Then
            Return array
        End If
        Dim ser = New DataContractJsonSerializer(GetType(List(Of IAirplaneInfo)))
        Dim ms = New MemoryStream()

        ser.WriteObject(ms, Airplanes)
        Dim json = Encoding.Default.GetString(ms.ToArray())

        Dim md5Hash As MD5 = MD5.Create()
        Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(json))

        Return data
    End Function
End Class
