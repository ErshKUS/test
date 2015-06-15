VBA-Web
=======

VBA-Web (formerly Excel-REST) makes working with complex webservices and APIs easy with VBA on Windows and Mac. It includes support for authentication, automatically converting and parsing JSON, working with cookies and headers, and much more.

Getting started
---------------

- Download the [latest release (v4.0.12)](https://github.com/VBA-tools/VBA-Web/releases)
- To install/upgrade in an existing file, use `VBA-Web - Installer.xlsm`
- To start from scratch in Excel, `VBA-Web - Blank.xlsm` has everything setup and ready to go

For more details see the [Wiki](https://github.com/VBA-tools/VBA-Web/wiki)

Upgrading
---------

To upgrade from Excel-REST to VBA-Web, follow the [Upgrading Guide](https://github.com/VBA-tools/VBA-Web/wiki/Upgrading-from-v3.*-to-v4.*)

Note: XML support has been temporarily removed from VBA-Web while parser issues for Mac are resolved.
XML support is still possible on Windows, follow [these instructions](https://github.com/VBA-tools/VBA-Web/wiki/XML-Support-in-4.0) to use a custom formatter.

Examples
-------

The following examples demonstrate using the Google Maps API to get directions between two locations.

### GetJSON Example
```VBA
Function GetDirections(Origin As String, Destination As String) As String
    ' Create a WebClient for executing requests
    ' and set a base url that all requests will be appended to
    Dim MapsClient As New WebClient
    MapsClient.BaseUrl = "https://maps.googleapis.com/maps/api/"
    
    ' Use GetJSON helper to execute simple request and work with response
    Dim Resource As String
    Dim Response As WebResponse
    
    Resource = "directions/json?" & _
        "origin=" & Origin & _
        "&destination=" & Destination & _
        "&sensor=false"
    Set Response = MapsClient.GetJSON(Resource)
    
    ' => GET https://maps.../api/directions/json?origin=...&destination=...&sensor=false
    
    ProcessDirections Response
End Function

Public Sub ProcessDirections(Response As WebResponse)
    If Response.StatusCode = WebStatusCode.Ok Then
        Dim Route As Dictionary
        Set Route = Response.Data("routes")(1)("legs")(1)

        Debug.Print "It will take " & Route("duration")("text") & _
            " to travel " & Route("distance")("text") & _
            " from " & Route("start_address") & _
            " to " & Route("end_address")
    Else
        Debug.Print "Error: " & Response.Content
    End If
End Sub
```
