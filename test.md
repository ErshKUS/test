VBA-Web
=

VBA-Web (formerly Excel-REST) makes working with complex webservices and APIs easy with VBA on Windows and Mac. It includes support for authentication, automatically converting and parsing JSON, working with cookies and headers, and much more.

Getting started
-

- Download the [latest release (v4.0.12)](https://github.com/VBA-tools/VBA-Web/releases)
- To install/upgrade in an existing file, use `VBA-Web - Installer.xlsm`
- To start from scratch in Excel, `VBA-Web - Blank.xlsm` has everything setup and ready to go

For more details see the [Wiki](https://github.com/VBA-tools/VBA-Web/wiki)

### GetJSON Example Пример
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
```

## Синтаксис тегов

### Для элементов и полей
Элемент				| В Evolution (Старый)	| В Revolution (Новый)		| Пример (для Revolution)
-|-|-|-
Шаблоны				| Нет					| Нет						|
Поля ресурсов		| `[*field*]`			| `[[*field]]`				| `[[*pagetitle]]`
Дополнительные поля	| `[*templatevar*]`		| `[[*templatevar]]`		| `[[*tags]]`
Чанки				| `{{chunk }}`			| `[[$chunk]]`				| `[[$header]]`
Сниппеты			| `[[snippet]]`			| `[[snippet]]`				| `[[getResources]]`
Плагины				| Нет					| Нет						|
