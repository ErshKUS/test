VBA-Web
=
[![Build Status](https://travis-ci.org/laravel/framework.svg)](https://travis-ci.org/laravel/framework)
[![Total Downloads](https://poser.pugx.org/laravel/framework/d/total.svg)](https://packagist.org/packages/laravel/framework)
[![Latest Stable Version](https://poser.pugx.org/laravel/framework/v/stable.svg)](https://packagist.org/packages/laravel/framework)
[![Latest Unstable Version](https://poser.pugx.org/laravel/framework/v/unstable.svg)](https://packagist.org/packages/laravel/framework)
[![License](https://poser.pugx.org/laravel/framework/license.svg)](https://packagist.org/packages/laravel/framework)

VBA-Web (formerly Excel-REST) makes working with complex webservices and APIs easy with VBA on Windows and Mac. It includes support for authentication, automatically converting and parsing JSON, [zolla.50webs.com/support][working] with cookies and headers, and much more.

<img src="http://saahov.ru/assets/2011/06/github-fork-link.png" alt="Fork" />

[working]:http://zolla.50webs.com

Getting started
-

- Download the [latest release (v4.0.12)](https://github.com/VBA-tools/VBA-Web/releases)
- To install/upgrade in an existing file, use `VBA-Web - Installer.xlsm`
- To start from scratch in Excel, `VBA-Web - Blank.xlsm` has everything setup and ready to go
* Download the [latest release (v4.0.12)](https://github.com/VBA-tools/VBA-Web/releases)
* To install/upgrade in an existing file, use `VBA-Web - Installer.xlsm`
* To start from scratch in Excel, `VBA-Web - Blank.xlsm` has everything setup and ready to go

For more details see the [Wiki](https://github.com/VBA-tools/VBA-Web/wiki)

### GetJSON Example Пример
```vba
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

#Заголовок?

## Синтаксис тегов

### Для элементов и полей
Элемент				| В Evolution (Старый)	| В Revolution (Новый)		| Пример (для Revolution)
---|---|---|---
Шаблоны				| Нет					| Нет						|
Поля ресурсов		| `[*field*]`			| `[[*field]]`				| `[[*pagetitle]]`
Дополнительные поля	| `[*templatevar*]`		| `[[*templatevar]]`		| `[[*tags]]`
Чанки				| `{{chunk }}`			| `[[$chunk]]`				| `[[$header]]`
Сниппеты			| `[[snippet]]`			| `[[snippet]]`				| `[[getResources]]`
Плагины				| Нет					| Нет						|

### Примеры форматирования

```markdown
# Заголовок документа (обязателен, иначе будет отображаться имя файла)

Текст документа.

## Заголовок второго уровня

### Заголовок третьего уровня

#### Заголовок четвёртого уровня

* Первый элемент списка
* Второй элемент списка
* Третий элемент списка

1. Первый элемент нумерованного списка
2. Второй элемент нумерованного списка
3. Третий элемент нумерованного списка

> Текст цитаты

[текст для внешней ссылки](http://example.com/)

[[текст для внутренней ссылки|имя документа, на который будет ссылка]] — имя документа должно быть расширения.

**Текст, выделенный жирным**

_Текст, выделенный курсивом_

```

### Подсветка кода на страницах

Примеры выделения кода:

1). Выделенный код в той же строке (например, так рекомендуется выделять все теги MT, пути, адрес URL в примерах):
```
`<mt:Example/>`
```

2) Блок кода с обеих сторон заключается тремя символами:
```
```
```

3) Для блока кода можно указывать определённый синтаксис:
```
```perl
```

Возможные примеры синтаксиса: xml (рекомендуется для тегов MT), html, perl, php, bash, apache, mysql, sql, js, css.
