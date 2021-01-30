# aspjson

Classic ASP JSON Class Reading/Writing

[ASPJSON](https://www.aspjson.com) is a free to use project for generating and reading JSON data into a classic ASP object.
The class can be used for reading a string of JSON data as well as writing JSON output from an AJAX file.
Below are 2 simple examples of both.

## Read Json Example:

```
{
  "firstName": "John",
  "lastName" : "Smith",
  "age"      : 25,
  "address"  :
  {
    "streetAddress": "21 2nd Street",
    "city"         : "New York",
    "state"        : "NY",
    "postalCode"   : "10021"
  },
  "phoneNumbers":
  [
    {
        "type"  : "home",
        "number": "212 555-1234"
      },
      {
        "type"  : "fax",
        "number": "646 555-4567"
    }
  ]
}
```

```
<%
Set oJSON = New aspJSON

'Load JSON string
oJSON.loadJSON(jsonstring)

'Get single value
Response.Write oJSON.data("firstName") & "<br>"
Response.Write oJSON.data("address").item("streetAddress") & "<br>"

'Loop through collection
For Each phonenr In oJSON.data("phoneNumbers")
    Set this = oJSON.data("phoneNumbers").item(phonenr)
    Response.Write _
    this.item("type") & ": " & _
    this.item("number") & "<br>"
Next

'Update/Add value
oJSON.data("firstName") = "James"

'Return json string
Response.Write oJSON.JSONoutput()
%>
```


## Write Json Example:

```
{
  "familyName": "Smith",
  "familyMembers":
  [
    {
      "firstName": "John",
      "age": 41,
      "job": 
      {
        "function": "Webdeveloper",
        "salary": 70000
      }
    },
    {
      "firstName": "Suzan",
      "age": 38,
      "interests":
      [
        "Reading",
        "Tennis",
        "Painting"
      ]
    },
    {
      "firstName": "John Jr.",
      "age": 2.5
    }
  ]
}
```

```
<%
Set oJSON = New aspJSON

With oJSON.data

    .Add "familyName", "Smith"                      'Create value
    .Add "familyMembers", oJSON.Collection()

    With oJSON.data("familyMembers")

        .Add 0, oJSON.Collection()                  'Create unnamed object
        With .item(0)
            .Add "firstName", "John"
            .Add "age", 41

            .Add "job", oJSON.Collection()          'Create named object
            With .item("job")
                .Add "function", "Webdeveloper"
                .Add "salary", 70000
            End With
        End With


        .Add 1, oJSON.Collection()
        With .item(1)
            .Add "firstName", "Suzan"
            .Add "age", 38
            .Add "interests", oJSON.Collection()    'Create array
            With .item("interests")
                .Add 0, "Reading"
                .Add 1, "Tennis"
                .Add 2, "Painting"
            End With
        End With

        .Add 2, oJSON.Collection()
        With .item(2)
            .Add "firstName", "John Jr."
            .Add "age", 2.5
        End With

    End With

End With

Response.Write oJSON.JSONoutput()                   'Return json string
```


## Changelog

### Version 1.19 January 2021
* Fix for locale decimal format settings when parsing json string

### Version 1.18 August 2020
* Decimal output comma/dot fix

### Version 1.17 Februari 2014
* Efficiency improvement large data

### Version 1.15 Februari 2014
* Trailing tabs fixed

### Version 1.14 Februari 2014
* Colon value within string bug fixed

### Version 1.13 December 2013
* Encoded data fix
* Now possible to load directly from a URL. For example: oJSON.loadJSON("http://www.aspjson.com/jsonstream.asp")

### Version 1.12 June 2013
* vbCrLf fix
* Compatible with Option Explicit
* JSON escape characters
