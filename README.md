# aspjson

Classic ASP JSON Class Reading/Writing

ASPJSON is a free to use project for generating and reading JSON data into a classic ASP object.
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
