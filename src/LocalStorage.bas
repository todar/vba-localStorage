Attribute VB_Name = "LocalStorage"
''
' LocalStorage is a way of saving and retreiving values in a key value format.
' This mimics LocalStorage found in browsers, and allows for an easy way of
' storing variables without the need of using a worksheet or some external
' file.
'
' @author Robert Todar <robert@roberttodar.com>
''
Option Explicit

''
' Created this property to get intellisence
' for ThisWorkbook.documentProperties
''
Private Property Get properties() As documentProperties
    Set properties = ThisWorkbook.CustomDocumentProperties
End Property

''
' Read from Local Storage.
' No errors are raised if it is not found, simply will return
' an Optional fallback value.
''
Public Function GetItem(ByVal key As String, Optional ByVal fallback As String = vbNullString) As String
On Error GoTo catch
    GetItem = properties.Item(key)
Exit Function
catch:
    GetItem = fallback
End Function

''
' Write to the Local Storage
' To mimic the browser, everything will be converted to a string.
''
Public Sub SetItem(ByVal key As String, ByVal value As String)
    On Error GoTo catch
    ' Error will occur if the item already exists.
    ' Always removing item first will make sure it doesn't error out.
    LocalStorage.RemoveItem key
    properties.Add key, False, msoPropertyTypeString, value
    ThisWorkbook.Save
Exit Sub
catch:
End Sub

''
' Remove a specified value using the key.
' No error thrown if value doesn't exist.
''
Public Sub RemoveItem(ByVal key As String)
    On Error GoTo catch
    properties.Item(key).Delete
catch:
End Sub

''
' Clear all values from Local Storage.
''
Public Sub Clear()
    Dim prop As DocumentProperty
    For Each prop In properties
        prop.Delete
    Next prop
End Sub

''
' To see the LocalStorage in an easier way, this will print out
' the Storage Object as JSON
''
Public Function ToString() As String
    Dim prop As DocumentProperty
    For Each prop In properties
        ToString = ToString & IIf(ToString <> "", "," & vbNewLine, "") & _
                   "  """ & prop.Name & """: """ & prop.value & """"
    Next prop
    
    ToString = IIf(ToString = "", "{}", "{" & vbNewLine & ToString & vbNewLine & "}")
End Function
