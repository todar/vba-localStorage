Attribute VB_Name = "LocalStorage_Tests"
''
' Testing for the LocalStorage module.
'
' @author Robert Todar <robert@roberttodar.com>
''
Option Explicit

''
' Main testing for LocalStorage
' Currently it is just a visual test of the data
' but could get better assertions in here as well.
''
Public Sub Test_LocalStorage()
    ' Clear LocalStorage for testing
    LocalStorage.Clear

    ' Add Some test data
    LocalStorage.SetItem "name", "Robert"
    LocalStorage.SetItem "age", 32
    
    ' Simply visualise that the data is modeled correctly.
    Debug.Print LocalStorage.ToString
End Sub

''
' This is just a visual test run after closing the workbook
' making sure the values persisted.
''
Public Sub Test_StorageExistsAfterClose()
    Debug.Print LocalStorage.ToString
End Sub
