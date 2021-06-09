# VBA Local Storage 🗃

`LocalStorage` is a way of saving and retreiving values in a key value format.

This VBA implementation mimics LocalStorage found in browsers, and allows for an easy way of storing variables without the need of using a worksheet or some external file.

<a href="https://www.buymeacoffee.com/todar" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" style="height: 51px !important;width: 217px !important;" ></a>

---

## Other Helpful Resources

- [www.roberttodar.com](https://www.roberttodar.com/) About me and my background and some of my other projects.
- [Style Guide](https://github.com/todar/VBA-Style-Guide) A guide for writing clean VBA code. Notes on how to take notes =)
- [Strings](https://github.com/todar/VBA-Strings) String function library. `ToString`, `Inject`, `StringSimilarity`, and more.
- [Analytics](https://github.com/todar/VBA-Analytics) Way of tracking code analytics and metrics. Useful when multiple users are running code within a shared network.
- [Userform EventListener](https://github.com/todar/VBA-Userform-EventListener) Listen to events such as `mouseover`, `mouseout`, `focus`, `blur`, and more.

---

## List of Available Methods

| Function Name          | Description                                             |
| :--------------------- | :------------------------------------------------------ |
| `GetItem`         | Read from Local Storage. |
| `SetItem` | Write to Local Storage. Optional fallback value as well. |
| `RemoveItem` | Remove a specified value using the key. |
| `Clear` | Remove all items from Local Storage. |
| `ToString` | Return a JSON string of all key/values in Local Storage. |

---

## Example of using Local Storage

```vba
''
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
    Debug.Print LocalStorage.GetItem("name")
    Debug.Print LocalStorage.GetItem("Null", "Fallback Value")

    ' Print all storage in JSON string
    Debug.Print LocalStorage.ToString
End Sub
```