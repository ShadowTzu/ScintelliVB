# ScintelliVB

ScintelliVB is a small class that allows you to add Intellisense functionality to a project using ScintillaNET with Visual Basic.NET

## Feature
- Uppercase first letter of keywords
- AutoIndentation
- AutoCompletion
- Intellisense (search in local/global variables and in namespace)
- Call tips (overloads supported)
- Auto close block (Class, Sub, Function, If)
- Folding
- Multiple selection

## Usage

```vb
Imports ScintillaNET
Imports ScintelliVB
'..
Private WithEvents TextArea As Scintilla
'..
TextArea = New ScintillaNET.Scintilla With {
         .Name = "TextArea"
}
TextArea.Dock = DockStyle.Fill
Me.Controls.Add(TextArea)
IntelliVB = New ScintelliVB.ScintelliVB(TextArea)
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://choosealicense.com/licenses/mit/)
