# ScintelliVB

ScintelliVB is a small class that allows you to add Intellisense functionality to a project using ScintillaNET with Visual Basic.NET


## Usage

```vb
Imports ScintillaNET
Imports ScintelliVB
'..
Private WithEvents TextArea As Scintilla
'..
TextArea = New Scintilla With {
         .Name = "TextArea"
}
TextArea.Dock = DockStyle.Fill
Me.Controls.Add(TextArea)
IntelliVB = New Scintelli(TextArea)
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://choosealicense.com/licenses/mit/)
