# ScintelliVB

ScintelliVB is a small class that allows you to add Intellisense functionality to a project using ScintillaNET with Visual Basic.NET

## Features
- Uppercase first letter of keywords
- AutoIndentation
- AutoCompletion
- Intellisense (search in local/global variables and in namespaces)
- Call tips (overloads supported)
- Auto close block (Class, Sub, Function, If)
- Folding
- Multiple selections

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
MIT License

Copyright (c) [2020] [ShadowTzu]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

