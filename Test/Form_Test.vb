Imports System.ComponentModel
Imports System.Reflection
Imports ScintelliVB
Imports ScintillaNET

Public Class Form_Test
    Private DefaultScript_Text As String = "Imports Microsoft.VisualBasic" & vbCrLf & "Imports System.Math" & vbCrLf & vbCrLf & "Public Class {0}" & vbCrLf & vbCrLf & vbTab & "Sub New()" & vbCrLf & vbTab & vbTab & " MyBase.New(""{0}"")" & vbCrLf & vbTab & "End Sub" & vbCrLf & vbCrLf & vbTab & "Public Overrides Sub Start()" & vbCrLf & vbTab & vbTab & vbCrLf & vbTab & "End Sub" & vbCrLf & vbCrLf & vbTab & "Public Overrides Sub Update()" & vbCrLf & vbTab & vbTab & vbCrLf & vbTab & "End Sub" & vbCrLf & vbCrLf & vbTab & "Public Overrides Sub Destroy()" & vbCrLf & vbTab & vbTab & vbCrLf & vbTab & "End Sub" & vbCrLf & "End Class"

    Private WithEvents TextArea As Scintilla
    Private IntelliVB As ScintelliVB.ScintelliVB

    Private Sub Form_Test_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextArea = New ScintillaNET.Scintilla With {
                 .Name = "TextArea"
        }
        Me.SplitContainer1.Panel2.Controls.Add(TextArea)
        TextArea.Text = String.Format(DefaultScript_Text, "Test")
        TextArea.Dock = DockStyle.Fill

        IntelliVB = New ScintelliVB.ScintelliVB(TextArea)

    End Sub



End Class
