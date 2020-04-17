Imports System.ComponentModel
Imports System.Reflection
Imports ScintelliVB
Imports ScintillaNET

Public Class FormTest
    Private DefaultScript_Text As String = "Imports Microsoft.VisualBasic" & vbCrLf & "Imports System.Math" & vbCrLf & vbCrLf & "Public Class {0}" & vbCrLf & vbCrLf & vbTab & "Sub New()" & vbCrLf & vbTab & vbTab & " MyBase.New(""{0}"")" & vbCrLf & vbTab & "End Sub" & vbCrLf & vbCrLf & vbTab & "public overrides sub start()" & vbCrLf & vbTab & vbTab & vbCrLf & vbTab & "End Sub" & vbCrLf & vbCrLf & vbTab & "Public Overrides Sub Update()" & vbCrLf & vbTab & vbTab & vbCrLf & vbTab & "End Sub" & vbCrLf & vbCrLf & vbTab & "Public Overrides Sub Destroy()" & vbCrLf & vbTab & vbTab & vbCrLf & vbTab & "End Sub" & vbCrLf & "End Class"

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

    Private Sub TimerText_Tick(sender As Object, e As EventArgs) Handles TimerText.Tick
        TextBox_0.Text = IntelliVB.Text0
        TextBox_1.Text = IntelliVB.Text1
        TextBox_2.Text = IntelliVB.Text2
        TextBox_3.Text = IntelliVB.Text3
        TextBox_4.Text = IntelliVB.Text4
    End Sub
End Class
