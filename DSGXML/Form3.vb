Public Class Form3

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "Please Select a XML file"
        OpenFileDialog1.InitialDirectory = "C:\Temp"
        OpenFileDialog1.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
        OpenFileDialog1.FilterIndex = 1

        OpenFileDialog1.ShowDialog()


    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

        TextBox1.Text = OpenFileDialog1.FileName.ToString()
        Debug.Print(TextBox1.Text)
        Me.Hide()

        Form1.Show()


    End Sub
End Class