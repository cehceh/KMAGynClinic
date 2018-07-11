Imports System.Xml
Imports System.IO
Imports System.Configuration
Public Class Form4

    Private Sub ChangeMyAppScopedSetting(ByVal newValue As String)
        Dim config As System.Configuration.Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Dim xmlDoc As New XmlDocument()

        ' Load an XML file into the XmlDocument object.
        Try
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(config.FilePath.Trim)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Dim i, j, k, l As Int32
        Try
            For i = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Count - 1
                For j = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Count - 1
                    For k = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Count - 1
                        If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).Name = "setting" Then
                            If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).Attributes.Item(0).Value = "MyAppScopedSetting" Then
                                For l = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Count - 1
                                    If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Item(l).Name = "value" Then
                                        xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Item(l).InnerText = newValue
                                    End If
                                Next l
                            End If
                        End If
                    Next k
                Next j
            Next i
            xmlDoc.Save(config.FilePath.Trim)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ChangeMyAppScopedSetting_2(ByVal newValue As String)
        Dim config As System.Configuration.Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Dim xmlDoc As New XmlDocument()

        ' Load an XML file into the XmlDocument object.
        Try
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(config.FilePath.Trim)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Dim i, j, k, l As Int32
        Try
            For i = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Count - 1
                For j = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Count - 1
                    For k = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Count - 1
                        If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).Name = "setting" Then
                            If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).Attributes.Item(0).Value = "MySysFile" Then
                                For l = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Count - 1
                                    If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Item(l).Name = "value" Then
                                        xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Item(l).InnerText = newValue
                                    End If
                                Next l
                            End If
                        End If
                    Next k
                Next j
            Next i
            xmlDoc.Save(config.FilePath.Trim)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ChangeMyAppScopedSetting_3(ByVal newValue As String)
        Dim config As System.Configuration.Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        Dim xmlDoc As New XmlDocument()

        ' Load an XML file into the XmlDocument object.
        Try
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(config.FilePath.Trim)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Dim i, j, k, l As Int32
        Try
            For i = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Count - 1
                For j = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Count - 1
                    For k = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Count - 1
                        If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).Name = "setting" Then
                            If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).Attributes.Item(0).Value = "Date_Days" Then
                                For l = 0 To xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Count - 1
                                    If xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Item(l).Name = "value" Then
                                        xmlDoc.GetElementsByTagName("applicationSettings").Item(i).ChildNodes.Item(j).ChildNodes.Item(k).ChildNodes.Item(l).InnerText = newValue
                                    End If
                                Next l
                            End If
                        End If
                    Next k
                Next j
            Next i
            xmlDoc.Save(config.FilePath.Trim)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        ChangeMyAppScopedSetting(TextBox1.Text)
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        ChangeMyAppScopedSetting_2(TextBox3.Text)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        ChangeMyAppScopedSetting_3(TextBox4.Text)
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Enabled = False
        Button1.Enabled = False
        TextBox3.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button7.Enabled = False

        lblSettings.Text = My.Settings.MyAppScopedSetting
        Label5.Text = My.Settings.Date_Days
        TextBox3.Text = My.Settings.MySysFile
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        My.Settings.Reload()
        MsgBox("MyAppScopeSettings value is" + " " + My.Settings.MyAppScopedSetting)
        lblSettings.Text = My.Settings.MyAppScopedSetting
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        My.Settings.Reload()
        MsgBox("MyAppScopeSettings value is" + " " + My.Settings.MySysFile)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        My.Settings.Reload()
        MsgBox("MyAppScopeSettings value is" + " " + My.Settings.Date_Days)
        Label5.Text = My.Settings.Date_Days
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text = "FSL_hggi1981923" Then   'textbox2.text must be equal your password
            TextBox1.Enabled = True
            Button1.Enabled = True
            TextBox3.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
            Button5.Enabled = True
            Button7.Enabled = True
            Button6.Enabled = True
            TextBox4.Enabled = True

            TextBox3.Text = My.Settings.MySysFile
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim f1 As New Form1
        Me.Close()  'Or .hide() 
        f1.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '"C:\Windows\System32\Python"
        If Not File.Exists(TextBox3.Text) Then
            File.Create(TextBox3.Text)
            MsgBox("Done")
        Else
            MsgBox("Already Exists")
        End If
        'File.Create(TextBox3.Text)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If Directory.Exists(TextBox3.Text) Then
            MsgBox("Directory Already Exists")
            Exit Sub
        Else
            Directory.CreateDirectory(TextBox3.Text)
            MsgBox("Done")
        End If
    End Sub
End Class
