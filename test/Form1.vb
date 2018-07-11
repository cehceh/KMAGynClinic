Option Strict On
Option Explicit On

Imports System.Data.OleDb
Imports System.IO
Imports System.Xml
Imports ApplicationEnhancement.Expiry



Public Class Form1

    Inherits System.Windows.Forms.Form

    Dim cs As String = "provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923"
    Dim conn As New OleDbConnection(cs)
    Dim cmd As New OleDbCommand
    Public dr As OleDbDataReader

    Dim f2 As Form2

    'To make the App. to open a fixed times
    Private WithEvents usage As ApplicationUsage
    Private _maxTimes As Integer = 100000
    Private _usageLimitExceeded As Boolean = False

    ''##from Daniweb.com  https://www.daniweb.com/programming/software-development/threads/339730/visual-basic-code-to-setup-your-program-to-expire-after-30-days#post1972097
    ''##Very Useful code for make expiration date 
    Public Function DateGood(NumDays As Integer) As Boolean
        'The purpose of this module is to allow you to place a time
        'limit on the unregistered use of your shareware application.
        'This module can not be defeated by rolling back the system clock.
        'Simply call the DateGood function when your application is first
        'loading, passing it the number of days it can be used without
        'registering.
        'Ex: If DateGood(30)=False Then
        ' Cripple Application
        ' End if
        'Register Parameters:
        ' CRD: Current Run Date
        ' LRD: Last Run Date
        ' FRD: First Run Date
        Dim TmpCRD As Date
        Dim TmpLRD As Date
        Dim TmpFRD As Date
        TmpCRD = CDate(Format(Now, ("dd/MM/yyyy").ToString))
        TmpLRD = CDate(GetSetting(Application.ExecutablePath, "Param", "LRD", "11/9/2017"))
        TmpFRD = CDate(GetSetting(Application.ExecutablePath, "Param", "FRD", "1/9/2017"))
        DateGood = False
        'If this is the applications first load, write initial settings
        'to the register
        If TmpLRD = CDate("11/9/2017") Then
            SaveSetting(Application.ExecutablePath, "Param", "LRD", CType(TmpCRD, String))
            SaveSetting(Application.ExecutablePath, "Param", "FRD", CType(TmpCRD, String))
        End If
        'Read LRD and FRD from register
        TmpLRD = CDate(GetSetting(Application.ExecutablePath, "Param", "LRD", "11/9/2017"))
        TmpFRD = CDate(GetSetting(Application.ExecutablePath, "Param", "FRD", "1/9/2017"))
        If TmpFRD > TmpCRD Then 'System clock rolled back
            DateGood = False
        ElseIf Now > DateAdd("d", NumDays, TmpFRD) Then 'Expiration expired
            DateGood = False
        ElseIf TmpCRD > TmpLRD Then 'Everything OK write New LRD date
            SaveSetting(Application.ExecutablePath, "Param", "LRD", CType(TmpCRD, String))
            DateGood = True
        ElseIf TmpCRD = CDate(Format(TmpLRD, "dd/MM/yyyy")) Then
            DateGood = True
        Else
            DateGood = False
        End If
    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        DTPickerNow()
        DateTimePicker1.Value = Now
        DTPicker.Value = Now

        LoadPicture()
        Me.AutoScroll = True
        FillAuto()
        txtNo.Select()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '_maxTimes = CInt(Val(My.Settings.MyAppScopedSetting))
        'DateGood(CInt(Val(My.Settings.Date_Days)))

        ' Initialize the variable "usage" by using
        ' the NEW keyword along with the number
        ' of maximum "hits" you want to allow:
        'usage = New ApplicationUsage(CInt(_maxTimes))
        usage = New ApplicationUsage(_maxTimes)

        ' Now check the usage. If the usage has been
        ' exceeded, the "MaximumExceeded" event will
        ' be raised:
        usage.CheckUsage()

        If _usageLimitExceeded Then
            MessageBox.Show(String.Format("The maximum usage of " & vbCrLf &
                               "{0: n0} times has been exceeded." & vbCrLf &
                               "For full version Call 01067174141 Or 01149003573",
                                _maxTimes), "Cannot Continue",
                               MessageBoxButtons.OK)
            'Exit Sub
            Close()
            End

        Else
            ' Just for demonstration here, if the
            ' maximum has not been exceeded then
            ' I'll just show the quantity of times
            ' this program has been run.

            Label46.Text = String.Format("{0:n0} Of " &
                           _maxTimes, usage.UsageQuantity)      'CType(usage.UsageQuantity, String)
            'MessageBox.Show(String.Format("Usage Quantity : {0:n0} Of " &
            'CInt(Val(My.Settings.MyAppScopedSetting)), usage.UsageQuantity))
        End If

        'Dim SysFile As String = "C:\Drivers\Video\AMD1\CCC\localdata.ini"
        'SysFile = My.Settings.MySysFile

        'If Not File.Exists(SysFile) Then
        '    'MsgBox("Call 01067174141 OR 01149003573", MsgBoxStyle.YesNo, "Unregistered Application")
        '    If MsgBox("Call 01067174141 OR 01149003573",
        '              MsgBoxStyle.YesNo, "Unregistered Application") = vbNo Then

        '        End

        '    Else
        '        Me.Hide()
        '        f4.ShowDialog()
        '    End If
        '    Close()
        '    End
        'End If

        ''##from Daniweb.com  https://www.daniweb.com/programming/software-development/threads/339730/visual-basic-code-to-setup-your-program-to-expire-after-30-days#post1972097
        '##Set timer And test
        '##This code for Expiration date it is very useful
        'If Not DateGood(10) Then  ''30 or any number for the expiration date
        '    MsgBox("Trial Period Expired!" & vbCrLf &
        '              "Call 01067174141", MsgBoxStyle.OkOnly,
        '              "Unregistered Date")
        '    Close()
        '    End

        'End If

    End Sub

    Private Sub _
        usage_MaximumExceeded(sender As Object,
                              e As System.EventArgs) _
                              Handles usage.MaximumExceeded

        _usageLimitExceeded = True

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        RDXmlPatNames()
        RDXmlPatNames1()
        RDXmlPatNames2()
        RDXmlDrugs()
        RDXmlDiaInter()
        RDXmlInv()
        RDXmlInv2()
        RDXmlInvRes()
        RDXmlPlan()

        If e.CloseReason = CloseReason.UserClosing Then
            usage.Save()
        End If
        End
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Location = New Point(0, 0)
        Me.Size = Screen.PrimaryScreen.WorkingArea.Size

        ShowPatTable()
        GynDisabled()
        Gyn2Disabled()

    End Sub
    ''##  https://forums.asp.net/t/1429696.aspx?Check+if+DBNull+value+and+replace+it+in+VB+NET
    ''##When you put this function in the save button click you will notice the difference
    ''##Till now it is useful ( Null value solution )
    Public Function CheckNull(ByVal fieldValue As String) As String
        If fieldValue.Equals(DBNull.Value) Then Return "" Else
        If fieldValue = "N/A" Then
            Return "value for N/A"
        Else
            Return ""
        End If
    End Function
    ''##To refresh 
    Private Sub loaddata()
        Me.Controls.Clear()
        InitializeComponent()
        Form1_Load(Nothing, Nothing)

        Me.Location = New Point(0, 0)
        Me.Size = Screen.PrimaryScreen.WorkingArea.Size
        Me.AutoScroll = True

        FillAuto()
        txtNo.Select()
        DTPickerNow()
        LoadPicture()
        DateTimePicker1.Value = Now

        GynDisabled()
        Gyn2Disabled()

    End Sub

    Sub DTPickerNow()
        Trace.WriteLine("DTPickerNow STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        DTPicker.Value = Now
        DTPickerMns.Value = Now
        DTPickerLMP.Value = Now
        DTPickerEDD.Value = Now
        DTPickerAtt.Value = Now

        Trace.WriteLine("DTPickerNow FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Sub

    Private Sub FillAuto()
        txtNo.Text = GetAutonumber("Pat", "Patient_no")
        txtVis1.Text = GetAutonumber("Gyn", "Vis_no")
        txtVis.Text = GetAutonumber("Gyn2", "Vis_no")

    End Sub

    Private Sub LoadPicture()
        'Dim PicPath As String = "D:\KMAClinic\Photos\GynClinic.png"
        'PicPath = My.Settings.PicFilePath
        Dim filename As String = Path.Combine(Application.StartupPath, "sabah2.png") 'Path.GetFileName("\DrAhEssmat.png")
        Me.PictureBox1.Image = Image.FromFile(filename)
    End Sub

    Function GetTable(SelectCommand As String) As DataTable
        Trace.WriteLine("GetTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Dim cmd As New OleDbCommand("", conn)
        Try
            Dim Data_table1 As New DataTable
            If conn.State = ConnectionState.Closed Then conn.Open()
            cmd.CommandText = SelectCommand
            Data_table1.Load(cmd.ExecuteReader())
            Return Data_table1
        Catch ex As Exception
            MsgBox(ex.Message)
            Return New DataTable
        Finally
            If conn.State = ConnectionState.Open Then conn.Close()
        End Try
        Trace.WriteLine("GetTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Function

    Function GetAutonumber(TableName As String, ColumnName As String) As String
        Trace.WriteLine("GetAutonumber STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Dim Str As String
        Str = "select max ( " & ColumnName & " ) + 1 from " & TableName
        Dim Data_table2 As New DataTable
        Data_table2 = GetTable(Str)
        Dim AutoNum As String
        If Data_table2.Rows(0)(0) Is DBNull.Value Then
            AutoNum = "1"
        Else
            AutoNum = CType(Data_table2.Rows(0)(0), String)
        End If
        Return AutoNum
        Trace.WriteLine("GetAutonumber FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Function

    Sub ClearGyn()
        Trace.WriteLine("ClearGyn STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        txtVis1.Text = ""
        txtA.Text = ""
        txtG.Text = ""
        txtP.Text = ""
        chbxNVD.Checked = False
        chbxCS.Checked = False
        txtGA.Text = ""
        txtElapsed.Text = ""

        DTPickerMns.Value = Now
        DTPickerLMP.Value = Now
        DTPickerEDD.Value = Now

        cbxLD.ResetText()
        cbxLC.ResetText()
        cbxMedH1.ResetText()
        cbxMedH2.ResetText()
        cbxMedH3.ResetText()
        cbxSurH1.ResetText()
        cbxSurH2.ResetText()
        cbxSurH3.ResetText()
        cbxGynH1.ResetText()
        cbxGynH2.ResetText()
        cbxGynH3.ResetText()
        cbxDrugH1.ResetText()
        cbxDrugH2.ResetText()
        cbxDrugH3.ResetText()
        chbxGyn.ResetText()
        txtVis1.Text = GetAutonumber("Gyn", "Vis_no")

        Trace.WriteLine("ClearGyn FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Sub

    Sub ClearGyn2()
        Trace.WriteLine("ClearGyn STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        txtVis.Text = ""
        cbxGL.Text = ""
        cbxPuls.Text = ""
        cbxBP.Text = ""
        cbxWeight.Text = ""
        cbxBodyBuilt.Text = ""
        cbxChtH.Text = ""
        cbxHdNe.Text = ""
        cbxExt.Text = ""
        cbxFunL.Text = ""
        cbxScars.Text = ""
        cbxEdema.Text = ""
        cbxUS.Text = ""
        txtAmount.Text = ""
        DTPickerAtt.Value = Now
        txtVis.Text = GetAutonumber("Gyn2", "Vis_no")

        Trace.WriteLine("ClearGyn FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Sub

    Private Sub btnclear_Click(sender As Object, e As EventArgs) Handles btnclear.Click
        Trace.WriteLine("btnclear_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Me.ListBox1.Items.Clear()
        Me.ListBox2.Items.Clear()
        Me.ListBox3.Items.Clear()

        cbxSearch.Text = ""
        cbxPatName.Text = ""
        cbxAddress.Text = ""
        DTPicker.Value = Now
        txtAge.Text = ""
        txtPhone.Text = ""
        cbxHusband.Text = ""

        TextBox1.Text = ""
        TextBox2.Text = ""
        ''##This line must come after the 'for loop' because for loop erase every textbox in the form
        ''##And the autonumber come after that
        txtNo.Text = GetAutonumber("Pat", "Patient_no")
        txtNo.Select()
        ClearGyn()
        ClearGyn2()

        GynDisabled()
        Gyn2Disabled()
        btnL.Enabled = True
        btnB.Enabled = True
        Trace.WriteLine("btnclear_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowGyn2Table()

        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "SELECT * FROM Gyn2 WHERE Vis_no=@Vis_no"
                cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVis.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtVis.Text = dt.Rows(0).Item("Vis_no").ToString
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        cbxGL.Text = dt.Rows(0).Item("GL").ToString
                        cbxPuls.Text = dt.Rows(0).Item("Pls").ToString
                        cbxBP.Text = dt.Rows(0).Item("BP").ToString
                        cbxWeight.Text = dt.Rows(0).Item("Wt").ToString
                        cbxBodyBuilt.Text = dt.Rows(0).Item("BdBt").ToString
                        cbxChtH.Text = dt.Rows(0).Item("ChtH").ToString
                        cbxHdNe.Text = dt.Rows(0).Item("HdNe").ToString
                        cbxExt.Text = dt.Rows(0).Item("Ext").ToString
                        cbxFunL.Text = dt.Rows(0).Item("FunL").ToString
                        cbxScars.Text = dt.Rows(0).Item("Scrs").ToString
                        cbxEdema.Text = dt.Rows(0).Item("Edm").ToString
                        cbxUS.Text = dt.Rows(0).Item("US").ToString
                        txtAmount.Text = dt.Rows(0).Item("Amount").ToString
                        DTPickerAtt.Text = dt.Rows(0).Item("AttDt").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub ShowGyn2PatTable()

        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Gyn2 WHERE Patient_no=@Patient_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))

                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtVis.Text = dt.Rows(0).Item("Vis_no").ToString
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        cbxGL.Text = dt.Rows(0).Item("GL").ToString
                        cbxPuls.Text = dt.Rows(0).Item("Pls").ToString
                        cbxBP.Text = dt.Rows(0).Item("BP").ToString
                        cbxWeight.Text = dt.Rows(0).Item("Wt").ToString
                        cbxBodyBuilt.Text = dt.Rows(0).Item("BdBt").ToString
                        cbxChtH.Text = dt.Rows(0).Item("ChtH").ToString
                        cbxHdNe.Text = dt.Rows(0).Item("HdNe").ToString
                        cbxExt.Text = dt.Rows(0).Item("Ext").ToString
                        cbxFunL.Text = dt.Rows(0).Item("FunL").ToString
                        cbxScars.Text = dt.Rows(0).Item("Scrs").ToString
                        cbxEdema.Text = dt.Rows(0).Item("Edm").ToString
                        cbxUS.Text = dt.Rows(0).Item("US").ToString
                        txtAmount.Text = dt.Rows(0).Item("Amount").ToString
                        DTPickerAtt.Text = dt.Rows(0).Item("AttDt").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub ShowGynTable()
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "SELECT * FROM Gyn WHERE Patient_no=@Patient_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        txtVis1.Text = dt.Rows(0).Item("Vis_no").ToString
                        txtG.Text = dt.Rows(0).Item("G").ToString
                        txtP.Text = dt.Rows(0).Item("P").ToString
                        txtA.Text = dt.Rows(0).Item("A").ToString
                        chbxNVD.Checked = CBool(dt.Rows(0).Item("NVD").ToString)
                        chbxCS.Checked = CBool(dt.Rows(0).Item("CS").ToString)
                        cbxHPOC.Text = dt.Rows(0).Item("HPOC").ToString
                        cbxLD.Text = dt.Rows(0).Item("LD").ToString
                        cbxLC.Text = dt.Rows(0).Item("LC").ToString
                        DTPickerMns.Text = dt.Rows(0).Item("MNSDate").ToString
                        DTPickerLMP.Text = dt.Rows(0).Item("LMPDate").ToString
                        DTPickerEDD.Text = dt.Rows(0).Item("EDDDate").ToString
                        txtElapsed.Text = dt.Rows(0).Item("ElapW").ToString
                        txtGA.Text = dt.Rows(0).Item("GAW").ToString
                        cbxMedH1.Text = dt.Rows(0).Item("MedH1").ToString
                        cbxMedH2.Text = dt.Rows(0).Item("MedH2").ToString
                        cbxMedH3.Text = dt.Rows(0).Item("MedH3").ToString
                        cbxSurH1.Text = dt.Rows(0).Item("SurH1").ToString
                        cbxSurH2.Text = dt.Rows(0).Item("SurH2").ToString
                        cbxSurH3.Text = dt.Rows(0).Item("SurH3").ToString
                        cbxGynH1.Text = dt.Rows(0).Item("GynH1").ToString
                        cbxGynH2.Text = dt.Rows(0).Item("GynH2").ToString
                        cbxGynH3.Text = dt.Rows(0).Item("GynH3").ToString
                        cbxDrugH1.Text = dt.Rows(0).Item("DrugH1").ToString
                        cbxDrugH2.Text = dt.Rows(0).Item("DrugH2").ToString
                        cbxDrugH3.Text = dt.Rows(0).Item("DrugH3").ToString
                        chbxGyn.Checked = CBool(dt.Rows(0).Item("Gyn").ToString)
                    End If
                End Using
            End Using
        End Using

    End Sub

    Sub ShowPatTable()
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con                               ''##

                cmd.CommandText = "SELECT * FROM Pat WHERE Patient_no=@Patient_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        'txtName.Text = dt.Rows(0).Item("Name")
                        cbxPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        cbxHusband.Text = dt.Rows(0).Item("HusName").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        conn.Open()
        txtNo.ResetText()
        cbxPatName.ResetText()
        cbxAddress.ResetText()
        DTPicker.ResetText()
        txtAge.ResetText()
        txtPhone.ResetText()
        cbxHusband.ResetText()
        Dim str As String = "SELECT * FROM [Pat] " &
        "WHERE Patient_no LIKE '%" & Me.cbxSearch.Text & "%' " &
        "ORDER BY Patient_no DESC"
        Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        dr = cmd.ExecuteReader
        While dr.Read
            txtNo.Text = dr("Patient_no").ToString
            cbxPatName.Text = dr("Name").ToString
            cbxAddress.Text = dr("Address").ToString
            DTPicker.Text = dr("Birthdate").ToString
            txtAge.Text = dr("Age").ToString
            txtPhone.Text = dr("Phone").ToString
            cbxHusband.Text = dr("HusName").ToString
        End While
        conn.Close()
    End Sub

    Private Sub cbxSearch_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxSearch.MouseClick
        If rdoName.Checked Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSearch.Items.AddRange(cbElements)
        ElseIf rdoHus.Checked Then
            '' Read the XML file from disk only once
            Dim xDoc1 = XElement.Load(Application.StartupPath + "\PatNames1.xml")
            '' Parse the XML document only once
            Dim cbElements1 = xDoc1.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSearch.Items.AddRange(cbElements1)
        End If
        btnF.Enabled = True
        btnL.Enabled = True
    End Sub

    Private Sub cbxSearch_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxSearch.Validating
        If rdoName.Checked Then
            SearchName()
        ElseIf rdoID.Checked Then
            SearchID()
        ElseIf rdoHus.Checked Then
            SearchHusband()
        ElseIf rdoPhone.Checked Then
            SearchPhone()
        End If
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        GynDisabled()
        Gyn2Disabled()
        btnF.Enabled = True
        btnL.Enabled = True

        ''##https://social.msdn.microsoft.com/Forums/vstudio/en-US/b2a15b26-6d51-49d5-81cf-20fef70e8316/when-datetimepicker-value-changed-this-error-occured?forum=vbgeneral
        operations.ToAgeString(DTPicker.Value)
        txtAge.Text = operations.Formatted
    End Sub

    Private Sub cbxSearch_TextChanged(sender As Object, e As EventArgs) Handles cbxSearch.TextChanged
        btnF.Enabled = True
        btnL.Enabled = True
    End Sub

    Private Sub txtNo_MouseClick(sender As Object, e As MouseEventArgs) Handles txtNo.MouseClick

    End Sub

    '## https://stackoverflow.com/questions/31561090/finding-records-in-a-database-using-a-textbox-and-search-button-in-vb-net
    Private Sub SearchName()
        'conn.Open()
        'txtNo.ResetText()
        'cbxPatName.ResetText()
        'cbxAddress.ResetText()
        'DTPicker.ResetText()
        'txtAge.ResetText()
        'txtPhone.ResetText()
        'cbxHusband.ResetText()
        'Dim str As String = "SELECT * FROM [Pat] " &
        '"WHERE Name LIKE '%" & Me.cbxSearch.Text & "%' " &
        '"ORDER BY Patient_no DESC"
        'Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        'dr = cmd.ExecuteReader
        'While dr.Read
        '    txtNo.Text = dr("Patient_no").ToString
        '    cbxPatName.Text = dr("Name").ToString
        '    cbxAddress.Text = dr("Address").ToString
        '    DTPicker.Text = dr("Birthdate").ToString
        '    txtAge.Text = dr("Age").ToString
        '    txtPhone.Text = dr("Phone").ToString
        '    cbxHusband.Text = dr("HusName").ToString
        'End While
        'conn.Close()

        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "SELECT * FROM [Pat] " &
                                  "WHERE Name LIKE '%" & Me.cbxSearch.Text & "%' " &  '& Me.cbxSearch.Text & " " 
                                  "ORDER BY Patient_no ASC"
                'cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(Me.cbxSearch.Text))

                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        cbxPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        cbxHusband.Text = dt.Rows(0).Item("HusName").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub SearchID()
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con                               ''##

                cmd.CommandText = "SELECT * FROM Pat " &
                                  "WHERE Patient_no = @Patient " &  '& Me.cbxSearch.Text & " " 
                                  "ORDER BY Patient_no ASC"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(Me.cbxSearch.Text))

                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        cbxPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        cbxHusband.Text = dt.Rows(0).Item("HusName").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub SearchHusband()
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con                               ''##

                cmd.CommandText = "SELECT * FROM Pat " &
                                  "WHERE HusName LIKE '%" & Me.cbxSearch.Text & "%' " &  '& Me.cbxSearch.Text & " " 
                                  "ORDER BY Patient_no ASC"
                'cmd.Parameters.Add("@HusName", OleDbType.VarChar).Value = CStr(Me.cbxSearch.Text)

                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        'txtName.Text = dt.Rows(0).Item("Name")
                        cbxPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        cbxHusband.Text = dt.Rows(0).Item("HusName").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub SearchPhone()
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con                               ''##

                cmd.CommandText = "SELECT * FROM Pat " &
                                  "WHERE Phone LIKE '%" & Me.cbxSearch.Text & "%' " &  '& Me.cbxSearch.Text & " " 
                                  "ORDER BY Patient_no ASC"
                'cmd.Parameters.Add("@Phone", OleDbType.VarChar).Value = CStr(Me.cbxSearch.Text)

                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        cbxPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        cbxHusband.Text = dt.Rows(0).Item("HusName").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub GynPatSearch()

    End Sub

    Sub Gyn2PatSearch()

    End Sub

    Private Sub SearchPattxtNo()

    End Sub

    Sub GynPattxtNo()

    End Sub

    Sub Gyn2PattxtNo()

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        CheckNull("")

        SaveButton()
        SaveGyn()

    End Sub

    Sub SaveButton()
        Try
            If txtNo.Text = GetAutonumber("Pat", "Patient_no") Then
                'cmd = New OleDbCommand("INSERT INTO Pat(Patient_no, Name, Address, Birthdate, Age, Phone, HusName)" &
                '                       "VALUES(" & txtNo.Text & ", '" & cbxPatName.Text & "', '" & cbxAddress.Text & "', '" & DTPicker.Text & "', '" & txtAge.Text & "', '" & txtPhone.Text & "', '" & cbxHusband.Text & "')", conn)
                '#Please taking care of the single quote here specially with phone.text and DateTimePicker
                'RunCommand("INSERT INTO Pat(Patient_no, Name, Address, Birthdate, Age, Phone, HusName)" &
                '   "VALUES(" & txtNo.Text & ", '" & cbxPatName.Text & "', '" & cbxAddress.Text & "', '" & DTPicker.Text & "', '" & txtAge.Text & "', '" & txtPhone.Text & "', '" & cbxHusband.Text & "')")

                cmd = New OleDbCommand("INSERT INTO Pat(Patient_no, Name, Address, Birthdate, Age, Phone, HusName)" &
                   "VALUES(@Patient_no, @Name, @Address, @Birthdate, @Age, @Phone, @HusName)", conn)

                With cmd.Parameters
                    .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                    .Add("@Name", OleDbType.VarChar).Value = cbxPatName.Text
                    .Add("@Address", OleDbType.VarChar).Value = cbxAddress.Text
                    .Add("@Birthdate", OleDbType.DBDate).Value = CDate(DTPicker.Value)
                    .Add("@Age", OleDbType.VarChar).Value = txtAge.Text
                    .Add("@Phone", OleDbType.VarChar).Value = txtPhone.Text
                    .Add("@HusName", OleDbType.VarChar).Value = cbxHusband.Text

                End With

                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub UpdatePatient()
        Try
            If cbxPatName.Text <> "" And txtNo.Text <> GetAutonumber("Pat", "Patient_no") Then
                'RunCommand("UPDATE Pat SET Name='" & cbxPatName.Text & "', Address='" & cbxAddress.Text & "', Birthdate='" & DTPicker.Text & "', Age='" & txtAge.Text & "', Phone='" & txtPhone.Text & "', HusName='" & cbxHusband.Text & "' WHERE Patient_no=" & txtNo.Text)
                'RunCommand("UPDATE Visits SET Name='" & txtName.Text & "' WHERE Patient_no=" & frm2.txtVisNo.Text)
                cmd = New OleDbCommand("UPDATE Pat SET Name=@Name, Address=@Address," &
                                   " Birthdate=@Birthdate, Age=@Age, Phone=@Phone," &
                                   " HusName=@HusName WHERE Patient_no=@Patient_no", conn)

                With cmd.Parameters
                    .Add("@Name", OleDbType.VarChar).Value = cbxPatName.Text
                    .Add("@Address", OleDbType.VarChar).Value = cbxAddress.Text
                    .Add("@Birthdate", OleDbType.DBDate).Value = CDate(DTPicker.Value)
                    .Add("@Age", OleDbType.VarChar).Value = txtAge.Text
                    .Add("@Phone", OleDbType.VarChar).Value = txtPhone.Text
                    .Add("@HusName", OleDbType.VarChar).Value = cbxHusband.Text
                    .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                End With

                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
                conn.Open()
                cmd.ExecuteNonQuery()
                conn.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub UpdateVisName()
        cmd = New OleDbCommand("UPDATE Pat INNER JOIN Visits ON Pat.Patient_no = Visits.Patient_no SET Visits.Name = [Pat].[Name];", conn)

        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()

    End Sub

    Sub UpdateInvName()
        cmd = New OleDbCommand("UPDATE Pat INNER JOIN Inves ON Pat.Patient_no = Inves.Patient_no SET Inves.Name = [Pat].[Name];", conn)

        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()

    End Sub

    Private Sub SaveGyn()

        If txtVis1.Text = GetAutonumber("Gyn", "Vis_no") Then
            cmd = New OleDbCommand("INSERT INTO Gyn(Patient_no, Vis_no, G, P, A, NVD, CS, HPOC, LD, LC," &
                               "MnsDate, LMPDate, EDDDate, ElapW, GAW, MedH1, MedH2, MedH3," &
                               "SurH1, SurH2, SurH3, GynH1, GynH2, GynH3, DrugH1, DrugH2, DrugH3, Gyn)" &
                               "VALUES(@Patient_no, @Vis_no, @G, @P, @A, @NVD, @CS, @HPOC, @LD, @LC," &
                               "@MnsDate, @LMPDate, @EDDDate, @ElapW, @GAW, @MedH1, @MedH2, @MedH3," &
                               "@SurH1, @SurH2, @SurH3, @GynH1, @GynH2, @GynH3, @DrugH1, @DrugH2, @DrugH3, @Gyn)", conn)
            With cmd.Parameters
                .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                .Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVis1.Text))
                .Add("@G", OleDbType.VarChar).Value = txtG.Text
                .Add("@P", OleDbType.VarChar).Value = txtP.Text
                .Add("@A", OleDbType.VarChar).Value = txtA.Text
                .Add("@NVD", OleDbType.Boolean).Value = chbxNVD.Checked
                .Add("@CS", OleDbType.Boolean).Value = chbxCS.Checked
                .Add("@HPOC", OleDbType.VarChar).Value = cbxHPOC.Text
                .Add("@LD", OleDbType.VarChar).Value = cbxLD.Text
                .Add("@LC", OleDbType.VarChar).Value = cbxLC.Text
                .Add("@MnsDate", OleDbType.DBDate).Value = CDate(DTPickerMns.Value)
                .Add("@LMPDate", OleDbType.DBDate).Value = CDate(DTPickerLMP.Value)
                .Add("@EDDDate", OleDbType.DBDate).Value = CDate(DTPickerEDD.Value)
                .Add("@ElapW", OleDbType.VarChar).Value = txtElapsed.Text
                .Add("@GAW", OleDbType.VarChar).Value = txtGA.Text
                .Add("@MedH1", OleDbType.VarChar).Value = cbxMedH1.Text
                .Add("@MedH2", OleDbType.VarChar).Value = cbxMedH2.Text
                .Add("@MedH3", OleDbType.VarChar).Value = cbxMedH3.Text
                .Add("@SurH1", OleDbType.VarChar).Value = cbxSurH1.Text
                .Add("@SurH2", OleDbType.VarChar).Value = cbxSurH2.Text
                .Add("@SurH3", OleDbType.VarChar).Value = cbxSurH3.Text
                .Add("@GynH1", OleDbType.VarChar).Value = cbxGynH1.Text
                .Add("@GynH2", OleDbType.VarChar).Value = cbxGynH2.Text
                .Add("@GynH3", OleDbType.VarChar).Value = cbxGynH3.Text
                .Add("@DrugH1", OleDbType.VarChar).Value = cbxDrugH1.Text
                .Add("@DrugH2", OleDbType.VarChar).Value = cbxDrugH2.Text
                .Add("@DrugH3", OleDbType.VarChar).Value = cbxDrugH3.Text
                .Add("@Gyn", OleDbType.Boolean).Value = chbxGyn.Checked
            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If

    End Sub

    Private Sub UpdateGyn()
        If txtVis1.Text <> GetAutonumber("Gyn", "Vis_no") Then
            cmd = New OleDbCommand("UPDATE Gyn SET G=@G, P=@P, A=@A, NVD=@NVD, CS=@CS, HPOC=@HPOC, LD=@LD, LC=@LC," &
                                  "MnsDate=@MnsDate, LMPDate=@LMPDate, EDDDate=@EDDDate, ElapW=@ElapW," &
                                  "GAW=@GAW, MedH1=@MedH1, MedH2=@MedH2, MedH3=@MedH3," &
                                  "SurH1=@SurH1, SurH2=@SurH2, SurH3=@SurH3, GynH1=@GynH1," &
                                  "GynH2=@GynH2, GynH3=@GynH3, DrugH1=@DrugH1, DrugH2=@DrugH2, 
                                  DrugH3=@DrugH3, Gyn=@Gyn WHERE Vis_no=@Vis_no", conn)

            With cmd.Parameters
                .Add("@G", OleDbType.VarChar).Value = txtG.Text
                .Add("@P", OleDbType.VarChar).Value = txtP.Text
                .Add("@A", OleDbType.VarChar).Value = txtA.Text
                .Add("@NVD", OleDbType.Boolean).Value = chbxNVD.Checked
                .Add("@CS", OleDbType.Boolean).Value = chbxCS.Checked
                .Add("@HPOC", OleDbType.VarChar).Value = cbxHPOC.Text
                .Add("@LD", OleDbType.VarChar).Value = cbxLD.Text
                .Add("@LC", OleDbType.VarChar).Value = cbxLC.Text
                .Add("@MnsDate", OleDbType.DBDate).Value = CDate(DTPickerMns.Value)
                .Add("@LMPDate", OleDbType.DBDate).Value = CDate(DTPickerLMP.Value)
                .Add("@EDDDate", OleDbType.DBDate).Value = CDate(DTPickerEDD.Value)
                .Add("@ElapW", OleDbType.VarChar).Value = txtElapsed.Text
                .Add("@GAW", OleDbType.VarChar).Value = txtGA.Text
                .Add("@MedH1", OleDbType.VarChar).Value = cbxMedH1.Text
                .Add("@MedH2", OleDbType.VarChar).Value = cbxMedH2.Text
                .Add("@MedH3", OleDbType.VarChar).Value = cbxMedH3.Text
                .Add("@SurH1", OleDbType.VarChar).Value = cbxSurH1.Text
                .Add("@SurH2", OleDbType.VarChar).Value = cbxSurH2.Text
                .Add("@SurH3", OleDbType.VarChar).Value = cbxSurH3.Text
                .Add("@GynH1", OleDbType.VarChar).Value = cbxGynH1.Text
                .Add("@GynH2", OleDbType.VarChar).Value = cbxGynH2.Text
                .Add("@GynH3", OleDbType.VarChar).Value = cbxGynH3.Text
                .Add("@DrugH1", OleDbType.VarChar).Value = cbxDrugH1.Text
                .Add("@DrugH2", OleDbType.VarChar).Value = cbxDrugH2.Text
                .Add("@DrugH3", OleDbType.VarChar).Value = cbxDrugH3.Text
                .Add("@Gyn", OleDbType.Boolean).Value = chbxGyn.Checked
                .Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVis1.Text))
            End With
            'MsgBox("Update Done1")
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
            'MsgBox("Update Done2")
        End If

    End Sub

    Sub SaveGyn2()
        If txtVis.Text = GetAutonumber("Gyn2", "Vis_no") And cbxPatName.Text <> "" Then

            cmd = New OleDbCommand("INSERT INTO Gyn2(Vis_no, Patient_no, GL, Pls, BP," &
                               "Wt, BdBt, ChtH, HdNe, Ext, FunL, Scrs, Edm, US, Amount, AttDt)" &
                               "VALUES(@Vis_no, @Patient_no, @GL, @Pls, @BP, @Wt, @BdBt," &
                               "@ChtH, @HdNe, @Ext, @FunL, @Scrs, @Edm, @US, @Amount, @AttDt)", conn)

            With cmd.Parameters
                .Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVis.Text))
                .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                .Add("@GL", OleDbType.VarChar).Value = cbxGL.Text
                .Add("@Pls", OleDbType.VarChar).Value = cbxPuls.Text
                .Add("@BP", OleDbType.VarChar).Value = cbxBP.Text
                .Add("@Wt", OleDbType.VarChar).Value = cbxWeight.Text
                .Add("@BdBt", OleDbType.VarChar).Value = cbxBodyBuilt.Text
                .Add("@ChtH", OleDbType.VarChar).Value = cbxChtH.Text
                .Add("@HdNe", OleDbType.VarChar).Value = cbxHdNe.Text
                .Add("@Ext", OleDbType.VarChar).Value = cbxExt.Text
                .Add("@FunL", OleDbType.VarChar).Value = cbxFunL.Text
                .Add("@Scrs", OleDbType.VarChar).Value = cbxScars.Text
                .Add("@Edm", OleDbType.VarChar).Value = cbxEdema.Text
                .Add("@US", OleDbType.VarChar).Value = cbxUS.Text
                .Add("@Amount", OleDbType.VarChar).Value = txtAmount.Text
                .Add("@AttDt", OleDbType.DBDate).Value = CDate(DTPickerAtt.Value)
            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If

    End Sub

    Sub UpdateGyn2()
        If txtVis.Text <> GetAutonumber("Gyn2", "Vis_no") Then

            cmd = New OleDbCommand("UPDATE Gyn2 SET GL=@GL, Pls=@Pls, BP=@BP, Wt=@Wt, BdBt=@BdBt," &
                                   "ChtH=@ChtH, HdNe=@HdNe, Ext=@Ext, FunL=@FunL, Scrs=@Scrs, Edm=@Edm," &
                                   "US=@US, Amount=@Amount, AttDt=@AttDt WHERE Vis_no=@Vis_no", conn)

            With cmd.Parameters
                .Add("@GL", OleDbType.VarChar).Value = cbxGL.Text
                .Add("@Pls", OleDbType.VarChar).Value = cbxPuls.Text
                .Add("@BP", OleDbType.VarChar).Value = cbxBP.Text
                .Add("@Wt", OleDbType.VarChar).Value = cbxWeight.Text
                .Add("@BdBt", OleDbType.VarChar).Value = cbxBodyBuilt.Text
                .Add("@ChtH", OleDbType.VarChar).Value = cbxChtH.Text
                .Add("@HdNe", OleDbType.VarChar).Value = cbxHdNe.Text
                .Add("@Ext", OleDbType.VarChar).Value = cbxExt.Text
                .Add("@FunL", OleDbType.VarChar).Value = cbxFunL.Text
                .Add("@Scrs", OleDbType.VarChar).Value = cbxScars.Text
                .Add("@Edm", OleDbType.VarChar).Value = cbxEdema.Text
                .Add("@US", OleDbType.VarChar).Value = cbxUS.Text
                .Add("@Amount", OleDbType.VarChar).Value = txtAmount.Text
                .Add("@AttDt", OleDbType.DBDate).Value = CDate(DTPickerAtt.Value)
                .Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVis.Text))
            End With
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If

    End Sub
    '## Paul https://social.msdn.microsoft.com/Forums/vstudio/en-US/35b9de93-e5fd-4e3f-a8f6-97516184d4c7/what-is-the-best-solution-for-compact-and-repair-access-2013-database?forum=vbgeneral
    Sub CompactAccessDatabase()

        '##Path For Real Projects by "Application.StartupPath"
        'Dim DatabasePath As String = "D:KMAClinic\bin\Release\Dr_T.accdb"
        Dim DatabasePath As String = Path.Combine(Application.StartupPath, "TestDB.accdb")

        '##For Real Projects
        Dim DatabasePathCompacted As String = "D:\KMAClinic\ComDB\_" & Format(Now(), "yyyyMMdd_hhmmss") & ".accdb"

        Dim CompactDB As New Microsoft.Office.Interop.Access.Dao.DBEngine

        '##Here you can write your database password with this method (DatabasePath, DatabasePathCompacted, , , ";pwd=mero1981923")
        CompactDB.CompactDatabase(DatabasePath, DatabasePathCompacted, , , ";pwd=mero1981923")
        CompactDB = Nothing

    End Sub

    Sub GotoVisitPH()
        Trace.WriteLine("GotoVisitPH started @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        f2 = New Form2(cbxPatName.Text, txtNo.Text)
        f2.Show()
        Me.Hide()

        Trace.WriteLine("GotoVisitPH FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnVisits_Click(sender As Object, e As EventArgs) Handles btnVisits.Click
        Trace.WriteLine("btnVisits_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtNo.Text <> GetAutonumber("Pat", "Patient_no") Then
            GotoVisitPH()

        ElseIf cbxPatName.Text = "" Then

            f2 = New Form2("", "")
            f2.Show()
            Me.Hide()

        End If
        Trace.WriteLine("btnVisits_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        If MsgBox("You Will Exit The Clinic" + vbCrLf +
                  "Are you sure ?", MsgBoxStyle.YesNo,
                  "Confirm Message") = vbNo Then
            Exit Sub
        Else
            Close()
            Application.ExitThread()
        End If


    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        RDXmlDiaInter()
        RDXmlDrugs()
        RDXmlPlan()
        RDXmlInv()
        RDXmlInv2()
        RDXmlInvRes()
        RDXmlPatNames()
        RDXmlPatNames1()
        RDXmlPatNames2()

        loaddata()
    End Sub

    Private Sub btnBackup_Click(sender As Object, e As EventArgs) Handles btnBackup.Click
        If MsgBox("Are you sure that the Network PC is turned off," & vbCrLf &
                  "In order to perform this action?",
                  MsgBoxStyle.YesNo,
                  "Turn Off the network PC") = vbNo Then
            Exit Sub
        End If
        CompactAccessDatabase()
        Dim BackupFilePath As String = "D:\KMAClinic\ComDB"
        MsgBox("Backup Done @" & vbCrLf & BackupFilePath, MsgBoxStyle.Information, "Backup")
    End Sub

    Private Sub btnNewGyn_Click(sender As Object, e As EventArgs) Handles btnNewGyn.Click
        ClearGyn()
        If txtVis1.Text <> GetAutonumber("Gyn", "Vis_no") Then
            txtVis1.Text = GetAutonumber("Gyn", "Vis_no")
            txtVis1.Select()
        ElseIf txtVis1.Text = GetAutonumber("Gyn", "Vis_no") And cbxPatName.Text <> "" Then
            SaveGyn()
            txtVis1.Select()
        End If
        GynEnabled()
        btnL.Enabled = True
    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        ClearGyn2()
        If txtVis.Text <> GetAutonumber("Gyn2", "Vis_no") Then
            txtVis.Text = GetAutonumber("Gyn2", "Vis_no")
            txtVis.Select()
        ElseIf txtVis.Text = GetAutonumber("Gyn2", "Vis_no") And cbxPatName.Text <> "" Then
            SaveGyn2()
            txtVis.Select()
        End If
        Gyn2Enabled()
        btnF.Enabled = True
    End Sub

    Private Sub cbxPatName_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPatName.Validating
        Dim ds As DataSet = New DataSet
        Dim da As OleDbDataAdapter = New OleDbDataAdapter("SELECT Name FROM Pat WHERE Name='" & cbxPatName.Text & "'", conn)
        da.Fill(ds, "Pat")
        Dim dv As DataView = New DataView(ds.Tables("Pat"))
        Dim cur As CurrencyManager
        cur = CType(Me.BindingContext(dv), CurrencyManager)

        If cur.Count <> 0 And txtNo.Text = GetAutonumber("Pat", "Patient_no") Then
            MsgBox("تأكد من الاسم. هذا الاسم موجود من قبل", MsgBoxStyle.OkOnly, "يجب تغيير الاسم")
            cbxPatName.ResetText()
            Exit Sub
        End If
        If cbxPatName.Text <> "" Then
            SaveButton()
            'SaveGyn()
            UpdatePatient()
        End If
    End Sub

    Private Sub cbxPatName_Validated(sender As Object, e As EventArgs) Handles cbxPatName.Validated

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxAddress.Items.AddRange(cbElements)
            SaveInXmlPatNames()
        End If

    End Sub

    Private Sub cbxPatName_Click(sender As Object, e As EventArgs) Handles cbxPatName.Click
        If CheckBox1.Checked = False Then
            InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        ElseIf CheckBox1.Checked = True Then
            InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        End If
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxPatName.Items.AddRange(cbElements)
    End Sub

    Private Sub cbxAddress_Validating(sender As Object, e As EventArgs) Handles cbxAddress.Validating
        If cbxPatName.Text <> String.Empty Then

            UpdatePatient()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxHusband.Items.AddRange(cbElements)
            SaveInXmlPatNames2()
        End If
    End Sub

    Private Sub cbxAddress_Click(sender As Object, e As EventArgs) Handles cbxAddress.Click
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames2.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxAddress.Items.AddRange(cbElements)
    End Sub

    Private Sub DTPicker_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPicker.Validating
        If cbxPatName.Text <> String.Empty Then
            UpdatePatient()
        End If
    End Sub

    '## Kareninstructor
    Private operations As DateTimeCalculations = New DateTimeCalculations
    Private Sub DTPicker_ValueChanged(sender As Object, e As EventArgs) Handles DTPicker.ValueChanged
        ''##https://social.msdn.microsoft.com/Forums/vstudio/en-US/b2a15b26-6d51-49d5-81cf-20fef70e8316/when-datetimepicker-value-changed-this-error-occured?forum=vbgeneral
        operations.ToAgeString(DTPicker.Value)
        txtAge.Text = operations.Formatted
    End Sub

    Private Sub txtAge_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtAge.Validating
        If cbxPatName.Text <> String.Empty Then
            UpdatePatient()
        End If
    End Sub

    Private Sub txtPhone_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtPhone.Validating
        If cbxPatName.Text <> String.Empty Then
            UpdatePatient()
        End If
    End Sub

    Private Sub cbxHusband_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxHusband.Validating
        If cbxPatName.Text <> String.Empty Then
            UpdatePatient()
            SaveInXmlPatNames1()
        End If

    End Sub

    Private Sub cbxHusband_Click(sender As Object, e As EventArgs) Handles cbxHusband.Click
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames1.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxHusband.Items.AddRange(cbElements)
    End Sub

    Private Sub PictureBox1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDoubleClick
        Close()
        End
    End Sub

    Private Sub EnabledMenst()
        Label1.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        DTPickerEDD.Visible = True
        txtElapsed.Visible = True
        txtGA.Visible = True

    End Sub

    Private Sub DisabledMenst()
        Label1.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        DTPickerEDD.Visible = False
        txtElapsed.Visible = False
        txtGA.Visible = False
    End Sub

    Private Sub chbxGyn_CheckedChanged(sender As Object, e As EventArgs) Handles chbxGyn.CheckedChanged
        Dim sum1 As Integer
        Dim sum2 As Integer
        If chbxGyn.Checked = False Then
            sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
            txtG.Text = CType(sum1, String)
            EnabledMenst()
        ElseIf chbxGyn.Checked = True Then
            sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
            txtG.Text = CType(sum2, String)
            DisabledMenst()
        End If

    End Sub

    Private Sub chbxGyn_CheckStateChanged(sender As Object, e As EventArgs) Handles chbxGyn.CheckStateChanged
        'Dim sum1 As Integer
        'Dim sum2 As Integer
        'If chbxGyn.Checked = False Then
        '    sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
        '    txtG.Text = CType(sum1, String)
        'ElseIf chbxGyn.Checked = True Then
        '    sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
        '    txtG.Text = CType(sum2, String)
        'End If
    End Sub

    Private Sub chbxGyn_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles chbxGyn.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub txtG_TextChanged(sender As Object, e As EventArgs) Handles txtG.TextChanged
        If txtA.Text = "" And txtP.Text = "" Then
            txtG.Text = ""
        End If
    End Sub

    Private Sub txtG_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtG.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub txtG_MouseClick(sender As Object, e As MouseEventArgs) Handles txtG.MouseClick
        'Dim sum1, sum2 As Integer
        'sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
        'sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
        'If txtG.Text = CType(sum1, String) Then
        '    txtG.Text = CType(sum2, String)
        'End If
    End Sub

    Private Sub txtA_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtA.Validating
        Dim sum1 As Integer
        Dim sum2 As Integer
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        ElseIf chbxGyn.Checked = False Then
            sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
            txtG.Text = CType(sum1, String)
        ElseIf chbxGyn.Checked = True Then
            sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
            txtG.Text = CType(sum2, String)
        End If
    End Sub

    Private Sub txtP_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtP.Validating
        Dim sum1 As Integer
        Dim sum2 As Integer
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        ElseIf chbxGyn.Checked = False Then
            sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
            txtG.Text = CType(sum1, String)
        ElseIf chbxGyn.Checked = True Then
            sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
            txtG.Text = CType(sum2, String)
        End If
    End Sub

    Private Sub txtA_TextChanged(sender As Object, e As EventArgs) Handles txtA.TextChanged
        Dim sum1 As Integer
        Dim sum2 As Integer
        If chbxGyn.Checked = False Then
            sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
            txtG.Text = CType(sum1, String)
        Else
            sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
            txtG.Text = CType(sum2, String)
        End If

    End Sub

    Private Sub txtP_TextChanged(sender As Object, e As EventArgs) Handles txtP.TextChanged
        Dim sum1 As Integer
        Dim sum2 As Integer
        If chbxGyn.Checked = False Then
            sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
            txtG.Text = CType(sum1, String)
        Else
            sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
            txtG.Text = CType(sum2, String)
        End If
    End Sub

    Private Sub chbxCS_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles chbxCS.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub chbxNVD_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles chbxNVD.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub


    Private Sub cbxHPOC_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxHPOC.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxLD.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxHPOC_Click(sender As Object, e As EventArgs) Handles cbxHPOC.Click
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxHPOC.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxLD_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxLD.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxLC.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxLD_Click(sender As Object, e As EventArgs) Handles cbxLD.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxLD.Items.AddRange(cbElements)
        End If
    End Sub
    Private Sub cbxLC_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxLC.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH1.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If

    End Sub

    Private Sub cbxLC_Click(sender As Object, e As EventArgs) Handles cbxLC.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxLC.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub DateTimePicker1_VisibleChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.VisibleChanged
        Dim date1 As Date = DateTimePicker1.Value  ''##Equal Now
        Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        txtElapsed.Text = weeks & "  Weeks"

        ''##https://social.msdn.microsoft.com/Forums/vstudio/en-US/b2a15b26-6d51-49d5-81cf-20fef70e8316/when-datetimepicker-value-changed-this-error-occured?forum=vbgeneral
        operations.ToAgeString(DTPicker.Value)
        txtAge.Text = operations.Formatted
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim date1 As Date = DateTimePicker1.Value  ''##Equal Now
        Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        txtElapsed.Text = weeks & "  Weeks"

        ''##https://social.msdn.microsoft.com/Forums/vstudio/en-US/b2a15b26-6d51-49d5-81cf-20fef70e8316/when-datetimepicker-value-changed-this-error-occured?forum=vbgeneral
        operations.ToAgeString(DTPicker.Value)
        txtAge.Text = operations.Formatted
    End Sub

    Private Sub DTPickerMns_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerMns.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub DTPickerMns_ValueChanged(sender As Object, e As EventArgs) Handles DTPickerMns.ValueChanged
        Dim date1 As Date = DTPickerMns.Value  ''##Equal Now
        Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        txtElapsed.Text = weeks & "  Weeks"

    End Sub

    Private Sub DTPickerEDD_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerEDD.Validating
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub DTPickerLMP_ValueChanged(sender As Object, e As EventArgs) Handles DTPickerLMP.ValueChanged
        Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        Dim date2 As Date = DTPickerLMP.Value
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        txtElapsed.Text = CStr(weeks) '& "  Weeks"

        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(7)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddMonths(9)
        '#For increasing 40 weeks = 280 days 
        DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)

        'Dim sum1, sum2 As Integer
        'sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
        'sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
        'If txtG.Text = CType(sum1, String) Then
        '    txtElapsed.Text = weeks & "  Weeks"
        '    DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
        'ElseIf txtg.Text = CType(sum2, String) Or txtG.Text = "" Then
        '    DTPickerEDD.Enabled = False
        '    txtGA.Text = ""
        '    txtElapsed.Text = ""
        'End If
    End Sub

    Private Sub DTPickerLMP_Validating(sender As Object, e As EventArgs) Handles DTPickerLMP.Validating
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(7)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddMonths(9)

        DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)

        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub txtGA_TextChanged(sender As Object, e As EventArgs) Handles txtGA.TextChanged

    End Sub

    Private Sub txtGA_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtGA.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub txtElapsed_TextChanged(sender As Object, e As EventArgs) Handles txtElapsed.TextChanged
        Dim date1 As Date = DateTimePicker1.Value  'Now 'DTPickerMns.Value
        Dim date2 As Date = DTPickerLMP.Value

        Dim weeks As Integer = CInt((date2 - date1).TotalDays / 7)
        txtGA.Text = CStr(weeks + 40) '& "  Weeks"

    End Sub

    Private Sub txtElapsed_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtElapsed.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub cbxMedH1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxMedH1.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH2.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxMedH1_Click(sender As Object, e As EventArgs) Handles cbxMedH1.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH1.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxMedH2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxMedH2.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH3.Items.AddRange(cbElements)
            SaveInXmlDiaInter()

        End If
    End Sub

    Private Sub cbxMedH2_Click(sender As Object, e As EventArgs) Handles cbxMedH2.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH2.Items.AddRange(cbElements)

        End If
    End Sub

    Private Sub cbxMedH3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxMedH3.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()

            '' Now fill the ComboBox's 
            cbxSurH1.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxMedH3_Click(sender As Object, e As EventArgs) Handles cbxMedH3.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxSurH1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxSurH1.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSurH2.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxSurH1_Click(sender As Object, e As EventArgs) Handles cbxSurH1.Click

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSurH1.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxSurH2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxSurH2.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSurH3.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxSurH2_Click(sender As Object, e As EventArgs) Handles cbxSurH2.Click

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSurH2.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxSurH3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxSurH3.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH1.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxSurH3_Click(sender As Object, e As EventArgs) Handles cbxSurH3.Click

        If cbxPatName.Text <> "" Then

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSurH3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxGynH1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxGynH1.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH2.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxGynH1_Click(sender As Object, e As EventArgs) Handles cbxGynH1.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH1.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxGynH2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxGynH2.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH3.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxGynH2_Click(sender As Object, e As EventArgs) Handles cbxGynH2.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH2.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxGynH3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxGynH3.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH1.Items.AddRange(cbElements)
            SaveInXmlDiaInter()
        End If
    End Sub

    Private Sub cbxGynH3_Click(sender As Object, e As EventArgs) Handles cbxGynH3.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxDrugH1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrugH1.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH2.Items.AddRange(cbElements)
            SaveInXmlDrugs()
        End If
    End Sub

    Private Sub cbxDrugH1_Click(sender As Object, e As EventArgs) Handles cbxDrugH1.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH1.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxDrugH2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrugH2.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH3.Items.AddRange(cbElements)
            SaveInXmlDrugs()
        End If
    End Sub

    Private Sub cbxDrugH2_Click(sender As Object, e As EventArgs) Handles cbxDrugH2.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH2.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxDrugH3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrugH3.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn()

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGL.Items.AddRange(cbElements)
            SaveInXmlDrugs()
        End If
    End Sub

    Private Sub cbxDrugH3_Click(sender As Object, e As EventArgs) Handles cbxDrugH3.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub txtVis_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtVis.Validating
        If txtVis.Text = GetAutonumber("Gyn2", "Vis_no") And cbxPatName.Text <> "" Then
            SaveGyn2()
        End If
    End Sub

    Private Sub txtVis1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtVis1.Validating
        If txtVis1.Text = GetAutonumber("Gyn", "Vis_no") And cbxPatName.Text <> "" Then
            SaveGyn()
        End If
    End Sub

    Private Sub cbxGL_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxGL.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPuls.Items.AddRange(cbElements)

        End If
    End Sub

    Private Sub cbxGL_Click(sender As Object, e As EventArgs) Handles cbxGL.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGL.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPuls_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPuls.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxBP.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlus_Click(sender As Object, e As EventArgs) Handles cbxPuls.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPuls.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxBP_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxBP.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxWeight.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxBP_Click(sender As Object, e As EventArgs) Handles cbxBP.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxBP.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxWeight_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxWeight.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxBodyBuilt.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxWeight_Click(sender As Object, e As EventArgs) Handles cbxWeight.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxWeight.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxBodyBuilt_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxBodyBuilt.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxChtH.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxBodyBuilt_Click(sender As Object, e As EventArgs) Handles cbxBodyBuilt.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxBodyBuilt.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxChtH_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxChtH.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxHdNe.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxChtH_Click(sender As Object, e As EventArgs) Handles cbxChtH.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxChtH.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxHdNe_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxHdNe.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxExt.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxHdNe_Click(sender As Object, e As EventArgs) Handles cbxHdNe.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxHdNe.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxExt_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxExt.Validating

        If cbxPatName.Text <> "" Then
            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxFunL.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxExt_Click(sender As Object, e As EventArgs) Handles cbxExt.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxExt.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxFunL_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxFunL.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxScars.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxFunL_Click(sender As Object, e As EventArgs) Handles cbxFunL.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxFunL.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxScars_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxScars.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxEdema.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxScars_Click(sender As Object, e As EventArgs) Handles cbxScars.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxScars.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxEdema_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxEdema.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxUS.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxEdema_Click(sender As Object, e As EventArgs) Handles cbxEdema.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxEdema.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxUS_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxUS.Validating

        If cbxPatName.Text <> "" Then

            UpdateGyn2()
            SaveInXmlInv2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxHPOC.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxUS_Click(sender As Object, e As EventArgs) Handles cbxUS.Click

        If cbxPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxUS.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub txtAmount_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtAmount.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn2()
        End If
    End Sub

    Private Sub DTPickerAtt_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerAtt.Validating
        If cbxPatName.Text <> "" Then
            UpdateGyn2()
        End If
    End Sub

    '##Save Patient Names
    Sub SaveInXmlPatNames()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\PatNames.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Names")
            .WriteElementString("Name", cbxPatName.Text)
            .WriteEndElement()
            .Close()
        End With
        xmldoc.Save(Directory.GetCurrentDirectory & "\PatNames.xml")

    End Sub
    '##Save Husband Names
    Sub SaveInXmlPatNames1()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\PatNames1.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Names")
            .WriteElementString("Name", cbxHusband.Text)
            .WriteEndElement()
            .Close()
        End With
        xmldoc.Save(Directory.GetCurrentDirectory & "\PatNames1.xml")

    End Sub
    '##Save Addresses
    Sub SaveInXmlPatNames2()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\PatNames2.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Names")
            .WriteElementString("Name", cbxAddress.Text)
            .WriteEndElement()
            .Close()
        End With

        xmldoc.Save(Directory.GetCurrentDirectory & "\PatNames2.xml")
    End Sub

    Sub SaveInXmlDiaInter()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\DiaInter.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxHPOC.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxLD.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxLC.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxMedH1.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxMedH2.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxMedH3.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxSurH1.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxSurH2.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxSurH3.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxGynH1.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxGynH2.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxGynH3.Text)
            .WriteEndElement()
            .Close()
        End With

        xmldoc.Save(Directory.GetCurrentDirectory & "\DiaInter.xml")
    End Sub

    Sub SaveInXmlDrugs()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\Drugs1.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrugH1.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrugH2.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrugH3.Text)
            .WriteEndElement()
            .Close()
        End With

        xmldoc.Save(Directory.GetCurrentDirectory & "\Drugs1.xml")
    End Sub

    Sub SaveInXmlInv2()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\Investigations2.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxGL.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxPuls.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxBP.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxWeight.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxBodyBuilt.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxChtH.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxHdNe.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxEdema.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxFunL.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxScars.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxExt.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxUS.Text)
            .WriteEndElement()
            .Close()
        End With

        xmldoc.Save(Directory.GetCurrentDirectory & "\Investigations2.xml")
    End Sub

    Sub RDXmlPlan()
        Trace.WriteLine("RDXmlPlan STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim fileName1 As String = "Plans.xml"
        Dim xdoc1 As XDocument = XDocument.Load(fileName1)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc1.Descendants("Plans")
                       Group n By Item = n.Element("Plan").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system
        xdoc1.Save(fileName1)
        Trace.WriteLine("RDXmlPlan FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub RDXmlInv()
        Trace.WriteLine("RDXmlInv STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim fileName2 As String = "Investigations.xml"
        Dim xdoc2 As XDocument = XDocument.Load(fileName2)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc2.Descendants("Invest")
                       Group n By Item = n.Element("Name").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system
        xdoc2.Save(fileName2)
        Trace.WriteLine("RDXmlInv FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub RDXmlInvRes()
        Trace.WriteLine("RDXmlInvRes STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim fileName As String = "InvRes.xml"
        Dim xdoc As XDocument = XDocument.Load(fileName)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc.Descendants("Invest")
                       Group n By Item = n.Element("Name").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system                                                
        xdoc.Save(fileName)
        Trace.WriteLine("RDXmlInvRes FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
    Sub RDXmlInv2()
        Dim fileName1 As String = "Investigations2.xml"
        Dim xdoc1 As XDocument = XDocument.Load(fileName1)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc1.Descendants("Invest")
                       Group n By Item = n.Element("Name").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system
        xdoc1.Save(fileName1)
    End Sub

    Sub RDXmlDiaInter()
        Dim fileName1 As String = "DiaInter.xml"
        Dim xdoc As XDocument = XDocument.Load(fileName1)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc.Descendants("Invest")
                       Group n By Item = n.Element("Name").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system
        xdoc.Save(fileName1)
    End Sub

    Sub RDXmlDrugs()
        Dim fileName1 As String = "Drugs1.xml"
        Dim xdoc As XDocument = XDocument.Load(fileName1)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc.Descendants("Drugs")
                       Group n By Item = n.Element("Drug").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system
        xdoc.Save(fileName1)
    End Sub
    ''##Removing Dplicates from Names Xml file 
    Sub RDXmlPatNames()
        Dim fileName1 As String = "PatNames.xml"
        Dim xdoc1 As XDocument = XDocument.Load(fileName1)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc1.Descendants("Names")
                       Group n By Item = n.Element("Name").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system
        xdoc1.Save(fileName1)
    End Sub

    ''##Removing Dplicates from Husband Xml file 
    Sub RDXmlPatNames1()
        Dim fileName1 As String = "PatNames1.xml"
        Dim xdoc1 As XDocument = XDocument.Load(fileName1)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc1.Descendants("Names")
                       Group n By Item = n.Element("Name").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system
        xdoc1.Save(fileName1)
    End Sub

    ''##Removing Dplicates from Addresses Xml file 
    Sub RDXmlPatNames2()
        Dim fileName1 As String = "PatNames2.xml"
        Dim xdoc1 As XDocument = XDocument.Load(fileName1)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc1.Descendants("Names")
                       Group n By Item = n.Element("Name").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system
        xdoc1.Save(fileName1)
    End Sub

    Private Sub btnF_Click(sender As Object, e As EventArgs) Handles btnF.Click

        ListBox2.Items.Clear()
        If cbxPatName.Text <> "" Then
            'btnF.Enabled = False
            Dim connection As OleDbConnection = New OleDbConnection()
            connection.ConnectionString = "provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923"
            Dim command As OleDbCommand = New OleDbCommand()
            command.Connection = connection
            command.CommandText = "SELECT Vis_no FROM Gyn2 WHERE Patient_no =" & txtNo.Text & " " &
            "ORDER BY Vis_no"
            command.CommandType = CommandType.Text
            connection.Open()

            Dim reader As OleDbDataReader = command.ExecuteReader()
            While reader.Read()
                'Dim PatientID As String = CStr(reader("Patient_no"))
                Dim VisitNO As String = CStr(reader("Vis_no"))
                Dim item As String = String.Format("{0}", VisitNO)
                Me.ListBox2.Items.Add(item).ToString()
            End While

            reader.Close()
            If connection.State = ConnectionState.Open Then connection.Close()


            TabControl1.SelectedTab = Me.TabPage2
            Label49.Text = "U/S Visits"
            Dim con As New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT * FROM Gyn2 WHERE Patient_no = @Patient_no
                               ORDER BY Vis_no DESC", con)
            cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
            'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
            Dim da As New OleDbDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds, "Gyn2")
            DataGridView2.DataSource = ds.Tables("Gyn2").DefaultView

            con.Close()
            Label50.Text = ((DataGridView1.Rows.Count) - 1).ToString()


        End If


    End Sub

    Private Sub btnL_Click(sender As Object, e As EventArgs) Handles btnL.Click
        Try

            ClearGyn()
            ListBox1.Items.Clear()
            'ListBox2.Items.Clear()
            If cbxPatName.Text <> "" Then
                'btnL.Enabled = False
                Dim connection As OleDbConnection = New OleDbConnection()
                connection.ConnectionString = "provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923"
                Dim command As OleDbCommand = New OleDbCommand()
                command.Connection = connection
                command.CommandText = "SELECT Vis_no FROM Gyn WHERE Patient_no =" & txtNo.Text & " " &
                                      "ORDER BY Vis_no"
                command.CommandType = CommandType.Text

                connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()
                While reader.Read()
                    'Dim PatientID As String = CStr(reader("Patient_no"))
                    Dim VisitNO As String = CStr(reader("Vis_no"))
                    Dim item As String = String.Format("{0}", VisitNO)
                    Me.ListBox1.Items.Add(item).ToString()

                End While

                reader.Close()
                If connection.State = ConnectionState.Open Then connection.Close()

                TabControl1.SelectedTab = Me.TabPage2
                Label48.Text = "Previous Visits"
                Dim con As New OleDbConnection(cs)
                con.Open()
                cmd = New OleDbCommand("SELECT * FROM Gyn WHERE Patient_no = @Patient_no
                               ORDER BY Vis_no DESC", con)
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
                Dim da As New OleDbDataAdapter(cmd)
                Dim ds As New DataSet
                da.Fill(ds, "Gyn")
                DataGridView1.DataSource = ds.Tables("Gyn").DefaultView

                con.Close()
                Label47.Text = ((DataGridView1.Rows.Count) - 1).ToString()

            End If

            GynDisabled()
            'Gyn2Disabled()
        Catch ex As Exception
            MsgBox(ErrorToString)
        End Try
    End Sub

    Private Sub RemoveTopItems()
        ' Determine if the currently selected item in the ListBox 
        ' is the item displayed at the top in the ListBox.
        If ListBox1.TopIndex <> ListBox1.SelectedIndex Then
            ' Make the currently selected item the top item in the ListBox.
            ListBox1.TopIndex = ListBox1.SelectedIndex
        End If
        ' Remove all items before the top item in the ListBox.
        Dim x As Integer
        For x = ListBox1.SelectedIndex - 1 To 0 Step -1
            ListBox1.Items.RemoveAt(x)
        Next x

        ' Clear all selections in the ListBox.
        ListBox1.ClearSelected()
    End Sub 'RemoveTopItems

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        TextBox1.Text = ""
        If ListBox1.SelectedIndex > -1 Then
            TextBox1.Text = CType(ListBox1.SelectedItem, String)
        End If
        Dim date1 As Date = DateTimePicker1.Value  ''##Equal Now
        Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        txtElapsed.Text = CStr(weeks) '& "  Weeks"
        'Dim sum1, sum2 As Integer
        'sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
        'sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
        If DTPickerLMP.Value = DTPickerMns.Value Then
            DTPickerEDD.Value = DTPickerMns.Value
            txtElapsed.Text = "0" '& "  Weeks"
            txtGA.Text = "0" '& "  Weeks"

            'ElseIf txtG.Text = CType(sum1, String) Then
            '    txtElapsed.Text = weeks & "  Weeks"
            '    DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
            'ElseIf txtG.Text = CType(sum2, String) Or txtG.Text = "" Then
            '    DTPickerEDD.Enabled = False
            '    txtGA.Text = ""
            '    txtElapsed.Text = ""
            'UpdateGyn()
        Else
            'Dim date1 As Date = DateTimePicker1.Value  ''##Equal Now
            'Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
            'Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
            txtElapsed.Text = CStr(weeks) '& "  Weeks"
            txtGA.Text = CStr(40 - weeks) '& "  Weeks"
        End If
        GynEnabled()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        conn.Open()

        txtNo.ResetText()
        txtVis1.ResetText()
        txtG.ResetText()
        txtP.ResetText()
        txtA.ResetText()
        chbxNVD.Checked = False
        chbxCS.Checked = False
        cbxHPOC.ResetText()
        cbxLD.ResetText()
        cbxLC.ResetText()
        DTPickerMns.ResetText()
        DTPickerLMP.ResetText()
        DTPickerEDD.ResetText()
        txtElapsed.ResetText()
        txtGA.ResetText()
        cbxMedH1.ResetText()
        cbxMedH2.ResetText()
        cbxMedH3.ResetText()
        cbxSurH1.ResetText()
        cbxSurH2.ResetText()
        cbxSurH3.ResetText()
        cbxGynH1.ResetText()
        cbxGynH2.ResetText()
        cbxGynH3.ResetText()
        cbxDrugH1.ResetText()
        cbxDrugH2.ResetText()
        cbxDrugH3.ResetText()
        chbxGyn.ResetText()

        Dim str As String = "SELECT * FROM Gyn WHERE Vis_no = @Vis_no" '& TextBox1.Text & " " '& 
        Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(TextBox1.Text))
        dr = cmd.ExecuteReader
        While dr.Read
            txtNo.Text = dr("Patient_no").ToString
            txtVis1.Text = dr("Vis_no").ToString
            txtG.Text = dr("G").ToString
            txtP.Text = dr("P").ToString
            txtA.Text = dr("A").ToString
            chbxNVD.Checked = CBool(dr("NVD").ToString)
            chbxCS.Checked = CBool(dr("CS").ToString)
            cbxHPOC.Text = dr("HPOC").ToString
            cbxLD.Text = dr("LD").ToString
            cbxLC.Text = dr("LC").ToString
            DTPickerMns.Text = dr("MNSDate").ToString
            DTPickerLMP.Text = dr("LMPDate").ToString
            DTPickerEDD.Text = dr("EDDDate").ToString
            txtElapsed.Text = dr("ElapW").ToString
            txtGA.Text = dr("GAW").ToString
            cbxMedH1.Text = dr("MedH1").ToString
            cbxMedH2.Text = dr("MedH2").ToString
            cbxMedH3.Text = dr("MedH3").ToString
            cbxSurH1.Text = dr("SurH1").ToString
            cbxSurH2.Text = dr("SurH2").ToString
            cbxSurH3.Text = dr("SurH3").ToString
            cbxGynH1.Text = dr("GynH1").ToString
            cbxGynH2.Text = dr("GynH2").ToString
            cbxGynH3.Text = dr("GynH3").ToString
            cbxDrugH1.Text = dr("DrugH1").ToString
            cbxDrugH2.Text = dr("DrugH2").ToString
            cbxDrugH3.Text = dr("DrugH3").ToString
            chbxGyn.Checked = CBool(dr("Gyn").ToString)
        End While
        conn.Close()

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        TextBox2.Text = ""
        If ListBox2.SelectedIndex > -1 Then
            TextBox2.Text = CType(ListBox2.SelectedItem, String)
        End If
        Gyn2Enabled()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        conn.Open()

        txtVis.ResetText()
        txtNo.ResetText()
        cbxGL.ResetText()
        cbxPuls.ResetText()
        cbxBP.ResetText()
        cbxWeight.ResetText()
        cbxBodyBuilt.ResetText()
        cbxChtH.ResetText()
        cbxHdNe.ResetText()
        cbxExt.ResetText()
        cbxFunL.ResetText()
        cbxScars.ResetText()
        cbxEdema.ResetText()
        cbxUS.ResetText()
        txtAmount.ResetText()
        DTPickerAtt.ResetText()

        Dim str As String = "SELECT * FROM Gyn2 WHERE Vis_no = @Vis_no" '& TextBox1.Text & " " '& 
        Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(TextBox2.Text))
        dr = cmd.ExecuteReader
        While dr.Read
            txtVis.Text = dr("Vis_no").ToString
            txtNo.Text = dr("Patient_no").ToString
            cbxGL.Text = dr("GL").ToString
            cbxPuls.Text = dr("Pls").ToString
            cbxBP.Text = dr("BP").ToString
            cbxWeight.Text = dr("Wt").ToString
            cbxBodyBuilt.Text = dr("BdBt").ToString
            cbxChtH.Text = dr("ChTH").ToString
            cbxHdNe.Text = dr("HdNe").ToString
            cbxExt.Text = dr("Ext").ToString
            cbxFunL.Text = dr("FunL").ToString
            cbxScars.Text = dr("Scrs").ToString
            cbxEdema.Text = dr("Edm").ToString
            cbxUS.Text = dr("US").ToString
            txtAmount.Text = dr("Amount").ToString
            DTPickerAtt.Text = dr("AttDt").ToString
        End While
        conn.Close()

    End Sub

    Private Sub GynEnabled()
        DTPickerMns.Enabled = True
        txtA.Enabled = True
        txtP.Enabled = True
        cbxLD.Enabled = True
        cbxLC.Enabled = True
        cbxHPOC.Enabled = True
        DTPickerLMP.Enabled = True
        DTPickerEDD.Enabled = True
        cbxMedH1.Enabled = True
        cbxMedH2.Enabled = True
        cbxMedH3.Enabled = True
        cbxGynH1.Enabled = True
        cbxGynH2.Enabled = True
        cbxGynH3.Enabled = True
        cbxSurH1.Enabled = True
        cbxSurH2.Enabled = True
        cbxSurH3.Enabled = True
        cbxDrugH1.Enabled = True
        cbxDrugH2.Enabled = True
        cbxDrugH3.Enabled = True
        chbxGyn.Enabled = True
        chbxNVD.Enabled = True
        chbxCS.Enabled = True
    End Sub

    Private Sub GynDisabled()
        DTPickerMns.Enabled = False
        txtA.Enabled = False
        txtP.Enabled = False
        chbxCS.Enabled = False
        chbxNVD.Enabled = False
        cbxLD.Enabled = False
        cbxLC.Enabled = False
        cbxHPOC.Enabled = False
        DTPickerLMP.Enabled = False
        DTPickerEDD.Enabled = False
        cbxMedH1.Enabled = False
        cbxMedH2.Enabled = False
        cbxMedH3.Enabled = False
        cbxGynH1.Enabled = False
        cbxGynH2.Enabled = False
        cbxGynH3.Enabled = False
        cbxSurH1.Enabled = False
        cbxSurH2.Enabled = False
        cbxSurH3.Enabled = False
        cbxDrugH1.Enabled = False
        cbxDrugH2.Enabled = False
        cbxDrugH3.Enabled = False
        chbxGyn.Enabled = False
    End Sub

    Private Sub Gyn2Enabled()
        DTPickerAtt.Enabled = True
        cbxGL.Enabled = True
        cbxPuls.Enabled = True
        cbxBP.Enabled = True
        cbxWeight.Enabled = True
        cbxBodyBuilt.Enabled = True
        cbxChtH.Enabled = True
        cbxHdNe.Enabled = True
        cbxExt.Enabled = True
        cbxFunL.Enabled = True
        cbxScars.Enabled = True
        cbxEdema.Enabled = True
        cbxUS.Enabled = True
        txtAmount.Enabled = True

    End Sub

    Private Sub Gyn2Disabled()
        DTPickerAtt.Enabled = False
        cbxGL.Enabled = False
        cbxPuls.Enabled = False
        cbxBP.Enabled = False
        cbxWeight.Enabled = False
        cbxBodyBuilt.Enabled = False
        cbxChtH.Enabled = False
        cbxHdNe.Enabled = False
        cbxExt.Enabled = False
        cbxFunL.Enabled = False
        cbxScars.Enabled = False
        cbxEdema.Enabled = False
        cbxUS.Enabled = False
        txtAmount.Enabled = False

    End Sub

    Private Sub lblAtt11_MouseHover(sender As Object, e As EventArgs) Handles lblAtt11.MouseHover
        'TextBox3.Visible = True
        If TextBox3.Text = "FSL_hggi1981923" Then
            Dim f4 As New Form4
            Me.Hide()
            f4.ShowDialog()
        End If

    End Sub

    Private Sub TextBox3_Validated(sender As Object, e As EventArgs) Handles TextBox3.Validated
        TextBox3.Visible = True
        If TextBox3.Text = "FSL_hggi1981923" Then
            Dim f4 As New Form4
            Me.Hide()
            f4.ShowDialog()
        End If
    End Sub

    Private Sub TextBox3_MouseHover(sender As Object, e As EventArgs) Handles TextBox3.MouseHover

        'TextBox3.Visible = True
        If TextBox3.Text = "FSL_hggi1981923" Then
            Label46.Visible = True
        End If
    End Sub

    Private Sub btnN_Click(sender As Object, e As EventArgs) Handles btnN.Click
        If TextBox3.Visible = False Then
            TextBox3.Visible = True
        ElseIf TextBox3.Visible = True Then
            TextBox3.Visible = False
        ElseIf TextBox3.Text = "FSL_hggi1981923" Then
            Dim f4 As New Form4
            Me.Hide()
            f4.ShowDialog()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.ListBox3.Items.Clear()
        'btnL.Enabled = True
        'btnF.Enabled = True

        Dim connection As OleDbConnection = New OleDbConnection()
        connection.ConnectionString = "provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923"
        Dim command As OleDbCommand = New OleDbCommand()
        command.Connection = connection

        command.CommandText = "SELECT Vis_no, EDDDate FROM Gyn WHERE (EDDDate >= ?) AND (GAW <= 4) AND (Gyn = 0) " &  '& txtNo.Text & " " &
                                  "ORDER BY EDDDate"
        command.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker1.Value
        'command.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker22.Value

        command.CommandType = CommandType.Text
        connection.Open()

        Dim reader As OleDbDataReader = command.ExecuteReader()
        While reader.Read()
            Dim Vis As String = CStr(reader("Vis_no"))
            Dim EDD As String = CStr(reader("EDDDate"))
            Dim item As String = String.Format("{0} : {1}", Vis, EDD & vbCrLf)
            Me.ListBox3.Items.Add(item).ToString()
        End While

        reader.Close()
        If connection.State = ConnectionState.Open Then connection.Close()

        TabControl1.SelectedTab = Me.TabPage2
        Label48.Text = "Expected Date Of Delivery"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT * FROM Gyn WHERE (EDDDate >= ?) AND (GAW <= 4)
                               ORDER BY EDDDate DESC", con)
        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker1.Value
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn")
        DataGridView1.DataSource = ds.Tables("Gyn").DefaultView

        con.Close()
        Label47.Text = ((DataGridView1.Rows.Count) - 1).ToString()
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        TextBox4.Text = ""
        If ListBox3.SelectedIndex > -1 Then
            TextBox4.Text = CStr(ListBox3.SelectedItem)
            'txtName.Text = CStr(ListBox1.SelectedItem)
        End If
        Dim date1 As Date = DateTimePicker1.Value   'Now
        Dim date2 As Date = DTPickerLMP.Value
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        If DTPickerEDD.Value <> DTPickerLMP.Value Then
            txtElapsed.Text = CStr(weeks)
            txtGA.Text = CStr(40 - weeks)
        ElseIf DTPickerEDD.Value = DTPickerLMP.Value Then
            txtElapsed.Text = "0"
            txtGA.Text = "0"
        End If
        ShowPatTable()
        GynEnabled()
        TabControl1.SelectedTab = Me.TabPage1
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        conn.Open()

        txtNo.ResetText()
        txtVis1.ResetText()
        txtG.ResetText()
        txtP.ResetText()
        txtA.ResetText()
        chbxNVD.Checked = False
        chbxCS.Checked = False
        cbxHPOC.ResetText()
        cbxLD.ResetText()
        cbxLC.ResetText()
        DTPickerMns.ResetText()
        DTPickerLMP.ResetText()
        DTPickerEDD.ResetText()
        txtElapsed.ResetText()
        txtGA.ResetText()
        cbxMedH1.ResetText()
        cbxMedH2.ResetText()
        cbxMedH3.ResetText()
        cbxSurH1.ResetText()
        cbxSurH2.ResetText()
        cbxSurH3.ResetText()
        cbxGynH1.ResetText()
        cbxGynH2.ResetText()
        cbxGynH3.ResetText()
        cbxDrugH1.ResetText()
        cbxDrugH2.ResetText()
        cbxDrugH3.ResetText()
        chbxGyn.ResetText()

        Dim str As String = "SELECT * FROM Gyn WHERE Vis_no=@Vis_no " &  '& txtNo.Text & " " &
                             "ORDER BY EDDDate"

        Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        'cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker1.Value
        cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(TextBox4.Text))
        dr = cmd.ExecuteReader
        While dr.Read
            txtNo.Text = dr("Patient_no").ToString
            txtVis1.Text = dr("Vis_no").ToString
            txtG.Text = dr("G").ToString
            txtP.Text = dr("P").ToString
            txtA.Text = dr("A").ToString
            chbxNVD.Checked = CBool(dr("NVD").ToString)
            chbxCS.Checked = CBool(dr("CS").ToString)
            cbxHPOC.Text = dr("HPOC").ToString
            cbxLD.Text = dr("LD").ToString
            cbxLC.Text = dr("LC").ToString
            DTPickerMns.Text = dr("MnsDate").ToString
            DTPickerLMP.Text = dr("LMPDate").ToString
            DTPickerEDD.Text = dr("EDDDate").ToString
            txtElapsed.Text = dr("ElapW").ToString
            txtGA.Text = dr("GAW").ToString
            cbxMedH1.Text = dr("MedH1").ToString
            cbxMedH2.Text = dr("MedH2").ToString
            cbxMedH3.Text = dr("MedH3").ToString
            cbxSurH1.Text = dr("SurH1").ToString
            cbxSurH2.Text = dr("SurH2").ToString
            cbxSurH3.Text = dr("SurH3").ToString
            cbxGynH1.Text = dr("GynH1").ToString
            cbxGynH2.Text = dr("GynH2").ToString
            cbxGynH3.Text = dr("GynH3").ToString
            cbxDrugH1.Text = dr("DrugH1").ToString
            cbxDrugH2.Text = dr("DrugH2").ToString
            cbxDrugH3.Text = dr("DrugH3").ToString
            chbxGyn.Checked = CBool(dr("Gyn").ToString)
        End While
        conn.Close()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.ListBox3.Items.Clear()
        'btnL.Enabled = True
        'btnF.Enabled = True

        Dim connection As OleDbConnection = New OleDbConnection()
        connection.ConnectionString = "provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923"
        Dim command As OleDbCommand = New OleDbCommand()
        command.Connection = connection

        command.CommandText = "SELECT Vis_no, EDDDate FROM Gyn WHERE (EDDDate >= ? AND ? >= EDDDate) AND (Gyn = 0) " &  '& txtNo.Text & " " &
                                  "ORDER BY EDDDate"
        command.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
        command.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value

        command.CommandType = CommandType.Text
        connection.Open()

        Dim reader As OleDbDataReader = command.ExecuteReader()
        While reader.Read()
            Dim Vis As String = CStr(reader("Vis_no"))
            Dim EDD As String = CStr(reader("EDDDate"))
            Dim item As String = String.Format("{0} : {1}", Vis, EDD)
            Me.ListBox3.Items.Add(item).ToString()
        End While

        reader.Close()
        If connection.State = ConnectionState.Open Then connection.Close()

        TabControl1.SelectedTab = Me.TabPage2
        Label48.Text = "Expected Date Of Delivery"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT * FROM Gyn WHERE (EDDDate >= ? AND ? >= EDDDate) AND (Gyn = 0)
                               ORDER BY EDDDate DESC", con)
        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn")
        DataGridView1.DataSource = ds.Tables("Gyn").DefaultView

        con.Close()
        Label47.Text = ((DataGridView1.Rows.Count) - 1).ToString()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        'Dim dgv As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim dgv As DataGridView
        dgv = DataGridView1
        'txtNo.Text = dgv.Cells(0).Value.ToString
        'txtVis1.Text = dgv.Cells(1).Value.ToString
        'txtG.Text = dgv.Cells(2).Value.ToString
        'txtP.Text = dgv.Cells(3).Value.ToString
        'txtA.Text = dgv.Cells(4).Value.ToString
        'chbxNVD.Checked = CBool(dgv.Cells(5).Value.ToString)
        'chbxCS.Checked = CBool(dgv.Cells(6).Value.ToString)
        'cbxHPOC.Text = dgv.Cells(7).Value.ToString
        'cbxLD.Text = dgv.Cells(8).Value.ToString
        'cbxLC.Text = dgv.Cells(9).Value.ToString
        'DTPickerMns.Value = CDate(dgv.Cells(10).Value.ToString)
        'DTPickerLMP.Value = CDate(dgv.Cells(11).Value.ToString)
        'DTPickerEDD.Value = CDate(dgv.Cells(12).Value.ToString)
        'txtElapsed.Text = dgv.Cells(13).Value.ToString
        'txtGA.Text = dgv.Cells(14).Value.ToString
        'cbxMedH1.Text = dgv.Cells(15).Value.ToString
        'cbxMedH2.Text = dgv.Cells(16).Value.ToString
        'cbxMedH3.Text = dgv.Cells(17).Value.ToString
        'cbxSurH1.Text = dgv.Cells(18).Value.ToString
        'cbxSurH2.Text = dgv.Cells(19).Value.ToString
        'cbxSurH3.Text = dgv.Cells(20).Value.ToString
        'cbxGynH1.Text = dgv.Cells(21).Value.ToString
        'cbxGynH2.Text = dgv.Cells(22).Value.ToString
        'cbxGynH3.Text = dgv.Cells(23).Value.ToString
        'cbxDrugH1.Text = dgv.Cells(24).Value.ToString
        'cbxDrugH2.Text = dgv.Cells(25).Value.ToString
        'cbxDrugH3.Text = dgv.Cells(26).Value.ToString
        'chbxGyn.Checked = CBool(dgv.Cells(27).Value.ToString)

        'txtNo.Text = dgv.Rows(0).Cells("Patient_no").Value.ToString
        'txtVis1.Text = dgv.Rows(1).Cells("Vis_no").Value.ToString
        'txtG.Text = dgv.Rows(2).Cells("G").Value.ToString
        'txtP.Text = dgv.Rows(3).Cells("P").Value.ToString
        'txtA.Text = dgv.Rows(4).Cells("A").Value.ToString
        'chbxNVD.Checked = CBool(dgv.Rows(5).Cells("NVD").Value.ToString)
        'chbxCS.Checked = CBool(dgv.Rows(6).Cells("CS").Value.ToString)
        'cbxHPOC.Text = dgv.Rows(7).Cells("HPOC").Value.ToString
        'cbxLD.Text = dgv.Rows(8).Cells("LD").Value.ToString
        'cbxLC.Text = dgv.Rows(9).Cells("LC").Value.ToString
        'DTPickerMns.Value = CDate(dgv.Rows(10).Cells("MnsDate").Value.ToString)
        'DTPickerLMP.Value = CDate(dgv.Rows(11).Cells("LMPDate").Value.ToString)
        'DTPickerEDD.Value = CDate(dgv.Rows(12).Cells("EDDDate").Value.ToString)
        'txtElapsed.Text = dgv.Rows(13).Cells("ElapW").Value.ToString
        'txtGA.Text = dgv.Rows(14).Cells("GAW").Value.ToString
        'cbxMedH1.Text = dgv.Rows(15).Cells("MedH1").Value.ToString
        'cbxMedH2.Text = dgv.Rows(16).Cells("MedH2").Value.ToString
        'cbxMedH3.Text = dgv.Rows(17).Cells("MedH3").Value.ToString
        'cbxSurH1.Text = dgv.Rows(18).Cells("SurH1").Value.ToString
        'cbxSurH2.Text = dgv.Rows(19).Cells("SurH2").Value.ToString
        'cbxSurH3.Text = dgv.Rows(20).Cells("SurH3").Value.ToString
        'cbxGynH1.Text = dgv.Rows(21).Cells("GynH1").Value.ToString
        'cbxGynH2.Text = dgv.Rows(22).Cells("GynH2").Value.ToString
        'cbxGynH3.Text = dgv.Rows(23).Cells("GynH3").Value.ToString
        'cbxDrugH1.Text = dgv.Rows(24).Cells("DrugH1").Value.ToString
        'cbxDrugH2.Text = dgv.Rows(25).Cells("DrugH2").Value.ToString
        'cbxDrugH3.Text = dgv.Rows(26).Cells("DrugH3").Value.ToString
        'chbxGyn.Checked = CBool(dgv.Rows(27).Cells("Gyn").Value.ToString)

        'TabControl1.SelectedTab = Me.TabPage1
        'GynEnabled()


    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Dim dgv As DataGridViewRow = DataGridView1.SelectedRows(0)
        'Dim dgv As DataGridView
        'dgv = DataGridView1
        txtNo.Text = dgv.Cells(0).Value.ToString
        txtVis1.Text = dgv.Cells(1).Value.ToString
        txtG.Text = dgv.Cells(2).Value.ToString
        txtP.Text = dgv.Cells(3).Value.ToString
        txtA.Text = dgv.Cells(4).Value.ToString
        chbxNVD.Checked = CBool(dgv.Cells(5).Value.ToString)
        chbxCS.Checked = CBool(dgv.Cells(6).Value.ToString)
        cbxHPOC.Text = dgv.Cells(7).Value.ToString
        cbxLD.Text = dgv.Cells(8).Value.ToString
        cbxLC.Text = dgv.Cells(9).Value.ToString
        DTPickerMns.Value = CDate(dgv.Cells(10).Value.ToString)
        DTPickerLMP.Value = CDate(dgv.Cells(11).Value.ToString)
        DTPickerEDD.Value = CDate(dgv.Cells(12).Value.ToString)
        txtElapsed.Text = dgv.Cells(13).Value.ToString
        txtGA.Text = dgv.Cells(14).Value.ToString
        cbxMedH1.Text = dgv.Cells(15).Value.ToString
        cbxMedH2.Text = dgv.Cells(16).Value.ToString
        cbxMedH3.Text = dgv.Cells(17).Value.ToString
        cbxSurH1.Text = dgv.Cells(18).Value.ToString
        cbxSurH2.Text = dgv.Cells(19).Value.ToString
        cbxSurH3.Text = dgv.Cells(20).Value.ToString
        cbxGynH1.Text = dgv.Cells(21).Value.ToString
        cbxGynH2.Text = dgv.Cells(22).Value.ToString
        cbxGynH3.Text = dgv.Cells(23).Value.ToString
        cbxDrugH1.Text = dgv.Cells(24).Value.ToString
        cbxDrugH2.Text = dgv.Cells(25).Value.ToString
        cbxDrugH3.Text = dgv.Cells(26).Value.ToString
        chbxGyn.Checked = CBool(dgv.Cells(27).Value.ToString)

        'txtNo.Text = dgv.Rows(0).Cells("Patient_no").Value.ToString
        'txtVis1.Text = dgv.Rows(1).Cells("Vis_no").Value.ToString
        'txtG.Text = dgv.Rows(2).Cells("G").Value.ToString
        'txtP.Text = dgv.Rows(3).Cells("P").Value.ToString
        'txtA.Text = dgv.Rows(4).Cells("A").Value.ToString
        'chbxNVD.Checked = CBool(dgv.Rows(5).Cells("NVD").Value.ToString)
        'chbxCS.Checked = CBool(dgv.Rows(6).Cells("CS").Value.ToString)
        'cbxHPOC.Text = dgv.Rows(7).Cells("HPOC").Value.ToString
        'cbxLD.Text = dgv.Rows(8).Cells("LD").Value.ToString
        'cbxLC.Text = dgv.Rows(9).Cells("LC").Value.ToString
        'DTPickerMns.Value = CDate(dgv.Rows(10).Cells("MnsDate").Value.ToString)
        'DTPickerLMP.Value = CDate(dgv.Rows(11).Cells("LMPDate").Value.ToString)
        'DTPickerEDD.Value = CDate(dgv.Rows(12).Cells("EDDDate").Value.ToString)
        'txtElapsed.Text = dgv.Rows(13).Cells("ElapW").Value.ToString
        'txtGA.Text = dgv.Rows(14).Cells("GAW").Value.ToString
        'cbxMedH1.Text = dgv.Rows(15).Cells("MedH1").Value.ToString
        'cbxMedH2.Text = dgv.Rows(16).Cells("MedH2").Value.ToString
        'cbxMedH3.Text = dgv.Rows(17).Cells("MedH3").Value.ToString
        'cbxSurH1.Text = dgv.Rows(18).Cells("SurH1").Value.ToString
        'cbxSurH2.Text = dgv.Rows(19).Cells("SurH2").Value.ToString
        'cbxSurH3.Text = dgv.Rows(20).Cells("SurH3").Value.ToString
        'cbxGynH1.Text = dgv.Rows(21).Cells("GynH1").Value.ToString
        'cbxGynH2.Text = dgv.Rows(22).Cells("GynH2").Value.ToString
        'cbxGynH3.Text = dgv.Rows(23).Cells("GynH3").Value.ToString
        'cbxDrugH1.Text = dgv.Rows(24).Cells("DrugH1").Value.ToString
        'cbxDrugH2.Text = dgv.Rows(25).Cells("DrugH2").Value.ToString
        'cbxDrugH3.Text = dgv.Rows(26).Cells("DrugH3").Value.ToString
        'chbxGyn.Checked = CBool(dgv.Rows(27).Cells("Gyn").Value.ToString)

        TabControl1.SelectedTab = Me.TabPage1
        TextBox5.Text = cbxPatName.Text
        TextBox6.Text = txtNo.Text
        GynEnabled()

    End Sub

    Private Sub txtNo_TextChanged(sender As Object, e As EventArgs) Handles txtNo.TextChanged
        ShowPatTable()
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim x As Integer
        Dim y As Integer
        x = CInt((Me.Width - Panel2.Width) / 2)
        y = CInt((Me.Height - Panel2.Height) / 2)
        Panel2.Location = New Point(x, y)

    End Sub
End Class



''##https://social.msdn.microsoft.com/Forums/vstudio/en-US/b2a15b26-6d51-49d5-81cf-20fef70e8316/when-datetimepicker-value-changed-this-error-occured?forum=vbgeneral
Public Class DateTimeCalculations
    Public Sub ToAgeString(ByVal dob As Date)
        Dim today As Date = Date.Today

        Dim months As Integer = today.Month - dob.Month
        Dim years As Integer = today.Year - dob.Year

        If today.Day < dob.Day Then
            months -= 1
        End If

        If months < 0 Then
            years -= 1
            months += 12
        End If

        Dim days As Integer = (today - dob.AddMonths((years * 12) + months)).Days

        mMonths = months
        mDays = days
        mYears = years

        mFormatted = String.Format("{0} Y{1}, {2} M{3} and {4} day{5}",
                                   years, If(years = 1, "", "s"),
                                   months, If(months = 1, "", "s"),
                                   days, If(days = 1, "", "s"))

    End Sub
    Private mFormatted As String
    Public ReadOnly Property Formatted As String
        Get
            Return mFormatted
        End Get
    End Property
    Private mMonths As Integer
    Public ReadOnly Property Months As Integer
        Get
            Return mMonths
        End Get
    End Property
    Private mDays As Integer
    Public ReadOnly Property Days As Integer
        Get
            Return mDays
        End Get
    End Property
    Private mYears As Integer
    Public ReadOnly Property Years As Integer
        Get
            Return mYears
        End Get
    End Property
End Class