'Option Strict On
Option Explicit On

Imports System.Data.OleDb
Imports System.IO
Imports System.Xml
Imports System.Collections
Imports System.Management
Imports ApplicationEnhancement.Expiry
Imports System.ComponentModel

Public Class Form1

    Inherits System.Windows.Forms.Form
    '("provider=microsoft.ace.oledb.12.0; data source=|DataDirectory|\TestDB.accdb;jet oledb:database password=mero1981923")
    Dim cs As String = "provider=microsoft.ace.oledb.12.0; data source=" & Application.StartupPath & "\TestDB.accdb;jet oledb:database password=hgpl]GGI"
    Dim conn As New OleDbConnection(cs)
    Dim cmd As New OleDbCommand
    Public dr As OleDbDataReader

    Dim f2 As Form2

    Dim hddserial As System.STAThreadAttribute()

    'To make the App. to open a fixed times
    'Private WithEvents usage As ApplicationUsage
    'Private _maxTimes As Integer = 1
    'Private _usageLimitExceeded As Boolean = False

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
        'HDDSer()
        'Expire()
        'If TextBox3.Text <> "amr_bakry" Then
        '    MsgBox("You Are Not Authorized," + vbCrLf +
        '           "Please Call 01067174141", MsgBoxStyle.Exclamation, "Error")
        '    End
        'End If

        Label81.Text = "Patient's Data"

        Me.AutoScroll = True
        DTPNow()
        DTPickerNow()
        DateTimePicker1.Value = Now
        DTPicker.Value = Now
        DateTimePicker2.Value = Now
        DateTimePicker3.Value = Now

        LoadPicture()

        FillAuto()
        'txtPatName.Select()
        'txtVisNo.Select()

    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim x As Integer
        Dim y As Integer
        x = CInt((Me.Width - Panel2.Width) / 2)
        y = CInt((Me.Height - Panel2.Height) / 2) + 40
        Panel2.Location = New Point(x, y)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Expire()
        'HDDSer()
        'If TextBox3.Text = TextBox15.Text Then
        '    Exit Sub
        'End If
        'MsgBox("You Are Not Authorized," + vbCrLf +
        '           "Please Call 01067174141", MsgBoxStyle.Exclamation, "Compatability Failed")
        'End

        'If TextBox3.Text <> "amr_bakry" Then
        '    MsgBox("You Are Not Authorized, Please Call 01067174141")
        '    End
        'End If
        ''If DateTimePicker4.Value < Now Then
        ''DateTimePicker4.Value = DateTimePicker4.Value.AddDays(50)
        ''End If

        'Label81.Text = "Patient's Data"

        '_maxTimes = CInt(Val(My.Settings.MyAppScopedSetting))
        'DateGood(CInt(Val(My.Settings.Date_Days)))

        '' Initialize the variable "usage" by using
        '' the NEW keyword along with the number
        '' of maximum "hits" you want to allow:
        ''usage = New ApplicationUsage(CInt(_maxTimes))
        'usage = New ApplicationUsage(_maxTimes)

        '' Now check the usage. If the usage has been
        '' exceeded, the "MaximumExceeded" event will
        '' be raised:
        'usage.CheckUsage()

        'If _usageLimitExceeded Then
        '    MessageBox.Show(String.Format("The maximum usage of " & vbCrLf &
        '                       "{0: n0} times has been exceeded." & vbCrLf &
        '                       "For full version Call 01067174141 Or 01149003573",
        '                        _maxTimes), "Cannot Continue",
        '                       MessageBoxButtons.OK)
        '    'Exit Sub
        '    Close()
        '    End

        'Else
        '    ' Just for demonstration here, if the
        '    ' maximum has not been exceeded then
        '    ' I'll just show the quantity of times
        '    ' this program has been run.

        '    Label46.Text = String.Format("{0:n0} Of " &
        '                   _maxTimes, usage.UsageQuantity)      'CType(usage.UsageQuantity, String)
        '    'MessageBox.Show(String.Format("Usage Quantity : {0:n0} Of " &
        '    'CInt(Val(My.Settings.MyAppScopedSetting)), usage.UsageQuantity))
        'End If

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

    'Private Sub _
    '    usage_MaximumExceeded(sender As Object,
    '                          e As System.EventArgs) _
    '                          Handles usage.MaximumExceeded

    '    _usageLimitExceeded = True

    'End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim databasepath As String = Path.Combine(Application.StartupPath,
                                                Directory.GetCurrentDirectory + "\TestDB.accdb")
        Dim backuppath As String = Path.Combine(Application.StartupPath,
                                                  Directory.GetCurrentDirectory + "\Backups\_" &
                                                                            Format(Now(), "dd_MM_yyyy_hhmmtt") & ".accdb")
        Dim backuppath1 As String = Path.Combine(Application.StartupPath,
                                                  Directory.GetCurrentDirectory + "\Backups\Docs\TestDB_" &
                                                                            Format(Now(), "MM_yyyy") & ".accdb")

        My.Computer.FileSystem.CopyFile(databasepath, backuppath, True)
        My.Computer.FileSystem.CopyFile(databasepath, backuppath1, True)

        RDXmlPatNames()
        RDXmlPatNames1()
        RDXmlPatNames2()
        RDXmlDrugs()
        RDXmlDiaInter()
        RDXmlInv()
        RDXmlInv2()
        RDXmlInvRes()
        RDXmlPlan()
        RDXmlJobs()
        RDXmlMob()


        'If e.CloseReason = CloseReason.UserClosing Then
        '    usage.Save()
        'End If
        'End
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Location = New Point(0, 0)
        Me.Size = Screen.PrimaryScreen.WorkingArea.Size
        'txtPatName.Select()
        'ShowPatTable()
        GynDisabled()
        Gyn2Disabled()
        Expire()
        If TextBox3.Text <> "amr_bakry" Then
            MsgBox("You Are Not Authorized," + vbCrLf +
                   "Please Call 01067174141", MsgBoxStyle.Exclamation, "Not Authorized")
            End
        End If
        'If TextBox3.Text = TextBox15.Text Then
        '    Exit Sub
        'End If
        'MsgBox("You Are Not Authorized," + vbCrLf +
        '           "Please Call 01067174141", MsgBoxStyle.Exclamation, "Compatability Failed")
        'End

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

    Private Sub Expire()
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "SELECT * FROM AttFile"
                'cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVis.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        DateTimePicker4.Value = CDate(dt.Rows(0).Item("AttDate").ToString)
                        TextBox3.Text = dt.Rows(0).Item("ex_pire").ToString
                    End If
                End Using
            End Using
        End Using

    End Sub

    Sub HDDSer()
        Dim HDD_Serial As String

        Dim hdd As New ManagementObjectSearcher("select * from Win32_DiskDrive")

        For Each hd In hdd.Get

            HDD_Serial = hd("SerialNumber")
            'MsgBox(HDD_Serial)
            TextBox15.Text = HDD_Serial
        Next
    End Sub
    'Private Sub HDDSer()
    '    Dim hdCollection As New ArrayList()
    '    Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")
    '    Dim wmi_HD As New ManagementObject()

    '    For Each wmi_HD In searcher.Get

    '        Dim hd As New Class1.HardDrive()

    '        hd.Model = wmi_HD("Model").ToString()
    '        hd.Type = wmi_HD("InterfaceType").ToString()
    '        hdCollection.Add(hd)
    '    Next

    '    Dim searcher1 As New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia")

    '    Dim i As Integer = 0
    '    For Each wmi_HD In searcher1.Get()

    '        '// get the hard drive from collection
    '        '// using index

    '        Dim hd As Class1.HardDrive
    '        hd = hdCollection(i)

    '        '// get the hardware serial no.
    '        If wmi_HD("SerialNumber") = "" Then
    '            hd.serialNo = "None"
    '        Else
    '            hd.serialNo = wmi_HD("SerialNumber").ToString()
    '            i += 1
    '        End If
    '    Next

    '    Dim hd1 As Class1.HardDrive
    '    Dim ii As Integer = 0

    '    For Each hd1 In hdCollection
    '        ii += 1

    '        TextBox15.Text = TextBox15.Text + "Serial No: " + hd1.serialNo + Chr(13) + Chr(10) + Chr(13) + Chr(10)
    '        'TextBox15.Text = TextBox15.Text + hd1.serialNo
    '    Next
    'End Sub

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
        'Trace.WriteLine("DTPickerNow STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        DTPicker.Value = Now
        DTPickerMns.Value = Now
        DTPickerLMP.Value = Now
        DTPickerEDD.Value = Now
        DTPickerAtt.Value = Now
        DateTimePicker2.Value = Now
        DateTimePicker3.Value = Now
        DateTimePicker5.Value = Now
        DateTimePicker6.Value = Now

        'Trace.WriteLine("DTPickerNow FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Sub

    Private Sub FillAuto()
        txtNo.Text = GetAutonumber("Pat", "Patient_no")
        txtVis1.Text = GetAutonumber("Gyn", "Vis_no")
        txtVis.Text = GetAutonumber("Gyn2", "Vis_no")
        txtVisNo.Text = GetAutonumber("Visits", "Visit_no")
    End Sub

    Private Sub LoadPicture()
        'Dim PicPath As String = "D:\KMAClinic\Photos\GynClinic.png"
        'PicPath = My.Settings.PicFilePath
        Dim filename As String = Path.Combine(Application.StartupPath, "sabah2.png") 'Path.GetFileName("\DrAhEssmat.png")
        Me.PictureBox1.Image = Image.FromFile(filename)
    End Sub

    Function GetTable(SelectCommand As String) As DataTable
        'Trace.WriteLine("GetTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        'Trace.WriteLine("GetTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Function

    Function GetAutonumber(TableName As String, ColumnName As String) As String
        'Trace.WriteLine("GetAutonumber STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        'Trace.WriteLine("GetAutonumber FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Function

    Sub ClearGyn()
        'Trace.WriteLine("ClearGyn STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        txtVis1.Text = ""
        txtA.Text = ""
        txtG.Text = ""
        txtP.Text = ""
        TextBox6.Text = ""
        TextBox5.Text = ""
        chbxGyn.Checked = False
        chbxNVD.Checked = False
        chbxCS.Checked = False
        txtGA.Text = ""
        txtElapsed.Text = ""

        DTPickerMns.Value = Now
        DTPickerLMP.Value = Now
        DTPickerEDD.Value = Now

        cbxLD.Text = ""
        cbxLC.Text = ""
        cbxHPOC.Text = ""
        cbxMedH1.Text = ""
        cbxMedH2.Text = ""
        cbxMedH3.Text = ""
        cbxSurH1.Text = ""
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

        'Trace.WriteLine("ClearGyn FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Sub

    Sub ClearGyn2()
        'Trace.WriteLine("ClearGyn2 STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        txtVis.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
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

        'Trace.WriteLine("ClearGyn2 FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Sub

    Private Sub btnclear_Click(sender As Object, e As EventArgs) Handles btnclear.Click
        'Trace.WriteLine("btnclear_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If Label81.Text = "Patient's Data" Then
            cbxSearch.Text = ""
            txtPatName.Text = ""
            cbxJob.Text = ""
            cbxAddress.Text = ""
            DTPicker.Value = Now
            txtAge.Text = ""
            txtPhone.Text = ""
            txtHusband.Text = ""
            cbxHusJob.Text = ""
            txtNo.Text = GetAutonumber("Pat", "Patient_no")
        ElseIf Label81.Text = "Expected Date Of Delivery" Then
            TextBox9.Text = ""
            DataGridView3.DataSource = Nothing
            Label53.Text = "0"
            DataGridView5.DataSource = Nothing
            Label78.Text = "0"
            DataGridView6.DataSource = Nothing
            Label80.Text = "0"
            DataGridView7.DataSource = Nothing
        ElseIf Label81.Text = "History" Then
            ClearGyn()
            ClearGyn2()
        ElseIf Label81.Text = "Income" Then
            TextBox11.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            DataGridView9.DataSource = Nothing
            Label86.Text = "0"
            DataGridView10.DataSource = Nothing
            Label85.Text = "0"
        ElseIf Label81.Text = "Visits" Then
            ListBox4.Items.Clear()
            txt1.Text = ""
            ClearData()
            ClearDrug()
            ClearInv()
            btnNewVisit.Enabled = False
            rdoVisit.Checked = True
            cbxVisSearch.Text = ""
            InvAndAttDisabled()
            DrugDisabled()
            lblcurTime.Text = Now.ToShortDateString
        End If

        'cbxSearch.Text = ""
        'txtPatName.Text = ""
        'cbxJob.Text = ""
        'cbxAddress.Text = ""
        'DTPicker.Value = Now
        'txtAge.Text = ""
        'txtPhone.Text = ""
        'txtHusband.Text = ""
        'cbxHusJob.Text = ""

        TextBox1.Text = ""
        TextBox2.Text = ""
        'TextBox5.Text = ""
        'TextBox6.Text = ""
        'TextBox7.Text = ""
        'TextBox8.Text = ""
        'TextBox9.Text = ""
        DateTimePicker2.Value = Now
        DateTimePicker3.Value = Now

        ''##This line must come after the 'for loop' because for loop erase every textbox in the form
        ''##And the autonumber come after that
        txtNo.Text = GetAutonumber("Pat", "Patient_no")
        txtPatName.Select()
        ClearGyn()
        ClearGyn2()

        DTPNow()

        GynDisabled()
        Gyn2Disabled()
        'DataGridView1.DataSource = Nothing
        'Label48.Text = "0"
        'DataGridView2.DataSource = Nothing
        'Label49.Text = "0"
        'DataGridView8.DataSource = Nothing
        'Label82.Text = "0"
        'FillDGV1()
        'FillDGV2()
        'FillDGV3()
        'FillDGV5()
        'FillDGV6()
        'DGVPatients()
        'FillDGV7()
        'btnL.Enabled = True
        'btnB.Enabled = True
        'Trace.WriteLine("btnclear_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowGyn2Table()

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

    Sub ShowGynTabletxtVis1()
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con

                'cmd.CommandText = "SELECT * FROM Gyn WHERE Patient_no=@Patient_no"
                cmd.CommandText = "SELECT * FROM Gyn WHERE Vis_no=@Vis_no"
                cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVis1.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        TextBox6.Text = dt.Rows(0).Item("Patient_no").ToString
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


    Sub ShowGynTable()
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con

                'cmd.CommandText = "SELECT * FROM Gyn WHERE Patient_no=@Patient_no"
                cmd.CommandText = "SELECT * FROM Gyn WHERE Vis_no=@Vis_no"
                cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(TextBox1.Text))
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
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con                               ''##

                cmd.CommandText = "SELECT * FROM Pat WHERE Patient_no=@Patient_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        txtPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxJob.Text = dt.Rows(0).Item("Job").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        txtHusband.Text = dt.Rows(0).Item("HusName").ToString
                        cbxHusJob.Text = dt.Rows(0).Item("HusJob").ToString

                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub ShowPatTableTextBox6()
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con                               ''##

                cmd.CommandText = "SELECT * FROM Pat WHERE Patient_no=@Patient_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(TextBox6.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        txtPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxJob.Text = dt.Rows(0).Item("Job").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        txtHusband.Text = dt.Rows(0).Item("HusName").ToString
                        cbxHusJob.Text = dt.Rows(0).Item("HusJob").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        conn.Open()
        txtNo.ResetText()
        txtPatName.ResetText()
        cbxAddress.ResetText()
        DTPicker.ResetText()
        txtAge.ResetText()
        txtPhone.ResetText()
        txtHusband.ResetText()
        Dim str As String = "SELECT * FROM [Pat] " &
        "WHERE Patient_no LIKE '%" & Me.cbxSearch.Text & "%' " &
        "ORDER BY Patient_no DESC"
        Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        dr = cmd.ExecuteReader
        While dr.Read
            txtNo.Text = dr("Patient_no").ToString
            txtPatName.Text = dr("Name").ToString
            cbxAddress.Text = dr("Address").ToString
            DTPicker.Text = dr("Birthdate").ToString
            txtAge.Text = dr("Age").ToString
            txtPhone.Text = dr("Phone").ToString
            txtHusband.Text = dr("HusName").ToString
        End While
        conn.Close()
    End Sub

    Private Sub cbxSearch_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxSearch.MouseClick
        If rdoName.Checked Then
            InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSearch.Items.AddRange(cbElements)
        ElseIf rdoHus.Checked Then
            InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
            '' Read the XML file from disk only once
            Dim xDoc1 = XElement.Load(Application.StartupPath + "\PatNames1.xml")
            '' Parse the XML document only once
            Dim cbElements1 = xDoc1.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSearch.Items.AddRange(cbElements1)
        ElseIf rdoPhone.Checked Then
            '' Read the XML file from disk only once
            Dim xDoc2 = XElement.Load(Application.StartupPath + "\Mob.xml")
            '' Parse the XML document only once
            Dim cbElements2 = xDoc2.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSearch.Items.AddRange(cbElements2)
        End If
        'btnF.Enabled = True
        'btnL.Enabled = True
    End Sub

    Private Sub cbxSearch_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxSearch.Validating
        'ClearGyn()

        If rdoName.Checked Then
            SearchName()
        ElseIf rdoID.Checked Then
            SearchID()
        ElseIf rdoHus.Checked Then
            SearchHusband()
        ElseIf rdoPhone.Checked Then
            SearchPhone()
        End If
        'ListBox1.Items.Clear()
        'ListBox2.Items.Clear()
        GynDisabled()
        Gyn2Disabled()
        'showgyn
        'FillDGV1()
        'FillDGV2()
        'FillDGV8()
        txtNo.Select()
        ''##https://social.msdn.microsoft.com/Forums/vstudio/en-US/b2a15b26-6d51-49d5-81cf-20fef70e8316/when-datetimepicker-value-changed-this-error-occured?forum=vbgeneral
        operations.ToAgeString(DTPicker.Value)
        txtAge.Text = operations.Formatted

        'Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        'Dim date2 As Date = DTPickerLMP.Value
        'Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        'If DTPickerEDD.Value = DTPickerLMP.Value Then
        '    Exit Sub
        'End If
        'txtElapsed.Text = CStr(weeks) '& "  Weeks"
        'txtGA.Text = CStr(40 - weeks)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
        'txtPatName.Focus()
    End Sub

    Sub FillDGV1()
        'If txtVis1.Text <> GetAutonumber("Gyn", "Vis_no") Then

        Label48.Text = "Previous Visits"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(MnsDate)AS[Visit Date],(NVD)AS[NVD],(CS)AS[CS],
                                (G)AS[G],(P)AS[P],(A)AS[A],(HPOC)AS[Previous Obstetric Complications],(LD)AS[LD],
                                (LC)AS[LC],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],
                                (ElapW)AS[Gestational age],(GAW)AS[Remaining],(MedH1)AS[Medical History1],(MedH2)AS[Medical History2],(MedH3)AS[Medical History3],
                                (SurH1)AS[Surgical History1],(SurH2)AS[Surgical History2],(SurH3)AS[Surgical History3],(GynH1)AS[Gynecological History1],
                                (GynH2)AS[Gynecological History2],(GynH3)AS[Gynecological History3],(DrugH1)AS[Drug History1],(DrugH2)AS[Drug History2],(DrugH3)AS[Drug History3],(Gyn)AS[Gyna]
                                FROM Gyn WHERE Patient_no = @Patient_no ORDER BY Vis_no DESC", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn")
        DataGridView1.DataSource = ds.Tables("Gyn").DefaultView

        con.Close()
        Label47.Text = (DataGridView1.Rows.Count).ToString()
        'End If
    End Sub

    Sub FillDGV2()
        Label49.Text = "U/S Visits"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(AttDt)AS[Visit Date],(GL)AS[General Look],(Pls)AS[Puls],
                               (BP)AS[Blood Pressure],(Wt)AS[Weight],(BdBt)AS[Body Built],(ChtH)AS[Chest and Heart],
                               (HdNe)AS[Head and Neck],(Ext)AS[Extremities],(FunL)AS[Fundal Level],(Scrs)AS[Scars],
                               (Edm)AS[Edema],(US)AS[Ultra Sound],(Amount)AS[Amount]
                               FROM Gyn2 WHERE Patient_no = @Patient_no
                               ORDER BY Vis_no DESC", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn2")
        DataGridView2.DataSource = ds.Tables("Gyn2").DefaultView
        'MsgBox("Patient ID = " & txtNo.Text)
        con.Close()
        Label50.Text = (DataGridView2.Rows.Count).ToString()
    End Sub

    Sub FillDGV3()
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],(ElapW)AS[Gestational age],
                              (GAW)AS[Remaining] FROM Gyn WHERE (EDDDate >= ?) AND (? >= EDDDate) AND (GAW <= 4) AND (GAW > -1) AND (EDDDate <> LMPDate) AND (Gyn = 0)
                               ORDER BY EDDDate DESC", con)
        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn")
        DataGridView3.DataSource = ds.Tables("Gyn").DefaultView

        con.Close()
        Label53.Text = (DataGridView3.Rows.Count).ToString()
    End Sub

    Sub FillDGV6()
        'If txtVis1.Text <> GetAutonumber("Gyn", "Vis_no") Then

        Label79.Text = "Previous Visits"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(MnsDate)AS[Visit Date],(NVD)AS[NVD],(CS)AS[CS],
                                (G)AS[G],(P)AS[P],(A)AS[A],(HPOC)AS[Previous Obstetric Complications],(LD)AS[LD],
                                (LC)AS[LC],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],
                                (ElapW)AS[Gestational age],(GAW)AS[Remaining],(MedH1)AS[Medical History1],(MedH2)AS[Medical History2],(MedH3)AS[Medical History3],
                                (SurH1)AS[Surgical History1],(SurH2)AS[Surgical History2],(SurH3)AS[Surgical History3],(GynH1)AS[Gynecological History1],
                                (GynH2)AS[Gynecological History2],(GynH3)AS[Gynecological History3],(DrugH1)AS[Drug History1],(DrugH2)AS[Drug History2],(DrugH3)AS[Drug History3],(Gyn)AS[Gyna]
                                FROM [Gyn] WHERE Patient_no = @Patient_no ORDER BY Vis_no DESC", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn")
        DataGridView6.DataSource = ds.Tables("Gyn").DefaultView

        con.Close()
        Label80.Text = (DataGridView6.Rows.Count).ToString()
        'End If
    End Sub


    Sub FillDGV6Empty()
        'If txtVis1.Text <> GetAutonumber("Gyn", "Vis_no") Then

        Label79.Text = "Previous Visits"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(MnsDate)AS[Visit Date],(NVD)AS[NVD],(CS)AS[CS],
                                (G)AS[G],(P)AS[P],(A)AS[A],(HPOC)AS[Previous Obstetric Complications],(LD)AS[LD],
                                (LC)AS[LC],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],
                                (ElapW)AS[Gestational age],(GAW)AS[Remaining],(MedH1)AS[Medical History1],(MedH2)AS[Medical History2],(MedH3)AS[Medical History3],
                                (SurH1)AS[Surgical History1],(SurH2)AS[Surgical History2],(SurH3)AS[Surgical History3],(GynH1)AS[Gynecological History1],
                                (GynH2)AS[Gynecological History2],(GynH3)AS[Gynecological History3],(DrugH1)AS[Drug History1],(DrugH2)AS[Drug History2],(DrugH3)AS[Drug History3],(Gyn)AS[Gyna]
                                FROM Gyn WHERE Patient_no = @Patient_no", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(TextBox9.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn")
        DataGridView6.DataSource = ds.Tables("Gyn").DefaultView

        con.Close()
        Label80.Text = (DataGridView6.Rows.Count).ToString()
        'End If
    End Sub

    Sub FillDGV5()
        Label77.Text = "U/S Visits"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(AttDt)AS[Visit Date],(GL)AS[General Look],(Pls)AS[Puls],
                               (BP)AS[Blood Pressure],(Wt)AS[Weight],(BdBt)AS[Body Built],(ChtH)AS[Chest and Heart],
                               (HdNe)AS[Head and Neck],(Ext)AS[Extremities],(FunL)AS[Fundal Level],(Scrs)AS[Scars],
                               (Edm)AS[Edema],(US)AS[Ultra Sound],(Amount)AS[Amount]
                               FROM Gyn2 WHERE Patient_no = @Patient_no
                               ORDER BY Vis_no DESC", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn2")
        DataGridView5.DataSource = ds.Tables("Gyn2").DefaultView
        'MsgBox("Patient ID = " & txtNo.Text)
        con.Close()
        Label78.Text = (DataGridView5.Rows.Count).ToString()
    End Sub

    Sub FillDGV5Empty()
        Label77.Text = "U/S Visits"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(AttDt)AS[Visit Date],(GL)AS[General Look],(Pls)AS[Puls],
                               (BP)AS[Blood Pressure],(Wt)AS[Weight],(BdBt)AS[Body Built],(ChtH)AS[Chest and Heart],
                               (HdNe)AS[Head and Neck],(Ext)AS[Extremities],(FunL)AS[Fundal Level],(Scrs)AS[Scars],
                               (Edm)AS[Edema],(US)AS[Ultra Sound],(Amount)AS[Amount]
                               FROM Gyn2 WHERE Patient_no = @Patient_no", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(TextBox9.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn2")
        DataGridView5.DataSource = ds.Tables("Gyn2").DefaultView
        'MsgBox("Patient ID = " & txtNo.Text)
        con.Close()
        Label78.Text = (DataGridView5.Rows.Count).ToString()
    End Sub

    Sub FillDGV7()
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Name) AS [Patient Name],(Address) AS [Address],
                                        (Birthdate) AS [Birth Date],(Age) AS [Age],(Phone) AS [Phone],(HusName) AS [Husband Name]
                                        FROM Pat WHERE Patient_no = @Patient_no", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Pat")
        DataGridView7.DataSource = ds.Tables("Pat").DefaultView

        con.Close()
    End Sub

    Sub FillDGV7Empty()
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Name) AS [Patient Name],(Address) AS [Address],
                                        (Birthdate) AS [Birth Date],(Age) AS [Age],(Phone) AS [Phone],(HusName) AS [Husband Name]
                                        FROM Pat WHERE Patient_no = @Patient_no", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(TextBox9.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Pat")
        DataGridView7.DataSource = ds.Tables("Pat").DefaultView

        con.Close()
    End Sub

    Sub FillDGV8()
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Visit_no)AS[Visit No],(Complain)AS[Complain],(Sign)AS[Sign],
                                        (Diagnosis)AS[Diagnosis],(Intervention)AS[Intervention],(Amount)AS[Amount],(VisDate)AS[Visit Date] 
                                        FROM Visits WHERE Patient_no = @Patient_no ORDER BY Visit_no DESC", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Visits")
        DataGridView8.DataSource = ds.Tables("Visits").DefaultView
        Label82.Text = DataGridView8.Rows.Count().ToString
        con.Close()
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
        'txtPatName.ResetText()
        'cbxAddress.ResetText()
        'DTPicker.ResetText()
        'txtAge.ResetText()
        'txtPhone.ResetText()
        'txtHusband.ResetText()
        'Dim str As String = "SELECT * FROM [Pat] " &
        '"WHERE Name LIKE '%" & Me.cbxSearch.Text & "%' " &
        '"ORDER BY Patient_no DESC"
        'Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        'dr = cmd.ExecuteReader
        'While dr.Read
        '    txtNo.Text = dr("Patient_no").ToString
        '    txtPatName.Text = dr("Name").ToString
        '    cbxAddress.Text = dr("Address").ToString
        '    DTPicker.Text = dr("Birthdate").ToString
        '    txtAge.Text = dr("Age").ToString
        '    txtPhone.Text = dr("Phone").ToString
        '    txtHusband.Text = dr("HusName").ToString
        'End While
        'conn.Close()

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

                        txtPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxJob.Text = dt.Rows(0).Item("Job").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        txtHusband.Text = dt.Rows(0).Item("HusName").ToString
                        cbxHusJob.Text = dt.Rows(0).Item("HusJob").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub SearchID()
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

                        txtPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxJob.Text = dt.Rows(0).Item("Job").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        txtHusband.Text = dt.Rows(0).Item("HusName").ToString
                        cbxHusJob.Text = dt.Rows(0).Item("HusJob").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub SearchHusband()
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

                        txtPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxJob.Text = dt.Rows(0).Item("Job").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        txtHusband.Text = dt.Rows(0).Item("HusName").ToString
                        cbxHusJob.Text = dt.Rows(0).Item("HusJob").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub SearchPhone()
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

                        txtPatName.Text = dt.Rows(0).Item("Name").ToString
                        cbxJob.Text = dt.Rows(0).Item("Job").ToString
                        cbxAddress.Text = dt.Rows(0).Item("Address").ToString
                        DTPicker.Text = dt.Rows(0).Item("Birthdate").ToString
                        txtAge.Text = dt.Rows(0).Item("Age").ToString
                        txtPhone.Text = dt.Rows(0).Item("Phone").ToString
                        txtHusband.Text = dt.Rows(0).Item("HusName").ToString
                        cbxHusJob.Text = dt.Rows(0).Item("HusJob").ToString
                    End If
                End Using
            End Using
        End Using
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
                '                       "VALUES(" & txtNo.Text & ", '" & txtPatName.Text & "', '" & cbxAddress.Text & "', '" & DTPicker.Text & "', '" & txtAge.Text & "', '" & txtPhone.Text & "', '" & txtHusband.Text & "')", conn)
                '#Please taking care of the single quote here specially with phone.text and DateTimePicker
                'RunCommand("INSERT INTO Pat(Patient_no, Name, Address, Birthdate, Age, Phone, HusName)" &
                '   "VALUES(" & txtNo.Text & ", '" & txtPatName.Text & "', '" & cbxAddress.Text & "', '" & DTPicker.Text & "', '" & txtAge.Text & "', '" & txtPhone.Text & "', '" & txtHusband.Text & "')")

                cmd = New OleDbCommand("INSERT INTO Pat(Patient_no, Name, Job, Address, Birthdate, Age, Phone, HusName,HusJob)" &
                   "VALUES(@Patient_no, @Name, @Job, @Address, @Birthdate, @Age, @Phone, @HusName, @HusJob)", conn)

                With cmd.Parameters
                    .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                    .Add("@Name", OleDbType.VarChar).Value = txtPatName.Text
                    .Add("@Job", OleDbType.VarChar).Value = cbxJob.Text
                    .Add("@Address", OleDbType.VarChar).Value = cbxAddress.Text
                    .Add("@Birthdate", OleDbType.DBDate).Value = (DTPicker.Value)
                    .Add("@Age", OleDbType.VarChar).Value = txtAge.Text
                    .Add("@Phone", OleDbType.VarChar).Value = txtPhone.Text
                    .Add("@HusName", OleDbType.VarChar).Value = txtHusband.Text
                    .Add("@HusJob", OleDbType.VarChar).Value = cbxHusJob.Text
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
            If txtPatName.Text <> "" And txtNo.Text <> GetAutonumber("Pat", "Patient_no") Then
                'RunCommand("UPDATE Pat SET Name='" & txtPatName.Text & "', Address='" & cbxAddress.Text & "', Birthdate='" & DTPicker.Text & "', Age='" & txtAge.Text & "', Phone='" & txtPhone.Text & "', HusName='" & txtHusband.Text & "' WHERE Patient_no=" & txtNo.Text)
                'RunCommand("UPDATE Visits SET Name='" & txtName.Text & "' WHERE Patient_no=" & frm2.txtVisNo.Text)
                cmd = New OleDbCommand("UPDATE Pat SET Name=@Name, Job=@Job, Address=@Address," &
                                   " Birthdate=@Birthdate, Age=@Age, Phone=@Phone," &
                                   " HusName=@HusName, HusJob=@HusJob WHERE Patient_no=@Patient_no", conn)

                With cmd.Parameters

                    .Add("@Name", OleDbType.VarChar).Value = txtPatName.Text
                    .Add("@Job", OleDbType.VarChar).Value = cbxJob.Text
                    .Add("@Address", OleDbType.VarChar).Value = cbxAddress.Text
                    .Add("@Birthdate", OleDbType.DBDate).Value = (DTPicker.Value)
                    .Add("@Age", OleDbType.VarChar).Value = txtAge.Text
                    .Add("@Phone", OleDbType.VarChar).Value = txtPhone.Text
                    .Add("@HusName", OleDbType.VarChar).Value = txtHusband.Text
                    .Add("@HusJob", OleDbType.VarChar).Value = cbxHusJob.Text
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

    Sub UpdateGAW()
        cmd = New OleDbCommand("UPDATE Gyn SET GAW = @GAW WHERE LMPDate = @LMPDate", conn)
        With cmd.Parameters
            .Add("@GAW", OleDbType.Integer).Value = CInt(Val(txtGA.Text))
            .Add("@LMPDate", OleDbType.DBDate).Value = DTPickerLMP.Value
        End With
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.Open()
        cmd.ExecuteNonQuery()
        conn.Close()

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
                .Add("@GAW", OleDbType.Integer).Value = CInt(Val(txtGA.Text))
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
                .Add("@GAW", OleDbType.Integer).Value = CInt(Val(txtGA.Text))
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
        If txtVis.Text = GetAutonumber("Gyn2", "Vis_no") And txtPatName.Text <> "" Then

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
        Dim DatabasePath As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\TestDB.accdb")

        '##For Real Projects
        Dim DatabasePathCompacted As String = Path.Combine(Application.StartupPath,
                                                           Directory.GetCurrentDirectory +
                                                           "\Backups\_" & Format(Now(),
                                                           "dd_MM_yyyy_hhmmsstt") & ".accdb")

        Dim CompactDB As New Microsoft.Office.Interop.Access.Dao.DBEngine

        '##Here you can write your database password with this method (DatabasePath, DatabasePathCompacted, , , ";pwd=mero1981923")
        CompactDB.CompactDatabase(DatabasePath, DatabasePathCompacted, , , ";pwd=hgpl]GGI")
        CompactDB = Nothing

        Dim backuppath As String = Path.Combine(Application.StartupPath,
                                                Directory.GetCurrentDirectory +
                                                "\Backups\Docs\TestDB_" &
                                                Format(Now(), "MM_yyyy") & ".accdb")
        My.Computer.FileSystem.CopyFile(DatabasePath,
                                        backuppath, True)
        My.Computer.FileSystem.CopyFile(DatabasePathCompacted,
               DatabasePath, True)

    End Sub

    Private Sub BackupXML()
        Dim SourcePatNames As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\PatNames.xml")
        Dim patnames As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\PatNames" + ".xml")

        Dim SourcePatNames1 As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\PatNames1.xml")
        Dim patnames1 As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\PatNames1" + ".xml")

        Dim SourcePatNames2 As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\PatNames2.xml")
        Dim patnames2 As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\PatNames2" + ".xml")

        Dim SourceInvestigations As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Investigations.xml")
        Dim investigations As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\Investigations" + ".xml")

        Dim Sourceinvestigations2 As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Investigations2.xml")
        Dim investigations2 As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\Investigations2" + ".xml")

        Dim SourceDrugs As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Drugs1.xml")
        Dim drugs As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\Drugs1" + ".xml")

        Dim SourcePlans As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Plans.xml")
        Dim plans As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\Plans" + ".xml")

        Dim SourceInvRes As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\InvRes.xml")
        Dim invres As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\InvRes" + ".xml")

        Dim SourceDiaInter As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\DiaInter.xml")
        Dim diainter As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\DiaInter" + ".xml")

        Dim SourceMob As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Mob.xml")
        Dim mob As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\Mob.xml")

        Dim SourceJobs As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Jobs.xml")
        Dim jobs As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\Jobs.xml")

        If Not File.Exists(patnames) Then
            File.Copy(SourcePatNames, patnames)
        ElseIf Not File.Exists(patnames1) Then
            File.Copy(SourcePatNames1, patnames1)
        ElseIf Not File.Exists(patnames2) Then
            File.Copy(SourcePatNames2, patnames2)
        ElseIf Not File.Exists(investigations2) Then
            File.Copy(Sourceinvestigations2, investigations2)
        ElseIf Not File.Exists(drugs) Then
            File.Copy(SourceDrugs, drugs)
        ElseIf Not File.Exists(plans) Then
            File.Copy(SourcePlans, plans)
        ElseIf Not File.Exists(investigations) Then
            File.Copy(SourceInvestigations, investigations)
        ElseIf Not File.Exists(invres) Then
            File.Copy(SourceInvRes, invres)
        ElseIf Not File.Exists(diainter) Then
            File.Copy(SourceDiaInter, diainter)
        ElseIf Not File.Exists(mob) Then
            File.Copy(SourceMob, mob)
        ElseIf Not File.Exists(jobs) Then
            File.Copy(SourceJobs, jobs)
        Else
            My.Computer.FileSystem.CopyFile(SourcePatNames, patnames, True)
            My.Computer.FileSystem.CopyFile(SourcePatNames1, patnames1, True)
            My.Computer.FileSystem.CopyFile(SourcePatNames2, patnames2, True)
            My.Computer.FileSystem.CopyFile(SourceDrugs, drugs, True)
            My.Computer.FileSystem.CopyFile(SourcePlans, plans, True)
            My.Computer.FileSystem.CopyFile(SourceInvRes, invres, True)
            My.Computer.FileSystem.CopyFile(SourceDiaInter, diainter, True)
            My.Computer.FileSystem.CopyFile(Sourceinvestigations2, investigations2, True)
            My.Computer.FileSystem.CopyFile(SourceInvestigations, investigations, True)
            My.Computer.FileSystem.CopyFile(SourceMob, mob, True)
            My.Computer.FileSystem.CopyFile(SourceJobs, jobs, True)
        End If

    End Sub

    Sub GotoVisitPH()
        'Trace.WriteLine("GotoVisitPH started @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        f2 = New Form2(txtPatName.Text, txtNo.Text)
        f2.Show()
        Me.Hide()

        'Trace.WriteLine("GotoVisitPH FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnVisits_Click(sender As Object, e As EventArgs) Handles btnVisits.Click
        'Trace.WriteLine("btnVisits_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtNo.Text <> GetAutonumber("Pat", "Patient_no") Then
            GotoVisitPH()

        ElseIf txtPatName.Text = "" Then

            f2 = New Form2("", "")
            f2.Show()
            Me.Hide()

        End If
        'Trace.WriteLine("btnVisits_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Dim databasepath As String = Path.Combine(Application.StartupPath,
                                                Directory.GetCurrentDirectory + "\TestDB.accdb")
        Dim backuppath As String = Path.Combine(Application.StartupPath,
                                                  Directory.GetCurrentDirectory + "\Backups\_" &
                                                                            Format(Now(), "dd_MM_yyyy_hhmmtt") & ".accdb")
        Dim backuppath1 As String = Path.Combine(Application.StartupPath,
                                                  Directory.GetCurrentDirectory + "\Backups\Docs\TestDB_" &
                                                                            Format(Now(), "MM_yyyy") & ".accdb")

        My.Computer.FileSystem.CopyFile(databasepath, backuppath, True)
        My.Computer.FileSystem.CopyFile(databasepath, backuppath1, True)

        If MsgBox("You Will Exit The Clinic" + vbCrLf +
                  "Are you sure ?", MsgBoxStyle.YesNo,
                  "Confirm Message") = vbNo Then
            Exit Sub
        Else
            RDXmlDiaInter()
            RDXmlDrugs()
            RDXmlPlan()
            RDXmlMob()
            RDXmlJobs()
            RDXmlInv()
            RDXmlInv2()
            RDXmlInvRes()
            RDXmlPatNames()
            RDXmlPatNames1()
            RDXmlPatNames2()
            Application.ExitThread()
        End If


    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        RDXmlDiaInter()
        RDXmlDrugs()
        RDXmlPlan()
        RDXmlMob()
        RDXmlJobs()
        RDXmlInv()
        RDXmlInv2()
        RDXmlInvRes()
        RDXmlPatNames()
        RDXmlPatNames1()
        RDXmlPatNames2()

        loaddata()
    End Sub

    Private Sub btnBackup_Click(sender As Object, e As EventArgs) Handles btnBackup.Click
        If MsgBox("We will make automatic Backup when you close the application," & vbCrLf &
                  "Are you sure to perform this action now?",
                  MsgBoxStyle.YesNo,
                  "Confirmation Backup") = vbNo Then
            Exit Sub
        End If
        CompactAccessDatabase()
        BackupXML()
        Dim BackupFilePath As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups")
        MsgBox("Backup Done @" & vbCrLf & BackupFilePath, MsgBoxStyle.Information, "Backup")
    End Sub

    Private Sub btnNewGyn_Click(sender As Object, e As EventArgs) Handles btnNewGyn.Click
        If txtPatName.Text = "" Then
            Exit Sub
        End If
        ClearGyn()
        TextBox6.Text = txtNo.Text
        TextBox5.Text = txtPatName.Text
        If txtVis1.Text <> GetAutonumber("Gyn", "Vis_no") Then
            txtVis1.Text = GetAutonumber("Gyn", "Vis_no")
            txtVis1.Select()
        ElseIf txtVis1.Text = GetAutonumber("Gyn", "Vis_no") And txtPatName.Text <> "" Then
            SaveGyn()
            txtVis1.Select()
        End If
        GynEnabled()
        btnL.Enabled = True
        txtVis1.BackColor = Color.LightSeaGreen
        txtVis1.ForeColor = Color.White
    End Sub

    Private Sub btnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        If txtPatName.Text = "" Then
            Exit Sub
        End If
        ClearGyn2()
        TextBox7.Text = txtNo.Text
        TextBox8.Text = txtPatName.Text
        If txtVis.Text <> GetAutonumber("Gyn2", "Vis_no") Then
            txtVis.Text = GetAutonumber("Gyn2", "Vis_no")
            txtVis.Select()
        ElseIf txtVis.Text = GetAutonumber("Gyn2", "Vis_no") And txtPatName.Text <> "" Then
            SaveGyn2()
            txtVis.Select()
        End If
        Gyn2Enabled()
        txtVis.BackColor = Color.LightSeaGreen
        txtVis.ForeColor = Color.White

    End Sub


    Private Sub txtPatName_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtPatName.Validating
        Dim ds As DataSet = New DataSet
        Dim da As OleDbDataAdapter = New OleDbDataAdapter("SELECT Name FROM Pat WHERE Name='" & txtPatName.Text & "'", conn)
        da.Fill(ds, "Pat")
        Dim dv As DataView = New DataView(ds.Tables("Pat"))
        Dim cur As CurrencyManager
        cur = CType(Me.BindingContext(dv), CurrencyManager)

        If cur.Count <> 0 And txtNo.Text = GetAutonumber("Pat", "Patient_no") Then
            MsgBox("تأكد من الاسم. هذا الاسم موجود من قبل", MsgBoxStyle.OkOnly, "يجب تغيير الاسم")
            txtPatName.ResetText()
            Exit Sub
        End If
        If txtPatName.Text <> "" Then
            SaveButton()
            'SaveGyn()
            UpdatePatient()
        End If
    End Sub

    Private Sub txtPatName_Click(sender As Object, e As EventArgs) Handles txtPatName.Click
        'If CheckBox1.Checked = False Then
        '    InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        'ElseIf CheckBox1.Checked = True Then
        '    InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        'End If
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)

        '' Read the XML file from disk only once
        'Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames.xml")
        '' Parse the XML document only once
        'Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        'txtPatName.Items.AddRange(cbElements)
    End Sub

    Private Sub txtPatName_Validated(sender As Object, e As EventArgs) Handles txtPatName.Validated
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Jobs.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxJob.Items.AddRange(cbElements)
            SaveInXmlPatNames()
        End If

    End Sub

    Private Sub txtPatName_TextChanged(sender As Object, e As EventArgs) Handles txtPatName.TextChanged
        'TextBox5.Text = txtPatName.Text
        'TextBox8.Text = txtPatName.Text
    End Sub

    Private Sub cbxJob_Validating(sender As Object, e As CancelEventArgs) Handles cbxJob.Validating
        If txtPatName.Text = String.Empty Then
            Exit Sub
        End If
        UpdatePatient()
        SaveInXmlJobs()
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames2.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxAddress.Items.AddRange(cbElements)
    End Sub

    Private Sub cbxJob_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxJob.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\Jobs.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxJob.Items.AddRange(cbElements)
    End Sub

    Private Sub cbxHusJob_Validating(sender As Object, e As CancelEventArgs) Handles cbxHusJob.Validating
        If txtPatName.Text = String.Empty Then
            Exit Sub
        End If
        UpdatePatient()
        SaveInXmlJobs()
    End Sub

    Private Sub cbxHusJob_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxHusJob.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\Jobs.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxHusJob.Items.AddRange(cbElements)
    End Sub


    Private Sub cbxAddress_Validating(sender As Object, e As CancelEventArgs) Handles cbxAddress.Validating
        If txtPatName.Text <> String.Empty Then
            UpdatePatient()
            SaveInXmlPatNames2()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Jobs.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxHusJob.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxAddress_Click(sender As Object, e As EventArgs) Handles cbxAddress.Click
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames2.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxAddress.Items.AddRange(cbElements)
    End Sub

    Private Sub DTPicker_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPicker.Validating
        If txtPatName.Text <> String.Empty Then
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
        If txtPatName.Text <> String.Empty Then
            UpdatePatient()
        End If
    End Sub

    Private Sub txtPhone_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtPhone.Validating
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtPatName.Text <> String.Empty Then
            UpdatePatient()
            SaveInXmlMob()
        End If
    End Sub

    Private Sub txtHusband_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtHusband.Validating
        'InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> String.Empty Then
            UpdatePatient()
            SaveInXmlPatNames1()
        End If

    End Sub

    Private Sub txtHusband_Click(sender As Object, e As EventArgs) Handles txtHusband.Click
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)

    End Sub

    Private Sub PictureBox1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDoubleClick

        End
    End Sub

    Private Sub EnabledMenst()
        Label1.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        DTPickerEDD.Visible = True
        txtElapsed.Visible = True
        txtGA.Visible = True
        Label73.Text = "Weeks"
        Label74.Text = "Weeks"
    End Sub

    Private Sub DisabledMenst()
        Label1.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        DTPickerEDD.Visible = False
        txtElapsed.Visible = False
        txtGA.Visible = False
        Label73.Text = ""
        Label74.Text = ""
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
        If txtPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub txtG_TextChanged(sender As Object, e As EventArgs) Handles txtG.TextChanged
        If txtA.Text = "" And txtP.Text = "" Then
            txtG.Text = ""
        End If
    End Sub

    Private Sub txtG_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtG.Validating
        If txtPatName.Text <> "" Then
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
        If txtPatName.Text <> "" Then
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
        If txtPatName.Text <> "" Then
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
        If txtPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub chbxNVD_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles chbxNVD.Validating
        If txtPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub


    Private Sub cbxHPOC_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxHPOC.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxHPOC.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxLD_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxLD.Validating

        If txtPatName.Text <> "" Then
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

        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxLD.Items.AddRange(cbElements)
        End If
    End Sub
    Private Sub cbxLC_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxLC.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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

        If txtPatName.Text <> "" Then
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
        txtElapsed.Text = CStr(weeks) '& "  Weeks"

        ''##https://social.msdn.microsoft.com/Forums/vstudio/en-US/b2a15b26-6d51-49d5-81cf-20fef70e8316/when-datetimepicker-value-changed-this-error-occured?forum=vbgeneral
        operations.ToAgeString(DTPicker.Value)
        txtAge.Text = operations.Formatted
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim date1 As Date = DateTimePicker1.Value  ''##Equal Now
        Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        txtElapsed.Text = CStr(weeks) '& "  Weeks"

        ''##https://social.msdn.microsoft.com/Forums/vstudio/en-US/b2a15b26-6d51-49d5-81cf-20fef70e8316/when-datetimepicker-value-changed-this-error-occured?forum=vbgeneral
        operations.ToAgeString(DTPicker.Value)
        txtAge.Text = operations.Formatted
    End Sub

    Private Sub DTPickerMns_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerMns.Validating
        If txtPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub DTPickerMns_ValueChanged(sender As Object, e As EventArgs) Handles DTPickerMns.ValueChanged
        Dim date1 As Date = DTPickerMns.Value  ''##Equal Now
        Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        txtElapsed.Text = CStr(weeks) '& "  Weeks"

    End Sub

    Private Sub DTPickerEDD_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerEDD.Validating
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
        If txtPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub DTPickerLMP_ValueChanged(sender As Object, e As EventArgs) Handles DTPickerLMP.ValueChanged
        Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        Dim date2 As Date = DTPickerLMP.Value
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        txtElapsed.Text = CStr(weeks) '& "  Weeks"
        txtGA.Text = CStr(40 - weeks)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(7)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddMonths(9)
        '#For increasing 40 weeks = 280 days 
        DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)

    End Sub

    Private Sub DTPickerLMP_Validating(sender As Object, e As EventArgs) Handles DTPickerLMP.Validating
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(7)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddMonths(9)

        DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)

        If txtPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub txtGA_TextChanged(sender As Object, e As EventArgs) Handles txtGA.TextChanged
        If txtPatName.Text = String.Empty Then
            Exit Sub
        End If
        UpdateGyn()
    End Sub

    Private Sub txtGA_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtGA.Validating
        If txtPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub txtElapsed_TextChanged(sender As Object, e As EventArgs) Handles txtElapsed.TextChanged

    End Sub

    Private Sub txtElapsed_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtElapsed.Validating
        If txtPatName.Text <> "" Then
            UpdateGyn()
        End If
    End Sub

    Private Sub cbxMedH1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxMedH1.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH1.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxMedH2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxMedH2.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH2.Items.AddRange(cbElements)

        End If
    End Sub

    Private Sub cbxMedH3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxMedH3.Validating
        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxMedH3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxSurH1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxSurH1.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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

        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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

        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxSurH3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxGynH1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxGynH1.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH1.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxGynH2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxGynH2.Validating

        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH2.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxGynH3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxGynH3.Validating

        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGynH3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxDrugH1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrugH1.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH1.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxDrugH2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrugH2.Validating

        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH2.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxDrugH3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrugH3.Validating

        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrugH3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub txtVis_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtVis.Validating
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVis.Text = GetAutonumber("Gyn2", "Vis_no") And txtPatName.Text <> "" Then
            SaveGyn2()
        End If
    End Sub

    Private Sub txtVis1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtVis1.Validating
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVis1.Text = GetAutonumber("Gyn", "Vis_no") And txtPatName.Text <> "" Then
            SaveGyn()
        End If
    End Sub

    Private Sub cbxGL_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxGL.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then

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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxGL.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPuls_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPuls.Validating

        If txtPatName.Text <> "" Then

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

        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPuls.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxBP_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxBP.Validating

        If txtPatName.Text <> "" Then

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

        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxBP.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxWeight_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxWeight.Validating

        If txtPatName.Text <> "" Then

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

        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxWeight.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxBodyBuilt_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxBodyBuilt.Validating

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then

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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxBodyBuilt.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxChtH_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxChtH.Validating

        If txtPatName.Text <> "" Then

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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxChtH.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxHdNe_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxHdNe.Validating

        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxHdNe.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxExt_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxExt.Validating

        If txtPatName.Text <> "" Then
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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxExt.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxFunL_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxFunL.Validating

        If txtPatName.Text <> "" Then

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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxFunL.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxScars_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxScars.Validating

        If txtPatName.Text <> "" Then

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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxScars.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxEdema_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxEdema.Validating

        If txtPatName.Text <> "" Then

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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxEdema.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxUS_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxUS.Validating

        If txtPatName.Text <> "" Then

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

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtPatName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations2.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxUS.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub txtAmount_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtAmount.Validating
        If txtPatName.Text <> "" Then
            UpdateGyn2()
        End If
    End Sub

    Private Sub DTPickerAtt_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerAtt.Validating
        If txtPatName.Text <> "" Then
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
            .WriteElementString("Name", txtPatName.Text)
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
            .WriteElementString("Name", txtHusband.Text)
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

    '##Save Mobile and Phone 
    Sub SaveInXmlMob()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\Mob.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Names")
            .WriteElementString("Name", txtPhone.Text)
            .WriteEndElement()
            .Close()
        End With
        xmldoc.Save(Directory.GetCurrentDirectory & "\Mob.xml")

    End Sub

    Sub SaveInXmlJobs()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\Jobs.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Names")
            .WriteElementString("Name", cbxJob.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Names")
            .WriteElementString("Name", cbxHusJob.Text)
            .WriteEndElement()
            .Close()
        End With
        xmldoc.Save(Directory.GetCurrentDirectory & "\Jobs.xml")

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
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxDia.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxInter.Text)
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
        'Trace.WriteLine("RDXmlPlan STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        'Trace.WriteLine("RDXmlPlan FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub RDXmlInv()
        'Trace.WriteLine("RDXmlInv STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        'Trace.WriteLine("RDXmlInv FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub RDXmlInvRes()
        'Trace.WriteLine("RDXmlInvRes STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        'Trace.WriteLine("RDXmlInvRes FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
    ''##Removing Duplicates from Names Xml file 
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

    Sub RDXmlMob()
        Dim fileName1 As String = "Mob.xml"
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

    Sub RDXmlJobs()
        Dim fileName1 As String = "Jobs.xml"
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
        cbxSearch.Text = ""
        Button1.BackColor = Color.LightSeaGreen
        Label81.Text = "Patient's Data"
        If Button1.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button6.BackColor = Color.SeaGreen
            Button8.BackColor = Color.SeaGreen
        End If

        If txtPatName.Text <> "" Then

            TabControl1.SelectedTab = Me.TabPage2
            Label49.Text = "U/S Visits"
            Dim con As New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(AttDt)AS[Visit Date],(GL)AS[General Look],(Pls)AS[Puls],
                               (BP)AS[Blood Pressure],(Wt)AS[Weight],(BdBt)AS[Body Built],(ChtH)AS[Chest and Heart],
                               (HdNe)AS[Head and Neck],(Ext)AS[Extremities],(FunL)AS[Fundal Level],(Scrs)AS[Scars],
                               (Edm)AS[Edema],(US)AS[Ultra Sound],(Amount)AS[Amount]
                               FROM Gyn2 WHERE Patient_no = @Patient_no
                               ORDER BY Vis_no DESC", con)
            cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
            'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
            Dim da As New OleDbDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds, "Gyn2")
            DataGridView2.DataSource = ds.Tables("Gyn2").DefaultView
            'MsgBox("Patient ID = " & txtNo.Text)
            con.Close()
            Label50.Text = (DataGridView2.Rows.Count).ToString()
            'MsgBox("Patient ID = " & txtNo.Text)
        End If


    End Sub

    Private Sub btnL_Click(sender As Object, e As EventArgs) Handles btnL.Click
        cbxSearch.Text = ""
        Button1.BackColor = Color.LightSeaGreen
        Label81.Text = "Patient's Data"
        If Button1.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button6.BackColor = Color.SeaGreen
            Button8.BackColor = Color.SeaGreen
        End If

        If txtPatName.Text <> "" Then

            TabControl1.SelectedTab = Me.TabPage2
            Label48.Text = "Previous Visits"
            Dim con As New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(MnsDate)AS[Visit Date],(NVD)AS[NVD],(CS)AS[CS],
                                    (G)AS[G],(P)AS[P],(A)AS[A],(HPOC)AS[Previous Obstetric Complications],(LD)AS[LD],
                                    (LC)AS[LC],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],
                                    (ElapW)AS[Gestational age],(GAW)AS[Remaining],(MedH1)AS[Medical History1],(MedH2)AS[Medical History2],(MedH3)AS[Medical History3],
                                    (SurH1)AS[Surgical History1],(SurH2)AS[Surgical History2],(SurH3)AS[Surgical History3],(GynH1)AS[Gynecological History1],
                                    (GynH2)AS[Gynecological History2],(GynH3)AS[Gynecological History3],(DrugH1)AS[Drug History1],(DrugH2)AS[Drug History2],(DrugH3)AS[Drug History3],(Gyn)AS[Gyna]
                                    FROM Gyn WHERE Patient_no = @Patient_no ORDER BY Vis_no DESC", con)
            cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
            'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
            Dim da As New OleDbDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds, "Gyn")
            DataGridView1.DataSource = ds.Tables("Gyn").DefaultView

            con.Close()
            Label47.Text = (DataGridView1.Rows.Count).ToString()

        End If
        'Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        'Dim date2 As Date = DTPickerLMP.Value
        'Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        'If DTPickerEDD.Value = DTPickerLMP.Value Then
        '    Exit Sub
        'End If
        'txtElapsed.Text = CStr(weeks) '& "  Weeks"
        'txtGA.Text = CStr(40 - weeks)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        'TextBox1.Text = ""
        'If ListBox1.SelectedIndex > -1 Then
        '    TextBox1.Text = CType(ListBox1.SelectedItem, String)
        'End If
        'Dim date1 As Date = DateTimePicker1.Value  ''##Equal Now
        'Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        'Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        'txtElapsed.Text = CStr(weeks) '& "  Weeks"
        ''Dim sum1, sum2 As Integer
        ''sum1 = CInt(Val(txtA.Text) + Val(txtP.Text)) + 1
        ''sum2 = CInt(Val(txtA.Text) + Val(txtP.Text))
        'If DTPickerLMP.Value = DTPickerMns.Value Then
        '    DTPickerEDD.Value = DTPickerMns.Value
        '    txtElapsed.Text = "0" '& "  Weeks"
        '    txtGA.Text = "0" '& "  Weeks"

        '    'ElseIf txtG.Text = CType(sum1, String) Then
        '    '    txtElapsed.Text = weeks & "  Weeks"
        '    '    DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
        '    'ElseIf txtG.Text = CType(sum2, String) Or txtG.Text = "" Then
        '    '    DTPickerEDD.Enabled = False
        '    '    txtGA.Text = ""
        '    '    txtElapsed.Text = ""
        '    'UpdateGyn()
        'Else
        '    'Dim date1 As Date = DateTimePicker1.Value  ''##Equal Now
        '    'Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        '    'Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        '    txtElapsed.Text = CStr(weeks) '& "  Weeks"
        '    txtGA.Text = CStr(40 - weeks) '& "  Weeks"
        'End If
        'GynEnabled()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ShowGynTable()
        'conn.Open()
        'txtNo.ResetText()
        'txtVis1.ResetText()
        'txtG.ResetText()
        'txtP.ResetText()
        'txtA.ResetText()
        'chbxNVD.Checked = False
        'chbxCS.Checked = False
        'cbxHPOC.ResetText()
        'cbxLD.ResetText()
        'cbxLC.ResetText()
        'DTPickerMns.ResetText()
        'DTPickerLMP.ResetText()
        'DTPickerEDD.ResetText()
        'txtElapsed.ResetText()
        'txtGA.ResetText()
        'cbxMedH1.ResetText()
        'cbxMedH2.ResetText()
        'cbxMedH3.ResetText()
        'cbxSurH1.ResetText()
        'cbxSurH2.ResetText()
        'cbxSurH3.ResetText()
        'cbxGynH1.ResetText()
        'cbxGynH2.ResetText()
        'cbxGynH3.ResetText()
        'cbxDrugH1.ResetText()
        'cbxDrugH2.ResetText()
        'cbxDrugH3.ResetText()
        'chbxGyn.ResetText()

        'Dim str As String = "SELECT * FROM Gyn WHERE Vis_no = @Vis_no" '& TextBox1.Text & " " '& 
        'Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        'cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(TextBox1.Text))
        'dr = cmd.ExecuteReader
        'While dr.Read
        '    txtNo.Text = dr("Patient_no").ToString
        '    txtVis1.Text = dr("Vis_no").ToString
        '    txtG.Text = dr("G").ToString
        '    txtP.Text = dr("P").ToString
        '    txtA.Text = dr("A").ToString
        '    chbxNVD.Checked = CBool(dr("NVD").ToString)
        '    chbxCS.Checked = CBool(dr("CS").ToString)
        '    cbxHPOC.Text = dr("HPOC").ToString
        '    cbxLD.Text = dr("LD").ToString
        '    cbxLC.Text = dr("LC").ToString
        '    DTPickerMns.Text = dr("MNSDate").ToString
        '    DTPickerLMP.Text = dr("LMPDate").ToString
        '    DTPickerEDD.Text = dr("EDDDate").ToString
        '    txtElapsed.Text = dr("ElapW").ToString
        '    txtGA.Text = dr("GAW").ToString
        '    cbxMedH1.Text = dr("MedH1").ToString
        '    cbxMedH2.Text = dr("MedH2").ToString
        '    cbxMedH3.Text = dr("MedH3").ToString
        '    cbxSurH1.Text = dr("SurH1").ToString
        '    cbxSurH2.Text = dr("SurH2").ToString
        '    cbxSurH3.Text = dr("SurH3").ToString
        '    cbxGynH1.Text = dr("GynH1").ToString
        '    cbxGynH2.Text = dr("GynH2").ToString
        '    cbxGynH3.Text = dr("GynH3").ToString
        '    cbxDrugH1.Text = dr("DrugH1").ToString
        '    cbxDrugH2.Text = dr("DrugH2").ToString
        '    cbxDrugH3.Text = dr("DrugH3").ToString
        '    chbxGyn.Checked = CBool(dr("Gyn").ToString)
        'End While
        'conn.Close()

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

        'txtVis.ResetText()
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
        'If TextBox3.Text = "FSL_hggi1981923" Then
        '    Dim f4 As New Form4
        '    Me.Hide()
        '    f4.ShowDialog()
        'End If

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
        cbxSearch.Text = ""
        Button1.BackColor = Color.LightSeaGreen
        TabControl1.SelectedTab = Me.TabPage2
        Label81.Text = "Patient's Data"
        If Button1.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button6.BackColor = Color.SeaGreen
            Button8.BackColor = Color.SeaGreen
            Button17.BackColor = Color.SeaGreen
        End If
        'Me.ListBox3.Items.Clear()
        'btnL.Enabled = True
        'btnF.Enabled = True
        DataGridView1.DataSource = Nothing
        Label47.Text = "0"
        DataGridView2.DataSource = Nothing
        Label50.Text = "0"

        If txtVisName.Text = "" Then
            Exit Sub
        End If
        txtNo.Text = txtVisPatNo.Text
        ShowPatTable()
        'Dim connection As OleDbConnection = New OleDbConnection()
        'connection.ConnectionString = "provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923"
        'Dim command As OleDbCommand = New OleDbCommand()
        'command.Connection = connection

        'command.CommandText = "SELECT Vis_no, EDDDate FROM Gyn WHERE (EDDDate >= ?) AND (GAW <= 4) AND (GAW > -1) AND (Gyn = 0) " &  '& txtNo.Text & " " &
        '                          "ORDER BY EDDDate DESC"
        'command.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker1.Value
        ''command.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker22.Value

        'command.CommandType = CommandType.Text
        'connection.Open()

        'Dim reader As OleDbDataReader = command.ExecuteReader()
        'While reader.Read()
        '    Dim Vis As String = CStr(reader("Vis_no"))
        '    Dim EDD As String = CStr(reader("EDDDate"))
        '    Dim item As String = String.Format("{0} : {1}", Vis, EDD & vbCrLf)
        '    Me.ListBox3.Items.Add(item).ToString()
        'End While

        'reader.Close()
        'If connection.State = ConnectionState.Open Then connection.Close()

        'Dim date1 As Date = DateTimePicker1.Value  ''##Equal Now
        'Dim date2 As Date = DTPickerLMP.Value  ''##First Date in Last Period
        'Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        'txtElapsed.Text = CStr(weeks) '& "  Weeks"
        'If DTPickerLMP.Value = DTPickerMns.Value Then
        '    DTPickerEDD.Value = DTPickerMns.Value
        '    txtElapsed.Text = "0" '& "  Weeks"
        '    txtGA.Text = "0" '& "  Weeks"
        'Else
        '    txtElapsed.Text = CStr(weeks) '& "  Weeks"
        '    txtGA.Text = CStr(40 - weeks) '& "  Weeks"
        'End If


        'TabControl1.SelectedTab = Me.TabPage2
        'Label7.Text = "Expected Date Of Delivery"
        'Dim con As New OleDbConnection(cs)
        'con.Open()
        'cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],(ElapW)AS[GA],
        '                      (GAW)AS[Remaining] FROM Gyn WHERE (EDDDate > LMPDate) AND (GAW <= 4) AND (GAW > -1) AND (Gyn = 0)
        '                       ORDER BY EDDDate DESC", con)
        ''cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker1.Value
        ''cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        'Dim da As New OleDbDataAdapter(cmd)
        'Dim ds As New DataSet
        'da.Fill(ds, "Gyn")
        'DataGridView3.DataSource = ds.Tables("Gyn").DefaultView

        'con.Close()
        'Label53.Text = ((DataGridView3.Rows.Count) - 1).ToString()
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
        DataGridView5.DataSource = Nothing
        Label78.Text = "0"
        DataGridView6.DataSource = Nothing
        Label80.Text = "0"
        'FillDGV6()
        'FillDGV5()
        DataGridView3.DataSource = Nothing
        Label53.Text = "0"
        DataGridView7.DataSource = Nothing

        'Button2.ForeColor = Color.BlueViolet
        'UpdateGyn()
        'txtGA.Text = dgv.Cells().Value.ToString

        'Dim connection As OleDbConnection = New OleDbConnection()
        'connection.ConnectionString = "provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923"
        'Dim command As OleDbCommand = New OleDbCommand()
        'command.Connection = connection

        'command.CommandText = "SELECT Vis_no, EDDDate FROM Gyn WHERE (EDDDate >= ? AND ? >= EDDDate) AND (GAW <= 4) AND (GAW > -1) AND (Gyn = 0) " &  '& txtNo.Text & " " &
        '                          "ORDER BY EDDDate DESC"
        'command.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
        'command.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value

        'command.CommandType = CommandType.Text
        'connection.Open()

        'Dim reader As OleDbDataReader = command.ExecuteReader()
        'While reader.Read()
        '    Dim Vis As String = CStr(reader("Vis_no"))
        '    Dim EDD As String = CStr(reader("EDDDate"))
        '    Dim item As String = String.Format("{0} : {1}", Vis, EDD)
        '    Me.ListBox3.Items.Add(item).ToString()
        'End While
        'reader.Close()
        'If connection.State = ConnectionState.Open Then connection.Close()
        Button2.BackColor = Color.LightSeaGreen
        If Button2.BackColor = Color.LightSeaGreen Then
            Button8.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button6.BackColor = Color.SeaGreen
            Button17.BackColor = Color.SeaGreen

        End If

        TabControl1.SelectedTab = Me.TabPage4
        Label81.Text = "Expected Date Of Delivery"
        'Label7.Text = "Expected Date Of Delivery"
        'Dim con As New OleDbConnection(cs)
        'con.Open()
        'cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],(ElapW)AS[Gestational age],
        '                      (GAW)AS[Remaining] FROM Gyn WHERE (DateDiff('ww',Date(),[EDDDate]) <= 4) AND (DateDiff('ww',Date(),[EDDDate]) >= 0) AND
        '                        (EDDDate <> LMPDate) AND (Gyn = 0)
        '                       ORDER BY EDDDate DESC", con)
        ''cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
        ''cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
        'Dim da As New OleDbDataAdapter(cmd)
        'Dim ds As New DataSet
        'da.Fill(ds, "Gyn")
        'DataGridView3.DataSource = ds.Tables("Gyn").DefaultView

        'con.Close()
        'Label53.Text = (DataGridView3.Rows.Count).ToString()

        Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        Dim date2 As Date = DTPickerLMP.Value
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        If DTPickerEDD.Value = DTPickerLMP.Value Then
            Exit Sub
        End If
        txtElapsed.Text = CStr(weeks) '& "  Weeks"
        txtGA.Text = CStr(40 - weeks)
        DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)


    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        TextBox16.Text = ""

        Dim dgv As DataGridViewRow = DataGridView1.SelectedRows(0)
        'txtNo.Text = dgv.Cells(0).Value.ToString
        TextBox6.Text = dgv.Cells(0).Value.ToString
        txtVis1.Text = dgv.Cells(1).Value.ToString
        TextBox16.Text = dgv.Cells(1).Value.ToString
        DTPickerMns.Value = CDate(dgv.Cells(2).Value.ToString)
        chbxNVD.Checked = CBool(dgv.Cells(3).Value.ToString)
        chbxCS.Checked = CBool(dgv.Cells(4).Value.ToString)
        txtG.Text = dgv.Cells(5).Value.ToString
        txtP.Text = dgv.Cells(6).Value.ToString
        txtA.Text = dgv.Cells(7).Value.ToString
        'chbxNVD.Checked = CBool(dgv.Cells(5).Value.ToString)
        'chbxCS.Checked = CBool(dgv.Cells(6).Value.ToString)
        cbxHPOC.Text = dgv.Cells(8).Value.ToString
        cbxLD.Text = dgv.Cells(9).Value.ToString
        cbxLC.Text = dgv.Cells(10).Value.ToString
        'DTPickerMns.Value = CDate(dgv.Cells(10).Value.ToString)
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

        'Button6.BackColor = Color.LightSeaGreen
        'Label81.Text = "History"
        'If Button6.BackColor = Color.LightSeaGreen Then
        '    Button2.BackColor = Color.SeaGreen
        '    Button1.BackColor = Color.SeaGreen
        '    Button8.BackColor = Color.SeaGreen
        'End If

        TabControl1.SelectedTab = Me.TabPage1
        TextBox5.Text = txtPatName.Text
        TextBox6.Text = txtNo.Text
        txtVisNo.Text = TextBox16.Text

        'TextBox8.Text = txtPatName.Text
        'TextBox7.Text = txtNo.Text

        'GynEnabled()
        'ClearForDGV2()
        Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        Dim date2 As Date = DTPickerLMP.Value
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)

        If DTPickerEDD.Value = DTPickerLMP.Value Then
            txtElapsed.Text = "0"
            txtGA.Text = "0"
            Exit Sub
        End If
        txtElapsed.Text = CStr(weeks) '& "  Weeks"
        txtGA.Text = CStr(40 - weeks)
        DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)



    End Sub

    Sub ClearForDGV1()
        'Trace.WriteLine("ClearGyn STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If TextBox6.Text = TextBox7.Text Then
            Exit Sub
        End If
        txtVis1.Text = ""
        txtA.Text = ""
        txtG.Text = ""
        txtP.Text = ""
        chbxGyn.Checked = False
        chbxNVD.Checked = False
        chbxCS.Checked = False
        txtGA.Text = ""
        txtElapsed.Text = ""

        TextBox5.Text = ""
        TextBox6.Text = ""
        DTPickerMns.Value = Now
        DTPickerLMP.Value = Now
        DTPickerEDD.Value = Now

        cbxLD.Text = ""
        cbxLC.Text = ""
        cbxHPOC.Text = ""
        cbxMedH1.Text = ""
        cbxMedH2.Text = ""
        cbxMedH3.Text = ""
        cbxSurH1.Text = ""
        cbxSurH2.Text = ""
        cbxSurH3.Text = ""
        cbxGynH1.Text = ""
        cbxGynH2.Text = ""
        cbxGynH3.Text = ""
        cbxDrugH1.Text = ""
        cbxDrugH2.Text = ""
        cbxDrugH3.Text = ""
        'chbxGyn.Checked = False
        txtVis1.Text = GetAutonumber("Gyn", "Vis_no")
        GynDisabled()
    End Sub

    Sub ClearForDGV2()
        'Trace.WriteLine("ClearGyn2 STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If TextBox7.Text = TextBox6.Text Then
            Exit Sub
        End If
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
        TextBox7.Text = ""
        TextBox8.Text = ""
        DTPickerAtt.Value = Now
        txtVis.Text = GetAutonumber("Gyn2", "Vis_no")
        Gyn2Disabled()
        'Trace.WriteLine("ClearGyn2 FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

    End Sub

    Sub ShowGyn2AttDT()

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "SELECT * FROM Gyn2 WHERE AttDt=@AttDt"
                cmd.Parameters.Add("@AttDt", OleDbType.DBDate).Value = DTPickerMns.Value
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
                        DTPickerAtt.Value = CDate(dt.Rows(0).Item("AttDt").ToString)
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub txtNo_TextChanged(sender As Object, e As EventArgs) Handles txtNo.TextChanged
        If txtNo.Text = GetAutonumber("Pat", "Patient_no") Then
            txtNo.BackColor = Color.Teal
            txtNo.ForeColor = Color.White
        ElseIf txtNo.Text <> GetAutonumber("Pat", "Patient_no") Then
            txtNo.BackColor = Color.MediumTurquoise
            txtNo.ForeColor = Color.White
        End If
        DataGridView1.DataSource = Nothing
        Label47.Text = "0"
        DataGridView2.DataSource = Nothing
        Label50.Text = "0"
        DataGridView8.DataSource = Nothing
        Label82.Text = "0"
        'FillDGV1()
        'FillDGV2()
        'FillDGV8()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TabControl1.SelectedTab = Me.TabPage1
        Button6.BackColor = Color.LightSeaGreen
        Label81.Text = "History"
        If Button6.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button8.BackColor = Color.SeaGreen
            Button17.BackColor = Color.SeaGreen
        End If

    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        TextBox17.Text = ""
        Dim dgv As DataGridViewRow = DataGridView2.SelectedRows(0)
        'txtNo.Text = dgv.Cells(0).Value.ToString
        TextBox7.Text = dgv.Cells(0).Value.ToString
        txtVis.Text = dgv.Cells(1).Value.ToString
        TextBox17.Text = dgv.Cells(1).Value.ToString
        DTPickerAtt.Value = CDate(dgv.Cells(2).Value.ToString)
        cbxGL.Text = dgv.Cells(3).Value.ToString
        cbxPuls.Text = dgv.Cells(4).Value.ToString
        cbxBP.Text = dgv.Cells(5).Value.ToString
        cbxWeight.Text = dgv.Cells(6).Value.ToString
        cbxBodyBuilt.Text = dgv.Cells(7).Value.ToString
        cbxChtH.Text = dgv.Cells(8).Value.ToString
        cbxHdNe.Text = dgv.Cells(9).Value.ToString
        cbxExt.Text = dgv.Cells(10).Value.ToString
        cbxFunL.Text = dgv.Cells(11).Value.ToString
        cbxScars.Text = dgv.Cells(12).Value.ToString
        cbxEdema.Text = dgv.Cells(13).Value.ToString
        cbxUS.Text = dgv.Cells(14).Value.ToString
        txtAmount.Text = dgv.Cells(15).Value.ToString
        'cbxGL.Text = dgv.Cells(15).Value.ToString
        'DTPickerAtt.Value = CDate(dgv.Cells(15).Value.ToString)

        Button6.BackColor = Color.LightSeaGreen
        Label81.Text = "History"
        If Button6.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button8.BackColor = Color.SeaGreen
            Button17.BackColor = Color.SeaGreen

        End If
        'TextBox5.Text = txtPatName.Text
        'TextBox6.Text = txtNo.Text
        TextBox8.Text = txtPatName.Text
        'TextBox7.Text = txtNo.Text
        TabControl1.SelectedTab = Me.TabPage1
        'ClearForDGV1()
        Gyn2Enabled()


    End Sub

    '######################### Visit Screen #############################

    Sub ClearData()
        'Trace.WriteLine("ClearData STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        'cbxVisSearch.Text = ""
        txtVisName.Text = ""
        txtComplain.Text = ""
        txtSign.Text = ""
        cbxDia.Text = ""
        cbxInter.Text = ""
        txtVisAmount.Text = ""

        DTPNow()

        txtVisNo.Text = GetAutonumber("Visits", "Visit_no")
        txtVisPatNo.Text = ""
        'Trace.WriteLine("ClearData FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ClearDrug()
        'Trace.WriteLine("ClearDrug STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        'For Each i As Control In Panel6.Controls
        '    If TypeOf i Is ComboBox Then
        '        i.Text = ""
        '    End If
        'Next
        cbxDrug1.Text = ""
        cbxDrug2.Text = ""
        cbxDrug3.Text = ""
        cbxDrug4.Text = ""
        cbxDrug5.Text = ""
        cbxDrug6.Text = ""
        cbxDrug7.Text = ""
        cbxPlan1.Text = ""
        cbxPlan2.Text = ""
        cbxPlan3.Text = ""
        cbxPlan4.Text = ""
        cbxPlan5.Text = ""
        cbxPlan6.Text = ""
        cbxPlan7.Text = ""

        'cbxDrug1.ResetText()

        'Trace.WriteLine("ClearDrug FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ClearInv()
        'Trace.WriteLine("ClearInv STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        'For Each i As Control In Panel4.Controls
        '    If TypeOf i Is TextBox Then
        '        i.Text = ""
        '    End If
        'Next
        'For Each combo As Control In Panel4.Controls
        '    If TypeOf combo Is ComboBox Then
        '        combo.Text = ""
        '    End If
        'Next
        cbxInvest.Text = ""
        cbxInvest1.Text = ""
        cbxInvest2.Text = ""
        cbxInvest3.Text = ""
        cbxInvest4.Text = ""
        cbxInvest5.Text = ""
        cbxResult.Text = ""
        cbxResult1.Text = ""
        cbxResult2.Text = ""
        cbxResult3.Text = ""
        cbxResult4.Text = ""
        cbxResult5.Text = ""

        txtAtt1.Text = ""
        txtAtt2.Text = ""
        txtAtt3.Text = ""
        txtAtt4.Text = ""
        txtAtt5.Text = ""
        txtCo1.Text = ""
        txtCo2.Text = ""
        txtCo3.Text = ""
        txtCo4.Text = ""
        txtCo5.Text = ""

        'for visit screen
        'DTPNow()
        'Trace.WriteLine("ClearInv FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub DTPNow()
        'Trace.WriteLine("DTPNow STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        'For Each L As Control In Panel4.Controls
        '    If TypeOf L Is DateTimePicker Then
        '        L.Text = CType(Now, String)
        '    End If
        'Next
        lblcurTime.Text = Now.ToShortDateString
        DTPAtt.Value = Now
        DTPickerInv.Value = Now
        DTPickerInv1.Value = Now
        DTPickerInv2.Value = Now
        DTPickerInv3.Value = Now
        DTPickerInv4.Value = Now
        DTPickerInv5.Value = Now


        'Trace.WriteLine("DTPNow FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub SaveInves()
        'Trace.WriteLine("SaveInves STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisNo.Text = GetAutonumber("Inves", "Vis_no") Then
            cmd = New OleDbCommand("INSERT INTO Inves(Patient_no, Vis_no, Name, Inves_name, Inv_Date, Result, Inves1, Inves2, Inves3, Inves4, Inves5," &
                               "Date1, Date2, Date3, Date4, Date5, Result1, Result2, Result3, Result4, Result5)" &
                               "VALUES(@Patient_no, @Vis_no, @Name, @Inves_name, @Inv_Date, @Result, @Inves1, @Inves2, @Inves3, @Inves4, @Inves5," &
                               "@Date1, @Date2, @Date3, @Date4, @Date5, @Result1, @Result2, @Result3, @Result4, @Result5)", conn)

            With cmd.Parameters
                .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtVisPatNo.Text))
                .Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
                .Add("@Name", OleDbType.VarChar).Value = txtVisName.Text
                .Add("@Inves_name", OleDbType.VarChar).Value = cbxInvest.Text
                .Add("@Inv_Date", OleDbType.DBDate).Value = CDate(DTPickerInv.Value)
                .Add("@Result", OleDbType.VarChar).Value = cbxResult.Text
                .Add("@Inves1", OleDbType.VarChar).Value = cbxInvest1.Text
                .Add("@Inves2", OleDbType.VarChar).Value = cbxInvest2.Text
                .Add("@Inves3", OleDbType.VarChar).Value = cbxInvest3.Text
                .Add("@Inves4", OleDbType.VarChar).Value = cbxInvest4.Text
                .Add("@Inves5", OleDbType.VarChar).Value = cbxInvest5.Text
                .Add("@Date1", OleDbType.DBDate).Value = CDate(DTPickerInv1.Value)
                .Add("@Date2", OleDbType.DBDate).Value = CDate(DTPickerInv2.Value)
                .Add("@Date3", OleDbType.DBDate).Value = CDate(DTPickerInv3.Value)
                .Add("@Date4", OleDbType.DBDate).Value = CDate(DTPickerInv4.Value)
                .Add("@Date5", OleDbType.DBDate).Value = CDate(DTPickerInv5.Value)
                .Add("@Result1", OleDbType.VarChar).Value = cbxResult1.Text
                .Add("@Result2", OleDbType.VarChar).Value = cbxResult2.Text
                .Add("@Result3", OleDbType.VarChar).Value = cbxResult3.Text
                .Add("@Result4", OleDbType.VarChar).Value = cbxResult4.Text
                .Add("@Result5", OleDbType.VarChar).Value = cbxResult5.Text
            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If
        'Trace.WriteLine("SaveInves FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub UpdateInves()
        'Trace.WriteLine("UpdateInves STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisNo.Text <> GetAutonumber("Inves", "Vis_no") Then
            cmd = New OleDbCommand("UPDATE Inves SET Name =@Name, Inves_name =@Inves_name, Inv_Date =@Inv_Date, Result =@Result, Inves1 =@nves1, Inves2 =@Inves2, Inves3 =@Inves3, Inves4 =@Inves4, Inves5 =@Inves5," &
                               "Date1 =@Date1, Date2 =@Date2, Date3 =@Date3, Date4 =@Date4, Date5 =@Date5, Result1 =@Result1, Result2 =@Result2, Result3 =@Result3, Result4 =@Result4, Result5 =@Result5 WHERE Vis_no =@Vis_no", conn)

            With cmd.Parameters
                .Add("@Name", OleDbType.VarChar).Value = txtVisName.Text
                .Add("@Inves_name", OleDbType.VarChar).Value = cbxInvest.Text
                .Add("@Inv_Date", OleDbType.DBDate).Value = CDate(DTPickerInv.Value)
                .Add("@Result", OleDbType.VarChar).Value = cbxResult.Text
                .Add("@Inves1", OleDbType.VarChar).Value = cbxInvest1.Text
                .Add("@Inves2", OleDbType.VarChar).Value = cbxInvest2.Text
                .Add("@Inves3", OleDbType.VarChar).Value = cbxInvest3.Text
                .Add("@Inves4", OleDbType.VarChar).Value = cbxInvest4.Text
                .Add("@Inves5", OleDbType.VarChar).Value = cbxInvest5.Text
                .Add("@Date1", OleDbType.DBDate).Value = CDate(DTPickerInv1.Value)
                .Add("@Date2", OleDbType.DBDate).Value = CDate(DTPickerInv2.Value)
                .Add("@Date3", OleDbType.DBDate).Value = CDate(DTPickerInv3.Value)
                .Add("@Date4", OleDbType.DBDate).Value = CDate(DTPickerInv4.Value)
                .Add("@Date5", OleDbType.DBDate).Value = CDate(DTPickerInv5.Value)
                .Add("@Result1", OleDbType.VarChar).Value = cbxResult1.Text
                .Add("@Result2", OleDbType.VarChar).Value = cbxResult2.Text
                .Add("@Result3", OleDbType.VarChar).Value = cbxResult3.Text
                .Add("@Result4", OleDbType.VarChar).Value = cbxResult4.Text
                .Add("@Result5", OleDbType.VarChar).Value = cbxResult5.Text
                .Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If
        'Trace.WriteLine("UpdateInves FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub UpdateVisitDP()
        'Trace.WriteLine("UpdateDrugs STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisNo.Text <> GetAutonumber("VisitDP", "Visit_no") Then
            cmd = New OleDbCommand("UPDATE VisitDP SET NameDrug =@NameDrug, NamePlan =@NamePlan, NameDrug1 =@NameDrug1, NamePlan1 =@NamePlan1, NameDrug2 =@NameDrug2 , NamePlan2 =@NamePlan2, NameDrug3 =@NameDrug3, NamePlan3 =@NamePlan3, NameDrug4 =@NameDrug4, NamePlan4 =@NamePlan4," &
                               "NameDrug5 =@NameDrug5, NamePlan5 =@NamePlan5, NameDrug6 =@NameDrug6, NamePlan6 =@NamePlan6, NameDrug7 =@NameDrug7, NamePlan7 =@NamePlan7, NameDrug8 =@NameDrug8, NamePlan8 =@NamePlan8, NameDrug9 =@NameDrug9, NamePlan9 =@NamePlan9," &
                               "NameDrug10 =@NameDrug10, NamePlan10 =@NamePlan10, NameDrug11 =@NameDrug11, NamePlan11 =@NamePlan11, NameDrug12 =@NameDrug12, NamePlan12 =@NamePlan12 where Visit_no =@Visit_no", conn)

            With cmd.Parameters
                .Add("@NameDrug", OleDbType.VarChar).Value = cbxDrug1.Text
                .Add("@NamePlan", OleDbType.VarChar).Value = cbxPlan1.Text
                .Add("@NameDrug1", OleDbType.VarChar).Value = cbxDrug2.Text
                .Add("@NamePlan1", OleDbType.VarChar).Value = cbxPlan2.Text
                .Add("@NameDrug2", OleDbType.VarChar).Value = cbxDrug3.Text
                .Add("@NamePlan2", OleDbType.VarChar).Value = cbxPlan3.Text
                .Add("@NameDrug3", OleDbType.VarChar).Value = cbxDrug4.Text
                .Add("@NamePlan3", OleDbType.VarChar).Value = cbxPlan4.Text
                .Add("@NameDrug4", OleDbType.VarChar).Value = cbxDrug5.Text
                .Add("@NamePlan4", OleDbType.VarChar).Value = cbxPlan5.Text
                .Add("@NameDrug5", OleDbType.VarChar).Value = cbxDrug6.Text
                .Add("@NamePlan5", OleDbType.VarChar).Value = cbxPlan6.Text
                .Add("@NameDrug6", OleDbType.VarChar).Value = cbxDrug7.Text
                .Add("@NamePlan6", OleDbType.VarChar).Value = cbxPlan7.Text
                .Add("@NameDrug7", OleDbType.VarChar).Value = cbxDrug8.Text
                .Add("@NamePlan7", OleDbType.VarChar).Value = cbxPlan8.Text
                .Add("@NameDrug8", OleDbType.VarChar).Value = cbxDrug9.Text
                .Add("@NamePlan8", OleDbType.VarChar).Value = cbxPlan9.Text
                .Add("@NameDrug9", OleDbType.VarChar).Value = cbxDrug10.Text
                .Add("@NamePlan9", OleDbType.VarChar).Value = cbxPlan10.Text
                .Add("@NameDrug10", OleDbType.VarChar).Value = txtDrug10.Text
                .Add("@NamePlan10", OleDbType.VarChar).Value = txtPlan10.Text
                .Add("@NameDrug11", OleDbType.VarChar).Value = txtDrug11.Text
                .Add("@NamePlan11", OleDbType.VarChar).Value = txtPlan11.Text
                .Add("@NameDrug12", OleDbType.VarChar).Value = txtDrug12.Text
                .Add("@NamePlan12", OleDbType.VarChar).Value = txtPlan12.Text
                .Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If
        'Trace.WriteLine("UpdateDrugs FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub SaveVisitDP()
        'Trace.WriteLine("SaveDrugs STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisNo.Text = GetAutonumber("VisitDP", "Visit_no") Then
            cmd = New OleDbCommand("INSERT INTO VisitDP(Visit_no, Patient_no, NameDrug, NamePlan, NameDrug1, NamePlan1, NameDrug2, NamePlan2, NameDrug3, NamePlan3, NameDrug4, NamePlan4, NameDrug5, NamePlan5," &
                               "NameDrug6, NamePlan6, NameDrug7, NamePlan7, NameDrug8, NamePlan8, NameDrug9, NamePlan9, NameDrug10, NamePlan10, NameDrug11, NamePlan11, NameDrug12, NamePlan12)" &
                               "VALUES(@Visit_no, @Patient_no,@NameDrug, @NamePlan, @NameDrug1, @NamePlan1, @NameDrug2, @NamePlan2, @NameDrug3, @NamePlan3, @NameDrug4, @NamePlan4, @NameDrug5, @NamePlan5, @NameDrug6, @NamePlan6," &
                               "@NameDrug7, @NamePlan7, @NameDrug8, @NamePlan8, @NameDrug9, @NamePlan9, @NameDrug10, @NamePlan10, @NameDrug11, @NamePlan11, @NameDrug12, @NamePlan12)", conn)

            With cmd.Parameters
                .Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
                .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtVisPatNo.Text))
                .Add("@NameDrug", OleDbType.VarChar).Value = cbxDrug1.Text
                .Add("@NamePlan", OleDbType.VarChar).Value = cbxPlan1.Text
                .Add("@NameDrug1", OleDbType.VarChar).Value = cbxDrug2.Text
                .Add("@NamePlan1", OleDbType.VarChar).Value = cbxPlan2.Text
                .Add("@NameDrug2", OleDbType.VarChar).Value = cbxDrug3.Text
                .Add("@NamePlan2", OleDbType.VarChar).Value = cbxPlan3.Text
                .Add("@NameDrug3", OleDbType.VarChar).Value = cbxDrug4.Text
                .Add("@NamePlan3", OleDbType.VarChar).Value = cbxPlan4.Text
                .Add("@NameDrug4", OleDbType.VarChar).Value = cbxDrug5.Text
                .Add("@NamePlan4", OleDbType.VarChar).Value = cbxPlan5.Text
                .Add("@NameDrug5", OleDbType.VarChar).Value = cbxDrug6.Text
                .Add("@NamePlan5", OleDbType.VarChar).Value = cbxPlan6.Text
                .Add("@NameDrug6", OleDbType.VarChar).Value = cbxDrug7.Text
                .Add("@NamePlan6", OleDbType.VarChar).Value = cbxPlan7.Text
                .Add("@NameDrug7", OleDbType.VarChar).Value = cbxDrug8.Text
                .Add("@NamePlan7", OleDbType.VarChar).Value = cbxPlan8.Text
                .Add("@NameDrug8", OleDbType.VarChar).Value = cbxDrug9.Text
                .Add("@NamePlan8", OleDbType.VarChar).Value = cbxPlan9.Text
                .Add("@NameDrug9", OleDbType.VarChar).Value = cbxDrug10.Text
                .Add("@NamePlan9", OleDbType.VarChar).Value = cbxPlan10.Text
                .Add("@NameDrug10", OleDbType.VarChar).Value = txtDrug10.Text
                .Add("@NamePlan10", OleDbType.VarChar).Value = txtPlan10.Text
                .Add("@NameDrug11", OleDbType.VarChar).Value = txtDrug11.Text
                .Add("@NamePlan11", OleDbType.VarChar).Value = txtPlan11.Text
                .Add("@NameDrug12", OleDbType.VarChar).Value = txtDrug12.Text
                .Add("@NamePlan12", OleDbType.VarChar).Value = txtPlan12.Text

            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If
        'Trace.WriteLine("SaveDrugs FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub UpdateAttach()
        'Trace.WriteLine("UpdateAttach STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisNo.Text <> GetAutonumber("AtFile", "Visit_no") Then
            cmd = New OleDbCommand("UPDATE AtFile SET LB1=@LB1, Co1=@Co1, LB2=@LB2, Co2=@Co2, LB3=@LB3, Co3=@Co3," &
                                   "LB4=@LB4, Co4=@Co4, LB5=@LB5, Co5=@Co5, AttDate=@AttDate WHERE Visit_no=@Visit_no", conn)

            With cmd.Parameters
                .Add("@LB1", OleDbType.VarChar).Value = txtAtt1.Text
                .Add("@Co1", OleDbType.VarChar).Value = txtCo1.Text
                .Add("@LB2", OleDbType.VarChar).Value = txtAtt2.Text
                .Add("@Co2", OleDbType.VarChar).Value = txtCo2.Text
                .Add("@LB3", OleDbType.VarChar).Value = txtAtt3.Text
                .Add("@Co3", OleDbType.VarChar).Value = txtCo3.Text
                .Add("@LB4", OleDbType.VarChar).Value = txtAtt4.Text
                .Add("@Co4", OleDbType.VarChar).Value = txtCo4.Text
                .Add("@LB5", OleDbType.VarChar).Value = txtAtt5.Text
                .Add("@Co5", OleDbType.VarChar).Value = txtCo5.Text
                .Add("@AttDate", OleDbType.DBDate).Value = CDate(DTPAtt.Value)
                .Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If
        'Trace.WriteLine("UpdateAttach FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub SaveAttach()
        'Trace.WriteLine("SaveAttach STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisNo.Text = GetAutonumber("AtFile", "Visit_no") Then
            'RunCommand("INSERT INTO AtFile(Patient_no, Visit_no, LB1, Co1, LB2, Co2, LB3, Co3, LB4, Co4, LB5, Co5, AttDate)" &
            '           "VALUES(@Visit_no, @LB1, @Co1, @LB2, @Co2, @LB3,  @Co3," &
            '           "@LB4, @Co4, @LB5, @Co5, @AttDate))
            cmd = New OleDbCommand("INSERT INTO AtFile(Patient_no, Visit_no, LB1, Co1, LB2, Co2, LB3, Co3, LB4, Co4, LB5, Co5, AttDate)" &
                                   "VALUES(Patient_no, @Visit_no, @LB1, @Co1, @LB2, @Co2, @LB3,  @Co3," &
                                   "@LB4, @Co4, @LB5, @Co5, @AttDate)", conn)

            With cmd.Parameters
                .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtVisPatNo.Text))
                .Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
                .Add("@LB1", OleDbType.VarChar).Value = txtAtt1.Text
                .Add("@Co1", OleDbType.VarChar).Value = txtCo1.Text
                .Add("@LB2", OleDbType.VarChar).Value = txtAtt2.Text
                .Add("@Co2", OleDbType.VarChar).Value = txtCo2.Text
                .Add("@LB3", OleDbType.VarChar).Value = txtAtt3.Text
                .Add("@Co3", OleDbType.VarChar).Value = txtCo3.Text
                .Add("@LB4", OleDbType.VarChar).Value = txtAtt4.Text
                .Add("@Co4", OleDbType.VarChar).Value = txtCo4.Text
                .Add("@LB5", OleDbType.VarChar).Value = txtAtt5.Text
                .Add("@Co5", OleDbType.VarChar).Value = txtCo5.Text
                .Add("@AttDate", OleDbType.DBDate).Value = CDate(DTPAtt.Value)

            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If

        'Trace.WriteLine("SaveAttach FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    'Sub UpdateVisName()

    '    cmd = New OleDbCommand("UPDATE Pat INNER JOIN Visits ON Pat.Patient_no = Visits.Patient_no SET Visits.Name = [Pat].[Name];", conn)

    '    If conn.State = ConnectionState.Open Then
    '        conn.Close()
    '    End If
    '    conn.Open()
    '    cmd.ExecuteNonQuery()
    '    conn.Close()

    'End Sub

    'Sub UpdateInvName()
    '    cmd = New OleDbCommand("UPDATE Pat INNER JOIN Inves ON Pat.Patient_no = Inves.Patient_no SET Inves.Name = [Pat].[Name];", conn)
    '    If conn.State = ConnectionState.Open Then
    '        conn.Close()
    '    End If
    '    conn.Open()
    '    cmd.ExecuteNonQuery()
    '    conn.Close()
    'End Sub

    Sub SaveVisits()
        If txtVisNo.Text = GetAutonumber("Visits", "Visit_no") And txtVisName.Text <> "" Then

            cmd = New OleDbCommand("INSERT INTO Visits(Visit_no, Patient_no, Name, Complain, Sign, Diagnosis, Intervention, Amount, VisDate)" &
                                   "VALUES(@Visit_no, @Patient_no, @Name, @Complain, @Sign, @Diagnosis, @Intervention, @Amount, @VisDate)", conn)
            '   '" & lblcurTime.Text & "')"
            With cmd.Parameters
                .Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
                .Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtVisPatNo.Text))
                .Add("@Name", OleDbType.VarChar).Value = txtVisName.Text
                .Add("@Complain", OleDbType.VarChar).Value = txtComplain.Text
                .Add("@Sign", OleDbType.VarChar).Value = txtSign.Text
                .Add("@Diagnosis", OleDbType.VarChar).Value = cbxDia.Text
                .Add("@Intervention", OleDbType.VarChar).Value = cbxInter.Text
                .Add("@Amount", OleDbType.VarChar).Value = txtVisAmount.Text
                .Add("@VisDate", OleDbType.DBDate).Value = CDate(lblcurTime.Text)
            End With
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()

        End If
    End Sub

    Sub UpdateVisits()
        If txtVisNo.Text <> GetAutonumber("Visits", "Visit_no") And txtVisName.Text <> "" Then
            'RunCommand("update Visits set Name='" & txtVisName.Text & "', Complain='" & txtComplain.Text & "', Sign='" & txtSign.Text & "', Diagnosis='" & cbxDia.Text & "', Intervention='" & cbxInter.Text & "', Amount='" & txtAmount.Text & "', VisDate='" & lblcurTime.Text & "' where Visit_no=" & txtVisNo.Text)
            cmd = New OleDbCommand("UPDATE Visits SET Name=@Name, Complain=@Complain, Sign=@Sign, Diagnosis=@Diagnosis," &
                                   "Intervention=@Intervention, Amount=@Amount, VisDate=@VisDate WHERE Visit_no=@Visit_no", conn)

            With cmd.Parameters
                .Add("@Name", OleDbType.VarChar).Value = txtVisName.Text
                .Add("@Complain", OleDbType.VarChar).Value = txtComplain.Text
                .Add("@Sign", OleDbType.VarChar).Value = txtSign.Text
                .Add("@Diagnosis", OleDbType.VarChar).Value = cbxDia.Text
                .Add("@Intervention", OleDbType.VarChar).Value = cbxInter.Text
                .Add("@Amount", OleDbType.VarChar).Value = txtVisAmount.Text
                .Add("@VisDate", OleDbType.DBDate).Value = CDate(lblcurTime.Text)
                .Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If

    End Sub

    Sub ShowAttachPatTable()
        'Trace.WriteLine("ShowAttachPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM AtFile WHERE Patient_no=@Patient_no " '&
                '"ORDER BY Visit_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtVisPatNo.Text))

                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)

                    If dt.Rows.Count > 0 Then
                        txtVisPatNo.Text = dt.Rows(0).Item("Patient_no").ToString

                        txtVisNo.Text = dt.Rows(0).Item("Visit_no").ToString

                        txtAtt1.Text = dt.Rows(0).Item("LB1").ToString

                        txtCo1.Text = dt.Rows(0).Item("Co1").ToString

                        txtAtt2.Text = dt.Rows(0).Item("LB2").ToString

                        txtCo2.Text = dt.Rows(0).Item("Co2").ToString

                        txtAtt3.Text = dt.Rows(0).Item("LB3").ToString

                        txtCo3.Text = dt.Rows(0).Item("Co3").ToString

                        txtAtt4.Text = dt.Rows(0).Item("LB4").ToString

                        txtCo4.Text = dt.Rows(0).Item("Co4").ToString

                        txtAtt5.Text = dt.Rows(0).Item("LB5").ToString

                        txtCo5.Text = dt.Rows(0).Item("Co5").ToString

                        DTPAtt.Text = dt.Rows(0).Item("AttDate").ToString
                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowAttachPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowAttachVisTable()
        'Trace.WriteLine("ShowAttachPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM AtFile WHERE Visit_no=@Visit_no " &
                                  "ORDER BY Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))

                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)

                    If dt.Rows.Count > 0 Then
                        txtVisPatNo.Text = dt.Rows(0).Item("Patient_no").ToString

                        txtVisNo.Text = dt.Rows(0).Item("Visit_no").ToString

                        txtAtt1.Text = dt.Rows(0).Item("LB1").ToString

                        txtCo1.Text = dt.Rows(0).Item("Co1").ToString

                        txtAtt2.Text = dt.Rows(0).Item("LB2").ToString

                        txtCo2.Text = dt.Rows(0).Item("Co2").ToString

                        txtAtt3.Text = dt.Rows(0).Item("LB3").ToString

                        txtCo3.Text = dt.Rows(0).Item("Co3").ToString

                        txtAtt4.Text = dt.Rows(0).Item("LB4").ToString

                        txtCo4.Text = dt.Rows(0).Item("Co4").ToString

                        txtAtt5.Text = dt.Rows(0).Item("LB5").ToString

                        txtCo5.Text = dt.Rows(0).Item("Co5").ToString

                        DTPAtt.Text = dt.Rows(0).Item("AttDate").ToString
                    End If

                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowAttachPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowAttachTable()
        'Trace.WriteLine("ShowAttachTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM AtFile WHERE Visit_no=@Visit_no " &
                                  "ORDER BY Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txt1.Text))

                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)

                    If dt.Rows.Count > 0 Then
                        txtVisPatNo.Text = dt.Rows(0).Item("Patient_no").ToString

                        txtVisNo.Text = dt.Rows(0).Item("Visit_no").ToString

                        txtAtt1.Text = dt.Rows(0).Item("LB1").ToString

                        txtCo1.Text = dt.Rows(0).Item("Co1").ToString

                        txtAtt2.Text = dt.Rows(0).Item("LB2").ToString

                        txtCo2.Text = dt.Rows(0).Item("Co2").ToString

                        txtAtt3.Text = dt.Rows(0).Item("LB3").ToString

                        txtCo3.Text = dt.Rows(0).Item("Co3").ToString

                        txtAtt4.Text = dt.Rows(0).Item("LB4").ToString

                        txtCo4.Text = dt.Rows(0).Item("Co4").ToString

                        txtAtt5.Text = dt.Rows(0).Item("LB5").ToString

                        txtCo5.Text = dt.Rows(0).Item("Co5").ToString

                        DTPAtt.Text = dt.Rows(0).Item("AttDate").ToString
                    End If

                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowAttachTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowInvPatTable()
        'Trace.WriteLine("ShowInvPatTable STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "SELECT * FROM Inves WHERE Patient_no=@Patient_no " '&
                '"ORDER BY Vis_no"

                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtVisPatNo.Text))

                Using dt2 As New DataTable
                    dt2.Load(cmd.ExecuteReader)
                    If dt2.Rows.Count > 0 Then

                        txtVisPatNo.Text = dt2.Rows(0).Item("Patient_no").ToString
                        txtVisNo.Text = dt2.Rows(0).Item("Vis_no").ToString
                        txtVisName.Text = dt2.Rows(0).Item("Name").ToString
                        cbxInvest.Text = dt2.Rows(0).Item("Inves_name").ToString
                        DTPickerInv.Text = dt2.Rows(0).Item("Inv_Date").ToString
                        cbxResult.Text = dt2.Rows(0).Item("Result").ToString
                        cbxInvest1.Text = dt2.Rows(0).Item("Inves1").ToString
                        cbxInvest2.Text = dt2.Rows(0).Item("Inves2").ToString
                        cbxInvest3.Text = dt2.Rows(0).Item("Inves3").ToString
                        cbxInvest4.Text = dt2.Rows(0).Item("Inves4").ToString
                        cbxInvest5.Text = dt2.Rows(0).Item("Inves5").ToString
                        DTPickerInv1.Text = dt2.Rows(0).Item("Date1").ToString
                        DTPickerInv2.Text = dt2.Rows(0).Item("Date2").ToString
                        DTPickerInv3.Text = dt2.Rows(0).Item("Date3").ToString
                        DTPickerInv4.Text = dt2.Rows(0).Item("Date4").ToString
                        DTPickerInv5.Text = dt2.Rows(0).Item("Date5").ToString
                        cbxResult1.Text = dt2.Rows(0).Item("Result1").ToString
                        cbxResult2.Text = dt2.Rows(0).Item("Result2").ToString
                        cbxResult3.Text = dt2.Rows(0).Item("Result3").ToString
                        cbxResult4.Text = dt2.Rows(0).Item("Result4").ToString
                        cbxResult5.Text = dt2.Rows(0).Item("Result5").ToString

                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowInvTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowInvVisTable()
        'Trace.WriteLine("ShowInvPatTable STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "SELECT * FROM Inves WHERE Vis_no=@Vis_no " &
                    "ORDER BY Vis_no"

                cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))

                Using dt2 As New DataTable
                    dt2.Load(cmd.ExecuteReader)
                    If dt2.Rows.Count > 0 Then

                        txtVisPatNo.Text = dt2.Rows(0).Item("Patient_no").ToString
                        txtVisNo.Text = dt2.Rows(0).Item("Vis_no").ToString
                        txtVisName.Text = dt2.Rows(0).Item("Name").ToString
                        cbxInvest.Text = dt2.Rows(0).Item("Inves_name").ToString
                        DTPickerInv.Text = dt2.Rows(0).Item("Inv_Date").ToString
                        cbxResult.Text = dt2.Rows(0).Item("Result").ToString
                        cbxInvest1.Text = dt2.Rows(0).Item("Inves1").ToString
                        cbxInvest2.Text = dt2.Rows(0).Item("Inves2").ToString
                        cbxInvest3.Text = dt2.Rows(0).Item("Inves3").ToString
                        cbxInvest4.Text = dt2.Rows(0).Item("Inves4").ToString
                        cbxInvest5.Text = dt2.Rows(0).Item("Inves5").ToString
                        DTPickerInv1.Text = dt2.Rows(0).Item("Date1").ToString
                        DTPickerInv2.Text = dt2.Rows(0).Item("Date2").ToString
                        DTPickerInv3.Text = dt2.Rows(0).Item("Date3").ToString
                        DTPickerInv4.Text = dt2.Rows(0).Item("Date4").ToString
                        DTPickerInv5.Text = dt2.Rows(0).Item("Date5").ToString
                        cbxResult1.Text = dt2.Rows(0).Item("Result1").ToString
                        cbxResult2.Text = dt2.Rows(0).Item("Result2").ToString
                        cbxResult3.Text = dt2.Rows(0).Item("Result3").ToString
                        cbxResult4.Text = dt2.Rows(0).Item("Result4").ToString
                        cbxResult5.Text = dt2.Rows(0).Item("Result5").ToString

                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowInvTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowInvTable()
        'Trace.WriteLine("ShowInvTable STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "select * from Inves where Vis_no=@Vis_no " &
                    "ORDER BY Vis_no"
                cmd.Parameters.Add("@Vis_no", OleDbType.Integer).Value = CInt(Val(txt1.Text))

                Using dt2 As New DataTable
                    dt2.Load(cmd.ExecuteReader)
                    If dt2.Rows.Count > 0 Then

                        txtVisPatNo.Text = dt2.Rows(0).Item("Patient_no").ToString
                        txtVisNo.Text = dt2.Rows(0).Item("Vis_no").ToString
                        txtVisName.Text = dt2.Rows(0).Item("Name").ToString
                        cbxInvest.Text = dt2.Rows(0).Item("Inves_name").ToString
                        DTPickerInv.Text = dt2.Rows(0).Item("Inv_Date").ToString
                        cbxResult.Text = dt2.Rows(0).Item("Result").ToString
                        cbxInvest1.Text = dt2.Rows(0).Item("Inves1").ToString
                        cbxInvest2.Text = dt2.Rows(0).Item("Inves2").ToString
                        cbxInvest3.Text = dt2.Rows(0).Item("Inves3").ToString
                        cbxInvest4.Text = dt2.Rows(0).Item("Inves4").ToString
                        cbxInvest5.Text = dt2.Rows(0).Item("Inves5").ToString
                        DTPickerInv1.Text = dt2.Rows(0).Item("Date1").ToString
                        DTPickerInv2.Text = dt2.Rows(0).Item("Date2").ToString
                        DTPickerInv3.Text = dt2.Rows(0).Item("Date3").ToString
                        DTPickerInv4.Text = dt2.Rows(0).Item("Date4").ToString
                        DTPickerInv5.Text = dt2.Rows(0).Item("Date5").ToString
                        cbxResult1.Text = dt2.Rows(0).Item("Result1").ToString
                        cbxResult2.Text = dt2.Rows(0).Item("Result2").ToString
                        cbxResult3.Text = dt2.Rows(0).Item("Result3").ToString
                        cbxResult4.Text = dt2.Rows(0).Item("Result4").ToString
                        cbxResult5.Text = dt2.Rows(0).Item("Result5").ToString

                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowInvTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisDPPatTable()
        'Trace.WriteLine("ShowVisDPPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM VisitDP WHERE Patient_no=@Patient_no " '&
                '"ORDER BY Visit_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtVisPatNo.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtVisNo.Text = dt.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        cbxDrug1.Text = dt.Rows(0).Item("NameDrug").ToString
                        cbxPlan1.Text = dt.Rows(0).Item("NamePlan").ToString
                        cbxDrug2.Text = dt.Rows(0).Item("NameDrug1").ToString
                        cbxPlan2.Text = dt.Rows(0).Item("NamePlan1").ToString
                        cbxDrug3.Text = dt.Rows(0).Item("NameDrug2").ToString
                        cbxPlan3.Text = dt.Rows(0).Item("NamePlan2").ToString
                        cbxDrug4.Text = dt.Rows(0).Item("NameDrug3").ToString
                        cbxPlan4.Text = dt.Rows(0).Item("NamePlan3").ToString
                        cbxDrug5.Text = dt.Rows(0).Item("NameDrug4").ToString
                        cbxPlan5.Text = dt.Rows(0).Item("NamePlan4").ToString
                        cbxDrug6.Text = dt.Rows(0).Item("NameDrug5").ToString
                        cbxPlan6.Text = dt.Rows(0).Item("NamePlan5").ToString
                        cbxDrug7.Text = dt.Rows(0).Item("NameDrug6").ToString
                        cbxPlan7.Text = dt.Rows(0).Item("NamePlan6").ToString
                        cbxDrug8.Text = dt.Rows(0).Item("NameDrug7").ToString
                        cbxPlan8.Text = dt.Rows(0).Item("NamePlan7").ToString
                        cbxDrug9.Text = dt.Rows(0).Item("NameDrug8").ToString
                        cbxPlan9.Text = dt.Rows(0).Item("NamePlan8").ToString
                        cbxDrug10.Text = dt.Rows(0).Item("NameDrug9").ToString
                        cbxPlan10.Text = dt.Rows(0).Item("NamePlan9").ToString
                        txtDrug10.Text = dt.Rows(0).Item("NameDrug10").ToString
                        txtPlan10.Text = dt.Rows(0).Item("NamePlan10").ToString
                        txtDrug11.Text = dt.Rows(0).Item("NameDrug11").ToString
                        txtPlan11.Text = dt.Rows(0).Item("NamePlan11").ToString
                        txtDrug12.Text = dt.Rows(0).Item("NameDrug12").ToString
                        txtPlan12.Text = dt.Rows(0).Item("NamePlan12").ToString
                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowVisDPPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisDPVisTable()
        'Trace.WriteLine("ShowVisDPPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM VisitDP WHERE Visit_no=@Visit_no " &
                                  "ORDER BY Visit_no DESC"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtVisNo.Text = dt.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        cbxDrug1.Text = dt.Rows(0).Item("NameDrug").ToString
                        cbxPlan1.Text = dt.Rows(0).Item("NamePlan").ToString
                        cbxDrug2.Text = dt.Rows(0).Item("NameDrug1").ToString
                        cbxPlan2.Text = dt.Rows(0).Item("NamePlan1").ToString
                        cbxDrug3.Text = dt.Rows(0).Item("NameDrug2").ToString
                        cbxPlan3.Text = dt.Rows(0).Item("NamePlan2").ToString
                        cbxDrug4.Text = dt.Rows(0).Item("NameDrug3").ToString
                        cbxPlan4.Text = dt.Rows(0).Item("NamePlan3").ToString
                        cbxDrug5.Text = dt.Rows(0).Item("NameDrug4").ToString
                        cbxPlan5.Text = dt.Rows(0).Item("NamePlan4").ToString
                        cbxDrug6.Text = dt.Rows(0).Item("NameDrug5").ToString
                        cbxPlan6.Text = dt.Rows(0).Item("NamePlan5").ToString
                        cbxDrug7.Text = dt.Rows(0).Item("NameDrug6").ToString
                        cbxPlan7.Text = dt.Rows(0).Item("NamePlan6").ToString
                        cbxDrug8.Text = dt.Rows(0).Item("NameDrug7").ToString
                        cbxPlan8.Text = dt.Rows(0).Item("NamePlan7").ToString
                        cbxDrug9.Text = dt.Rows(0).Item("NameDrug8").ToString
                        cbxPlan9.Text = dt.Rows(0).Item("NamePlan8").ToString
                        cbxDrug10.Text = dt.Rows(0).Item("NameDrug9").ToString
                        cbxPlan10.Text = dt.Rows(0).Item("NamePlan9").ToString
                        txtDrug10.Text = dt.Rows(0).Item("NameDrug10").ToString
                        txtPlan10.Text = dt.Rows(0).Item("NamePlan10").ToString
                        txtDrug11.Text = dt.Rows(0).Item("NameDrug11").ToString
                        txtPlan11.Text = dt.Rows(0).Item("NamePlan11").ToString
                        txtDrug12.Text = dt.Rows(0).Item("NameDrug12").ToString
                        txtPlan12.Text = dt.Rows(0).Item("NamePlan12").ToString
                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowVisDPPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisDPTable()
        'Trace.WriteLine("ShowVisDPTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM VisitDP WHERE Visit_no=@Visit_no " &
                "ORDER BY Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txt1.Text))
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        txtVisNo.Text = dt.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt.Rows(0).Item("Patient_no").ToString
                        cbxDrug1.Text = dt.Rows(0).Item("NameDrug").ToString
                        cbxPlan1.Text = dt.Rows(0).Item("NamePlan").ToString
                        cbxDrug2.Text = dt.Rows(0).Item("NameDrug1").ToString
                        cbxPlan2.Text = dt.Rows(0).Item("NamePlan1").ToString
                        cbxDrug3.Text = dt.Rows(0).Item("NameDrug2").ToString
                        cbxPlan3.Text = dt.Rows(0).Item("NamePlan2").ToString
                        cbxDrug4.Text = dt.Rows(0).Item("NameDrug3").ToString
                        cbxPlan4.Text = dt.Rows(0).Item("NamePlan3").ToString
                        cbxDrug5.Text = dt.Rows(0).Item("NameDrug4").ToString
                        cbxPlan5.Text = dt.Rows(0).Item("NamePlan4").ToString
                        cbxDrug6.Text = dt.Rows(0).Item("NameDrug5").ToString
                        cbxPlan6.Text = dt.Rows(0).Item("NamePlan5").ToString
                        cbxDrug7.Text = dt.Rows(0).Item("NameDrug6").ToString
                        cbxPlan7.Text = dt.Rows(0).Item("NamePlan6").ToString
                        cbxDrug8.Text = dt.Rows(0).Item("NameDrug7").ToString
                        cbxPlan8.Text = dt.Rows(0).Item("NamePlan7").ToString
                        cbxDrug9.Text = dt.Rows(0).Item("NameDrug8").ToString
                        cbxPlan9.Text = dt.Rows(0).Item("NamePlan8").ToString
                        cbxDrug10.Text = dt.Rows(0).Item("NameDrug9").ToString
                        cbxPlan10.Text = dt.Rows(0).Item("NamePlan9").ToString
                        txtDrug10.Text = dt.Rows(0).Item("NameDrug10").ToString
                        txtPlan10.Text = dt.Rows(0).Item("NamePlan10").ToString
                        txtDrug11.Text = dt.Rows(0).Item("NameDrug11").ToString
                        txtPlan11.Text = dt.Rows(0).Item("NamePlan11").ToString
                        txtDrug12.Text = dt.Rows(0).Item("NameDrug12").ToString
                        txtPlan12.Text = dt.Rows(0).Item("NamePlan12").ToString
                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowVisDPTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    'Sub ShowVisitsPatTable()
    '    'Trace.WriteLine("ShowVisitsPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    '    Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
    '        con.Open()
    '        Using cmd As New OleDbCommand
    '            cmd.Connection = con
    '            cmd.CommandText = "SELECT * FROM Visits WHERE Patient_no=@Patient_no"
    '            '"ORDER BY Visit_no"
    '            cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
    '            Using dt1 As New DataTable
    '                dt1.Load(cmd.ExecuteReader)
    '                If dt1.Rows.Count > 0 Then
    '                    txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
    '                    txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
    '                    txtVisName.Text = dt1.Rows(0).Item("Name").ToString
    '                    txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
    '                    txtSign.Text = dt1.Rows(0).Item("Sign").ToString
    '                    cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
    '                    cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
    '                    txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
    '                    lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
    '                End If
    '            End Using
    '        End Using
    '    End Using
    '    'Trace.WriteLine("ShowVisitsPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    'End Sub
    Sub ShowVisitsPatTable()
        'Trace.WriteLine("ShowVisitsPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits WHERE Patient_no=@Patient_no"
                '"ORDER BY Visit_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtVisPatNo.Text))
                Using dt1 As New DataTable
                    dt1.Load(cmd.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
                        txtVisName.Text = dt1.Rows(0).Item("Name").ToString
                        txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
                        txtSign.Text = dt1.Rows(0).Item("Sign").ToString
                        cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
                        cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
                        txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowVisitsPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisitsVisTable()
        'Trace.WriteLine("ShowVisitsPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits WHERE Visit_no=@Visit_no " &
                                  "ORDER BY Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(cbxSearch.Text))
                Using dt1 As New DataTable
                    dt1.Load(cmd.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
                        txtVisName.Text = dt1.Rows(0).Item("Name").ToString
                        txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
                        txtSign.Text = dt1.Rows(0).Item("Sign").ToString
                        cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
                        cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
                        txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowVisitsPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisitsTable()
        'Trace.WriteLine("ShowVisitsTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits WHERE Visit_no=@Visit_no " &
                "ORDER BY Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txt1.Text))
                Using dt1 As New DataTable
                    dt1.Load(cmd.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
                        txtVisName.Text = dt1.Rows(0).Item("Name").ToString
                        txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
                        txtSign.Text = dt1.Rows(0).Item("Sign").ToString
                        cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
                        cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
                        txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowVisitsTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisits()
        'Trace.WriteLine("ShowVisitsTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits WHERE Visit_no = @Visit_no " &
                "ORDER BY Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
                Using dt1 As New DataTable
                    dt1.Load(cmd.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
                        txtVisName.Text = dt1.Rows(0).Item("Name").ToString
                        txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
                        txtSign.Text = dt1.Rows(0).Item("Sign").ToString
                        cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
                        cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
                        txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        'Trace.WriteLine("ShowVisitsTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnVisSave_Click(sender As Object, e As EventArgs) Handles btnVisSave.Click
        'Trace.WriteLine("btnVisSave_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        CheckNull("")

        SaveVisits()
        SaveVisitDP()
        SaveInves()
        SaveAttach()
        btnNewVisit.Enabled = True

        'Trace.WriteLine("btnVisSave_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnPatient_Click(sender As Object, e As EventArgs) Handles btnPatient.Click
        'Trace.WriteLine("btnPatient_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        'If txtVisPatNo.Text = "" And txtVisName.Text = "" Then
        '    GoToPatients()
        '    f1.txtNo.Text = GetAutonumber("Pat", "Patient_no")
        '    f1.txtNo.Select()
        'ElseIf txtVisNo.Text <> GetAutonumber("Visits", "Visit_no") Then
        '    GoToPatients()
        '    f1.txtNo.Select()

        'End If
        'Trace.WriteLine("btnPatient_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub GoToPatients()
        'Trace.WriteLine("GoToPatient STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        'f1 = New Form1
        'f1.Show()
        'Me.Hide()

        'f1.txtNo.Text = txtVisPatNo.Text
        'f1.cbxnnn.Text = txtVisName.Text

        'Trace.WriteLine("GoToPatient FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnVisClear_Click(sender As Object, e As EventArgs) Handles btnVisClear.Click
        'Trace.WriteLine("btnVisClear_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ListBox4.Items.Clear()
        txt1.Text = ""
        ClearData()
        ClearDrug()
        ClearInv()
        btnNewVisit.Enabled = False
        rdoVisit.Checked = True
        cbxVisSearch.Text = ""
        InvAndAttDisabled()
        DrugDisabled()
        lblcurTime.Text = Now.ToShortDateString

        'Trace.WriteLine("btnVisClear_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnNewVisit_Click(sender As Object, e As EventArgs) Handles btnNewVisit.Click
        'Trace.WriteLine("btnNewVisit_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text = "" Then
            Exit Sub
        End If
        txtComplain.Text = ""
        txtSign.Text = ""
        cbxDia.Text = ""
        cbxInter.Text = ""
        txtAmount.Text = ""
        ClearDrug()
        ClearInv()
        DTPNow()

        If txtVisNo.Text <> GetAutonumber("Visits", "Visit_no") Then

            txtVisNo.Text = GetAutonumber("Visits", "Visit_no")

        End If
        lblcurTime.Text = Now.ToShortDateString
        btnNewVisit.Enabled = False

        btnVisSave_Click(Nothing, Nothing)
        txtVisNo.Select()
        txtVisNo.BackColor = Color.LightSeaGreen
        txtVisNo.ForeColor = Color.White

        'Trace.WriteLine("btnNewVisit_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private PrescriptionStr As String
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        'Trace.WriteLine("btnPrint_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        'PrescriptionStr = "                " & Label1.Text & "                                                  " & Label7.Text & vbNewLine &
        '                      "          " & Label2.Text & "                                        " & Label8.Text & vbNewLine &
        '                      "          " & Label3.Text & "                                                  " & Label9.Text & vbCrLf &
        '                      "          " & Label4.Text & "                                             " & Label10.Text & vbNewLine &
        '                      "          " & Label5.Text & "                                                 " & Label11.Text & vbCrLf &
        '                      "          " & Label6.Text & "                                             " & Label12.Text & vbCrLf & vbCrLf &
        '                      "          " & "Name  : " & txtVisName.Text & "                    " & "ID : " & txtVisPatNo.Text & vbCrLf &
        '                      "          " & "Diagnosis : " & cbxDia.Text & "     " & lblcurTime.Text & "        " & "No.: " & txtVisNo.Text & vbCrLf &
        '                      "          " & "R/: " & vbNewLine &
        '                      "          " & cbxDrug1.Text & vbNewLine & "                                                   " & cbxPlan1.Text & vbCrLf & "        " & cbxDrug2.Text & vbCrLf & "                                                  " & cbxPlan2.Text & vbCrLf &
        '                      "          " & cbxDrug3.Text & vbCrLf & "                                                  " & cbxPlan3.Text & vbCrLf & "        " & cbxDrug4.Text & vbCrLf & "                                                  " & cbxPlan4.Text & vbCrLf &
        '                      "          " & cbxDrug5.Text & vbCrLf & "                                                  " & cbxPlan5.Text & vbCrLf & "        " & cbxDrug6.Text & vbCrLf & "                                                  " & cbxPlan6.Text & vbCrLf &
        '                      "          " & cbxDrug7.Text & vbCrLf & "                                                  " & cbxPlan7.Text & vbCrLf & "        " & cbxDrug8.Text & vbCrLf & "                                                  " & cbxPlan8.Text & vbCrLf &
        '                      "          " & cbxDrug9.Text & vbCrLf & "                                                  " & cbxPlan9.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf &
        '                      "               " & Label13.Text

        PrescriptionStr = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine &
                            "          " & "Name  : " & txtVisName.Text & vbCrLf &
                            "          " & lblcurTime.Text & "          " & "Visit No.: " & txtVisNo.Text & "          " & "Patient ID : " & txtVisPatNo.Text & vbCrLf &        '"    " & "R/: " & vbNewLine &
                            "          " & "Diagnosis : " & cbxDia.Text & vbCrLf & vbCrLf & vbCrLf &
                            "                    " & cbxDrug1.Text & vbNewLine & "                                                   " & cbxPlan1.Text & vbCrLf & vbCrLf & vbCrLf &
                            "                    " & cbxDrug2.Text & vbCrLf & "                                                  " & cbxPlan2.Text & vbCrLf & vbCrLf &
                            "                    " & cbxDrug3.Text & vbCrLf & "                                                  " & cbxPlan3.Text & vbCrLf & vbCrLf &
                            "                    " & cbxDrug4.Text & vbCrLf & "                                                  " & cbxPlan4.Text & vbCrLf & vbCrLf &
                            "                    " & cbxDrug5.Text & vbCrLf & "                                                  " & cbxPlan5.Text & vbCrLf & vbCrLf &
                            "                    " & cbxDrug6.Text & vbCrLf & "                                                  " & cbxPlan6.Text & vbCrLf & vbCrLf &
                            "                    " & cbxDrug7.Text & vbCrLf & "                                                  " & cbxPlan7.Text & vbCrLf & vbCrLf &
                            "                    " & cbxDrug8.Text & vbCrLf & "                                                  " & cbxPlan8.Text & vbCrLf &
                            "                    " & cbxDrug9.Text & vbCrLf & "                                                  " & cbxPlan9.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf

        If cbxDrug1.Text <> "" And txtVisName.Text <> "" Then
            PrintDocument1.Print()
        End If
        'Trace.WriteLine("btnPrint_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnPrintInv_Click(sender As Object, e As EventArgs) Handles btnPrintInv.Click
        ''##This condition for investigation's prescription with report if checkbox7 is "checked"

        PrescriptionStr = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine &
                              "               " & "Name : " & txtVisName.Text & vbCrLf & vbCrLf &
                              "               " & "Visit No. : " & txtVisNo.Text & "     " & "Patient ID : " & txtVisPatNo.Text & "          " & lblcurTime.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf &
                              "                                        " & " PLEASE FOR " & vbCrLf & vbCrLf & vbCrLf & vbCrLf &
                              "                      " & cbxInvest.Text & vbCrLf & vbCrLf & vbCrLf &
                              "                      " & cbxInvest1.Text & vbCrLf & vbCrLf & vbCrLf &
                              "                      " & cbxInvest2.Text & vbCrLf & vbCrLf & vbCrLf &
                              "                      " & cbxInvest3.Text & vbCrLf & vbCrLf & vbCrLf &
                              "                      " & cbxInvest4.Text & vbCrLf & vbCrLf & vbCrLf &
                              "                      " & cbxInvest5.Text & vbCrLf & vbCrLf & vbCrLf

        If txtVisName.Text <> "" And cbxInvest.Text <> "" Then
            PrintDocument1.Print()
        End If

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'Trace.WriteLine("PrintDocument1_PrintPage STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim numChars As Integer
        Dim numLines As Integer
        Dim stringForPage As String
        Dim strFormat As New StringFormat()
        Dim PrintFont As Font

        PrintFont = Label1.Font : PrintFont = Label7.Font : PrintFont = Label2.Font : PrintFont = Label8.Font
        PrintFont = Label3.Font : PrintFont = Label9.Font : PrintFont = Label4.Font : PrintFont = Label10.Font
        PrintFont = Label5.Font : PrintFont = Label11.Font : PrintFont = Label6.Font : PrintFont = Label12.Font
        PrintFont = txtVisName.Font : PrintFont = txtVisPatNo.Font : PrintFont = txtDiagnosis.Font
        PrintFont = lblcurTime.Font : PrintFont = txtVisNo.Font
        PrintFont = cbxDrug1.Font : PrintFont = cbxPlan1.Font : PrintFont = cbxDrug2.Font : PrintFont = cbxPlan2.Font
        PrintFont = cbxDrug3.Font : PrintFont = cbxDrug4.Font : PrintFont = cbxPlan4.Font : PrintFont = cbxDrug5.Font
        PrintFont = cbxPlan5.Font : PrintFont = cbxDrug6.Font : PrintFont = cbxPlan6.Font : PrintFont = cbxDrug7.Font
        PrintFont = cbxPlan7.Font : PrintFont = cbxDrug8.Font : PrintFont = cbxPlan8.Font : PrintFont = Label13.Font
        PrintFont = cbxPlan9.Font : PrintFont = cbxDrug10.Font : PrintFont = cbxPlan10.Font
        'PrintFont = txtDrug10.Font :PrintFont = txtDrug11.Font : PrintFont = txtPlan11.Font : PrintFont = txtDrug12.Font : PrintFont = txtPlan12.Font

        PrintFont = cbxInvest.Font : PrintFont = cbxInvest1.Font : PrintFont = cbxInvest2.Font
        PrintFont = cbxInvest3.Font : PrintFont = cbxInvest4.Font : PrintFont = cbxInvest5.Font

        Dim rectDraw As New RectangleF(e.MarginBounds.Left, e.MarginBounds.Top, e.MarginBounds.Width, e.MarginBounds.Height)
        Dim sizeMeasure As New SizeF(e.MarginBounds.Width, e.MarginBounds.Height - PrintFont.GetHeight(e.Graphics))
        strFormat.Trimming = StringTrimming.Word
        e.Graphics.MeasureString(PrescriptionStr, PrintFont, sizeMeasure, strFormat, numChars, numLines)
        stringForPage = PrescriptionStr.Substring(0, numChars)
        e.Graphics.DrawString(stringForPage, PrintFont, Brushes.Black, rectDraw, strFormat)
        If numChars < PrescriptionStr.Length Then
            PrescriptionStr = PrescriptionStr.Substring(numChars)
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
        'Trace.WriteLine("PrintDocument1_PrintPage FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub SearchVisitNo()

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits WHERE Visit_no=@Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(cbxVisSearch.Text))
                Using dt1 As New DataTable
                    dt1.Load(cmd.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
                        txtVisName.Text = dt1.Rows(0).Item("Name").ToString
                        txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
                        txtSign.Text = dt1.Rows(0).Item("Sign").ToString
                        cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
                        cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
                        txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub SearchVisName()
        'MsgBox("why")
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM [Visits]
                                   WHERE Name LIKE '%" & cbxVisSearch.Text & "%' "
                'cmd.Parameters.Add("@Name", OleDbType.Integer).Value = CInt(Val(txtSearch.Text))
                Using dt1 As New DataTable
                    dt1.Load(cmd.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
                        txtVisName.Text = dt1.Rows(0).Item("Name").ToString
                        txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
                        txtSign.Text = dt1.Rows(0).Item("Sign").ToString
                        cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
                        cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
                        txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        'MsgBox("execute")
    End Sub

    Private Sub SearchVisID()

        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits" &
                                  " WHERE Patient_no LIKE '%" & cbxVisSearch.Text & "%'" & 'LIKE '%" & txtSearch.Text & "%'
                                  " ORDER BY Visit_no"
                'cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtSearch.Text))
                Using dt1 As New DataTable
                    dt1.Load(cmd.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
                        txtVisName.Text = dt1.Rows(0).Item("Name").ToString
                        txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
                        txtSign.Text = dt1.Rows(0).Item("Sign").ToString
                        cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
                        cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
                        txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Sub SearchDiagnosis()
        'MsgBox("why it doesn't open")
        ListBox4.Items.Clear()
        Dim con As New OleDbConnection(cs)
        con.Open()
        Dim str As String = "SELECT Visit_no FROM Visits " &
            "WHERE Diagnosis = @Diagnosis " &
            "ORDER BY Visit_no"
        Dim cmd As OleDbCommand = New OleDbCommand(str, con)
        cmd.Parameters.Add("@Diagnosis", OleDbType.VarChar).Value = cbxVisSearch.Text

        Dim reader As OleDbDataReader = cmd.ExecuteReader
        While reader.Read
            Dim VisitNo As String = CStr(reader("Visit_no"))
            Dim item As String = String.Format("{0}", VisitNo)
            ListBox4.Items.Add(item).ToString()
        End While
        reader.Close()
        con.Close()

        InvAndAttEnabled()
        ShowInvVisTable()
        ShowAttachVisTable()

        'ShowAttachTable() 'txt1
        'ShowInvTable()    'txt1
    End Sub

    Private Sub SearchVisDia()
        MsgBox("why")
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits 
                                   WHERE Diagnosis = @Diagnosis    
                                   ORDER BY Visit_no"  'LIKE '%" & cbxVisSearch.Text & "%'
                cmd.Parameters.Add("@Diagnosis", OleDbType.VarChar).Value = cbxVisSearch.Text
                Using dt1 As New DataTable
                    dt1.Load(cmd.ExecuteReader)
                    If dt1.Rows.Count > 0 Then
                        txtVisNo.Text = dt1.Rows(0).Item("Visit_no").ToString
                        txtVisPatNo.Text = dt1.Rows(0).Item("Patient_no").ToString
                        txtVisName.Text = dt1.Rows(0).Item("Name").ToString
                        txtComplain.Text = dt1.Rows(0).Item("Complain").ToString
                        txtSign.Text = dt1.Rows(0).Item("Sign").ToString
                        cbxDia.Text = dt1.Rows(0).Item("Diagnosis").ToString
                        cbxInter.Text = dt1.Rows(0).Item("Intervention").ToString
                        txtVisAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        MsgBox("why")
    End Sub

    Private Sub SearchtxtVisName()
        'The Code Below is not Mine, But I modified it to work with my code. This Code below belongs to Christopher Tubig, Code from: http://goo.gl/113Jd7 (Url have been shortend for convenience) User Profile:

        'Dim con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=Dr_T.accdb;jet oledb:database password=mero1981923")
        'Dim cmd As New OleDbCommand
        'Dim da As New OleDbDataAdapter
        'Dim dt As New DataTable
        'Dim sSQL As String = String.Empty

        'Try
        '    con.Open()
        '    cmd.Connection = con
        '    cmd.CommandType = CommandType.Text
        '    sSQL = "SELECT * FROM [Visits] " &
        '           "WHERE Name ='" & Me.txtVisName.Text & "' " &
        '           "ORDER BY Visit_no"

        '    cmd.CommandText = sSQL
        '    da.SelectCommand = cmd
        '    da.Fill(dt)

        '    dgvVisit.DataSource = dt
        '    If dt.Rows.Count = 0 Then
        '        Exit Sub
        '        'MsgBox("No record found!")
        '    End If

        'Catch ex As Exception
        '    MsgBox(ErrorToString)
        'Finally
        '    con.Close()
        'End Try
    End Sub

    Sub SaveInXmlInvRes()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\InvRes.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxResult.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxResult1.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxResult2.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxResult3.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxResult4.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxResult5.Text)
            .WriteEndElement()
            .Close()
        End With
        xmldoc.Save(Directory.GetCurrentDirectory & "\InvRes.xml")
    End Sub

    'Sub SaveInXmlDiaInter()
    '    ''Writing XML content...
    '    Dim xmldoc As XmlDocument = New XmlDocument()
    '    xmldoc.Load(Directory.GetCurrentDirectory & "\DiaInter.xml")

    '    With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
    '        .WriteStartElement("Invest")
    '        .WriteElementString("Name", cbxDia.Text)
    '        .WriteEndElement()
    '        .Close()
    '    End With
    '    With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
    '        .WriteStartElement("Invest")
    '        .WriteElementString("Name", cbxInter.Text)
    '        .WriteEndElement()
    '        .Close()
    '    End With

    '    xmldoc.Save(Directory.GetCurrentDirectory & "\DiaInter.xml")
    'End Sub

    Sub SaveInXmlInv()
        'Trace.WriteLine("SaveInXmlInv STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\Investigations.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxInvest.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxInvest1.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxInvest2.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxInvest3.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxInvest4.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Invest")
            .WriteElementString("Name", cbxInvest5.Text)
            .WriteEndElement()
            .Close()
        End With


        xmldoc.Save(Directory.GetCurrentDirectory & "\Investigations.xml")
        'Trace.WriteLine("SaveInXmlInv FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub SaveInXml()
        ''## These codes from https://stackoverflow.com/questions/33237851/write-combobox-selected-item-to-xml-file?answertab=votes#tab-top
        ''Writing XML content...
        'Trace.WriteLine("SaveInXml STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\Drugs1.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug1.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug2.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug3.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug4.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug5.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug6.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug7.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug8.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug9.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Drugs")
            .WriteElementString("Drug", cbxDrug10.Text)
            .WriteEndElement()
            .Close()
        End With
        xmldoc.Save(Directory.GetCurrentDirectory & "\Drugs1.xml")
        'Trace.WriteLine("SaveInXml FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    ''##For saving a new item in the Plan Xml file in the run time
    Sub SaveInXmlPlan()
        'Trace.WriteLine("SaveInXmlPlan STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\Plans.xml")

        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan1.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan2.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan3.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan4.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan5.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan6.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan7.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan8.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan9.Text)
            .WriteEndElement()
            .Close()
        End With
        With xmldoc.SelectSingleNode("/dataroot").CreateNavigator().AppendChild()
            .WriteStartElement("Plans")
            .WriteElementString("Plan", cbxPlan10.Text)
            .WriteEndElement()
            .Close()
        End With
        xmldoc.Save(Directory.GetCurrentDirectory & "\Plans.xml")
        'Trace.WriteLine("SaveInXmlPlan FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
    ''##Fernado Soto from expert-exchange  https://www.experts-exchange.com/questions/29061270/In-which-event-of-a-combo-box-can-i-Load-XML-File.html?anchor=a42323104&notificationFollowed=198557353#a42323104
    Sub LoadXmlInv()
        'Trace.WriteLine("LoadXmlInv STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxInvest.Items.AddRange(cbElements)
        cbxInvest1.Items.AddRange(cbElements)
        cbxInvest3.Items.AddRange(cbElements)
        cbxInvest4.Items.AddRange(cbElements)
        cbxInvest5.Items.AddRange(cbElements)
        cbxResult.Items.AddRange(cbElements)
        cbxResult1.Items.AddRange(cbElements)
        cbxResult2.Items.AddRange(cbElements)
        cbxResult3.Items.AddRange(cbElements)
        cbxResult4.Items.AddRange(cbElements)
        cbxResult5.Items.AddRange(cbElements)
        'Trace.WriteLine("LoadXmlInv FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub LoadXmlPlan()
        Trace.WriteLine("LoadXmlPlan STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxPlan1.Items.AddRange(cbElements)
        cbxPlan2.Items.AddRange(cbElements)
        cbxPlan3.Items.AddRange(cbElements)
        cbxPlan4.Items.AddRange(cbElements)
        cbxPlan5.Items.AddRange(cbElements)
        cbxPlan6.Items.AddRange(cbElements)
        cbxPlan7.Items.AddRange(cbElements)
        cbxPlan8.Items.AddRange(cbElements)
        cbxPlan9.Items.AddRange(cbElements)
        cbxPlan10.Items.AddRange(cbElements)
        Trace.WriteLine("LoadXmlPlan FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub LoadXmlDrug()
        'Trace.WriteLine("LoadXmlDrug STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        '' Read the XML file from disk only once
        Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
        '' Parse the XML document only once
        Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
        '' Now fill the ComboBox's 
        cbxDrug1.Items.AddRange(cbElements)
        cbxDrug2.Items.AddRange(cbElements)
        cbxDrug3.Items.AddRange(cbElements)
        cbxDrug4.Items.AddRange(cbElements)
        cbxDrug5.Items.AddRange(cbElements)
        cbxDrug6.Items.AddRange(cbElements)
        cbxDrug7.Items.AddRange(cbElements)
        cbxDrug8.Items.AddRange(cbElements)
        cbxDrug9.Items.AddRange(cbElements)
        cbxDrug10.Items.AddRange(cbElements)
        'Trace.WriteLine("LoadXmlDrug FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub RemoveDuplicateXml()
        ''##By Fernando Soto in this site: https://www.experts-exchange.com/questions/28454848/Help-with-removing-duplicates-in-an-xml-file-using-vb-net.html
        'Trace.WriteLine("RemoveDuplicateXml STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim fileName As String = "Drugs1.xml"
        Dim xdoc As XDocument = XDocument.Load(fileName)
        ' Find the duplicate nodes in the XML document                                             
        Dim results = (From n In xdoc.Descendants("Drugs")
                       Group n By Item = n.Element("Drug").Value.ToLower() Into itemGroup = Group
                       Where itemGroup.Count > 1
                       From i In itemGroup.Skip(1)
                       Select i).ToList()
        ' Remove the duplicates from xdoc                                                           
        results.ForEach(Sub(d) d.Remove())
        ' Save the modified xdoc to the file system                                                
        xdoc.Save(fileName)
        'Trace.WriteLine("RemoveDuplicateXml FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtVisNo_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtVisNo.Validating
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisNo.Text = GetAutonumber("Visits", "Visit_no") And txtVisName.Text <> "" Then
            'btnVisSave_Click(Nothing, Nothing)
            CheckNull("")
            SaveVisits()
            SaveVisitDP()
            SaveInves()
            SaveAttach()
            btnNewVisit.Enabled = True
        End If
    End Sub

    Private Sub txtComplain_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtComplain.Validating
        'Trace.WriteLine("txtComplain_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisits()
        End If

        'Trace.WriteLine("txtComplain_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtSign_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtSign.Validating
        'Trace.WriteLine("txtSign_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisits()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            ' Now fill the ComboBox's 
            cbxDia.Items.AddRange(cbElements)
        End If

        'Trace.WriteLine("txtSign_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDia_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDia.Validating
        'Trace.WriteLine("cbxDia_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisits()
            SaveInXmlDiaInter()

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInter.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDia_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDia_Click(sender As Object, e As EventArgs) Handles cbxDia.Click
        'Trace.WriteLine("txtDia_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDia.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("txtDia_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInter_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInter.Validating
        'Trace.WriteLine("cbxInter_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisits()
            SaveInXmlDiaInter()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug1.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInter_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInter_Click(sender As Object, e As EventArgs) Handles cbxInter.Click
        'Trace.WriteLine("cbxInter_Click STRTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInter.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInter_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtVisAmount_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtVisAmount.Validating
        'Trace.WriteLine("txtAmount_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisits()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug1.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("txtAmount_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug1.Validating
        'Trace.WriteLine("cbxDrug1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ''##This condition i made it to prevent the error when moving between Drugs comboboxes by Tab or by mouse
        ''##And to avoid "try and catch" statement
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()

            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan1.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug1_Click(sender As Object, e As EventArgs) Handles cbxDrug1.Click
        'Trace.WriteLine("cbxDrug1_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0) 'for English Lang
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug1.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug1_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug2.Validating
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            'btnVisSave_Click(Nothing, Nothing)
            UpdateVisitDP()
            SaveInXml()
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan2.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug2_Click(sender As Object, e As EventArgs) Handles cbxDrug2.Click
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug2.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug3.Validating
        'Trace.WriteLine("cbxDrug3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan3.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug3_Click(sender As Object, e As EventArgs) Handles cbxDrug3.Click
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug3.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug4.Validating
        'Trace.WriteLine("cbxDrug4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then

            UpdateVisitDP()
            SaveInXml()
            '' Read the XML file from disk only once
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan4.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug4_Click(sender As Object, e As EventArgs) Handles cbxDrug4.Click
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug4.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug5.Validating
        'Trace.WriteLine("cbxDrug5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()
            '' Read the XML file from disk only once
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan5.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug5_Click(sender As Object, e As EventArgs) Handles cbxDrug5.Click
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug5.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug6_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug6.Validating
        'Trace.WriteLine("cbxDrug6_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()
            '' Read the XML file from disk only once
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan6.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug6_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug6_Click(sender As Object, e As EventArgs) Handles cbxDrug6.Click
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug6.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug7_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug7.Validating
        'Trace.WriteLine("cbxDrug7_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()
            '' Read the XML file from disk only once
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan7.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug7_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug7_Click(sender As Object, e As EventArgs) Handles cbxDrug7.Click
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug7.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug8_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug8.Validating
        'Trace.WriteLine("cbxDrug8_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()
            '' Read the XML file from disk only once
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan8.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug8_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug8_Click(sender As Object, e As EventArgs) Handles cbxDrug8.Click
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug8.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug9_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug9.Validating
        'Trace.WriteLine("cbxDrug9_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()
            '' Read the XML file from disk only once
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan9.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug9_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug9_Click(sender As Object, e As EventArgs) Handles cbxDrug9.Click
        'Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug9.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug10_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug10.Validating
        'Trace.WriteLine("cbxDrug10_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()
            '' Read the XML file from disk only once
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan10.Items.AddRange(PlElements)
        End If
        'Trace.WriteLine("cbxDrug10_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan1.Validating
        'Trace.WriteLine("cbxPlan1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug2.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan2.Validating
        'Trace.WriteLine("cbxPlan2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug3.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan3.Validating
        'Trace.WriteLine("cbxPlan3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug4.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan4.Validating
        'Trace.WriteLine("cbxPlan4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug5.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan5.Validating
        'Trace.WriteLine("cbxPlan5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug6.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan6_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan6.Validating
        ' Trace.WriteLine("cbxPlan6_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug7.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan6_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan7_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan7.Validating
        'Trace.WriteLine("cbxPlan7_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug8.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan7_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan8_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan8.Validating
        'Trace.WriteLine("cbxPlan8_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug9.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan8_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan9_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan9.Validating
        'Trace.WriteLine("cbxPlan9_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug10.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxPlan9_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan10_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan10.Validating
        'Trace.WriteLine("cbxPlan10_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
        End If
        'Trace.WriteLine("cbxPlan10_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan1_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan1.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan1.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan2_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan2.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan2.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan3_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan3.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan3.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan4_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan4.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan4.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan5_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan5.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan5.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan6_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan6.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan6.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan7_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan7.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan7.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan8_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan8.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan8.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan9_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan9.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan9.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub cbxPlan10_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan10.MouseClick
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan10.Items.AddRange(cbElements)
        End If
    End Sub

    Private Sub Panel4_Paint(sender As Object, e As PaintEventArgs) Handles Panel4.Paint
        'Trace.WriteLine("Panel4_Paint STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Panel4.AutoScroll = True
        'Trace.WriteLine("Panel4_Paint FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest.Validating
        'Trace.WriteLine("cbxInvest_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInv()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest1.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest_Click(sender As Object, e As EventArgs) Handles cbxInvest.Click
        'Trace.WriteLine("cbxInvest_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest1_Validating(sender As Object, e As EventArgs) Handles cbxInvest1.Validating
        'Trace.WriteLine("cbxInvest1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInv()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest2.Items.AddRange(cbElements)
            'cbxResult1.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest1_Click(sender As Object, e As EventArgs) Handles cbxInvest1.Click
        'Trace.WriteLine("cbxInvest1_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest1.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest1_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest2.Validating
        'Trace.WriteLine("cbxInvest2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInv()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest3.Items.AddRange(cbElements)
            'cbxResult2.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest2_Click(sender As Object, e As EventArgs) Handles cbxInvest2.Click
        'Trace.WriteLine("cbxInvest2_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest2.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest2_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest3.Validating
        'Trace.WriteLine("cbxInvest3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInv()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest4.Items.AddRange(cbElements)
            'cbxResult3.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest3_Click(sender As Object, e As EventArgs) Handles cbxInvest3.Click
        'Trace.WriteLine("cbxInvest3_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest3.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest3_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest4.Validating
        'Trace.WriteLine("cbxInvest4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInv()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest5.Items.AddRange(cbElements)
            'cbxResult4.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest4_Click(sender As Object, e As EventArgs) Handles cbxInvest4.Click
        'Trace.WriteLine("cbxInvest4_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest4.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest4_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest5.Validating
        'Trace.WriteLine("cbxInvest5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInv()
        End If
        'Trace.WriteLine("cbxInvest5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest5_Click(sender As Object, e As EventArgs) Handles cbxInvest5.Click
        'Trace.WriteLine("cbxInvest5_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest5.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxInvest5_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv.Validating
        'Trace.WriteLine("DTPickerInv_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        'Trace.WriteLine("DTPickerInv_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv1.Validating
        'Trace.WriteLine("DTPickerInv1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        'Trace.WriteLine("DTPickerInv1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv2.Validating
        'Trace.WriteLine("DTPickerInv2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        'Trace.WriteLine("DTPickerInv2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv3.Validating
        'Trace.WriteLine("DTPickerInv3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        'Trace.WriteLine("DTPickerInv3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv4.Validating
        'Trace.WriteLine("DTPickerInv4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        'Trace.WriteLine("DTPickerInv4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv5.Validating
        'Trace.WriteLine("DTPickerInv5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        'Trace.WriteLine("DTPickerInv5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult.Validating
        'Trace.WriteLine("cbxResult_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInvRes()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            'cbxInvest1.Items.AddRange(cbElements)
            cbxResult1.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult_Click(sender As Object, e As EventArgs) Handles cbxResult.Click
        'Trace.WriteLine("cbxResult_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult_MouseEnter(sender As Object, e As EventArgs) Handles cbxResult.MouseEnter
        'Trace.WriteLine("cbxResult_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult1.Validating
        'Trace.WriteLine("cbxResult1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInvRes()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            'cbxInvest2.Items.AddRange(cbElements)
            cbxResult2.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult1_Click(sender As Object, e As EventArgs) Handles cbxResult1.Click
        'Trace.WriteLine("cbxResult1_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult1.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult1_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult2.Validating
        'Trace.WriteLine("cbxResult2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInvRes()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            'cbxInvest3.Items.AddRange(cbElements)
            cbxResult3.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult2_Click(sender As Object, e As EventArgs) Handles cbxResult2.Click
        'Trace.WriteLine("cbxResult2_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult2.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult2_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult3.Validating
        'Trace.WriteLine("cbxResult3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInvRes()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            'cbxInvest4.Items.AddRange(cbElements)
            cbxResult4.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult3_Click(sender As Object, e As EventArgs) Handles cbxResult3.Click
        'Trace.WriteLine("cbxResult3_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult3.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult3_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult4.Validating
        'Trace.WriteLine("cbxResult4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInvRes()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            'cbxInvest5.Items.AddRange(cbElements)
            cbxResult5.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult4_Click(sender As Object, e As EventArgs) Handles cbxResult4.Click
        'Trace.WriteLine("cbxResult4_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult4.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult4_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult5.Validating
        'Trace.WriteLine("cbxResult5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        'Trace.WriteLine("cbxResult5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult5_Click(sender As Object, e As EventArgs) Handles cbxResult5.Click
        'Trace.WriteLine("cbxResult5_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult5.Items.AddRange(cbElements)
        End If
        'Trace.WriteLine("cbxResult5_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo1.Validating
        'Trace.WriteLine("txtCo1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        'Trace.WriteLine("txtCo1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo2.Validating
        'Trace.WriteLine("txtCo2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        'Trace.WriteLine("txtCo2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo3.Validating
        'Trace.WriteLine("txtCo3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
            'btnVisSave_Click(Nothing, Nothing)
        End If
        'Trace.WriteLine("txtCo3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo4.Validating
        'Trace.WriteLine("txtCo4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
            'btnVisSave_Click(Nothing, Nothing)
        End If
        'Trace.WriteLine("txtCo4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo5.Validating
        'Trace.WriteLine("txtCo5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
            'btnVisSave_Click(Nothing, Nothing)
        End If
        'Trace.WriteLine("txtCo5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt1.MouseDoubleClick
        'Trace.WriteLine("txtAtt1_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ''##This condition for prevent error when "txtAtt textbox" is empty
        If txtAtt1.Text <> "" Then
            Process.Start(Me.txtAtt1.Text)
        End If
        'Trace.WriteLine("txtAtt1_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt2.MouseDoubleClick
        'Trace.WriteLine("txtAtt2_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtAtt2.Text <> "" Then
            Process.Start(Me.txtAtt2.Text)
        End If
        'Trace.WriteLine("txtAtt2_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt3_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt3.MouseDoubleClick
        'Trace.WriteLine("txtAtt3_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtAtt3.Text <> "" Then
            Process.Start(Me.txtAtt3.Text)
        End If
        'Trace.WriteLine("txtAtt3_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt4_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt4.MouseDoubleClick
        'Trace.WriteLine("txtAtt4_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtAtt4.Text <> "" Then
            Process.Start(Me.txtAtt4.Text)
        End If
        'Trace.WriteLine("txtAtt4_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt5_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt5.MouseDoubleClick
        'Trace.WriteLine("txtAtt5_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtAtt5.Text <> "" Then
            Process.Start(Me.txtAtt5.Text)
        End If
        'Trace.WriteLine("txtAtt5_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPAtt_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPAtt.Validating
        'Trace.WriteLine("DTPickerAtt_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        'Trace.WriteLine("DTPickerAtt_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Dim f As New OpenFileDialog
    Private Sub btnOpen1_Click(sender As Object, e As EventArgs) Handles btnOpen1.Click
        ''##from https://answers.microsoft.com/en-us/windows/forum/windows8_1-winapps-appother/storing-and-retrieving-a-file-path-using-access/374b9f15-77c3-4348-bf75-676658c9bb6b?tm=1506765470714&rtAction=1506810163404
        'Trace.WriteLine("btnOpen1_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        With f
            .Title = "Please Select a File"
            .InitialDirectory = "C:\"
            .RestoreDirectory = True
            If f.ShowDialog() <> DialogResult.OK Then
                ' user canceled - so exit code
                Exit Sub
            End If
        End With
        ' user selected a file - put file name into text box
        Me.txtAtt1.Text = f.FileName
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        'Trace.WriteLine("btnOpen1_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnopen2_Click(sender As Object, e As EventArgs) Handles btnOpen2.Click
        'Trace.WriteLine("btnOpen2_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        With f
            .Title = "Please Select a File"
            .InitialDirectory = "C:\"
            .RestoreDirectory = True
            If f.ShowDialog() <> DialogResult.OK Then
                ' user canceled - so exit code
                Exit Sub
            End If
        End With
        ' user selected a file - put file name into text box
        Me.txtAtt2.Text = f.FileName
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        'Trace.WriteLine("btnOpen2_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnopen3_Click(sender As Object, e As EventArgs) Handles btnOpen3.Click
        'Trace.WriteLine("btnOpen3_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        With f
            .Title = "Please Select a File"
            .InitialDirectory = "C:\"
            .RestoreDirectory = True
            If f.ShowDialog() <> DialogResult.OK Then
                ' user canceled - so exit code
                Exit Sub
            End If
        End With
        ' user selected a file - put file name into text box
        Me.txtAtt3.Text = f.FileName
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        'Trace.WriteLine("btnOpen3_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnopen4_Click(sender As Object, e As EventArgs) Handles btnOpen4.Click
        'Trace.WriteLine("btnOpen4_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        With f
            .Title = "Please Select a File"
            .InitialDirectory = "C:\"
            .RestoreDirectory = True
            If f.ShowDialog() <> DialogResult.OK Then
                ' user canceled - so exit code
                Exit Sub
            End If
        End With
        ' user selected a file - put file name into text box
        Me.txtAtt4.Text = f.FileName
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        'Trace.WriteLine("btnOpen4_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnopen5_Click(sender As Object, e As EventArgs) Handles btnOpen5.Click
        'Trace.WriteLine("btnOpen5_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        With f
            .Title = "Please Select a File"
            .InitialDirectory = "C:\"
            .RestoreDirectory = True
            If f.ShowDialog() <> DialogResult.OK Then
                ' user canceled - so exit code
                Exit Sub
            End If
        End With
        ' user selected a file - put file name into text box
        Me.txtAtt5.Text = f.FileName
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        'Trace.WriteLine("btnOpen5_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint
        'Trace.WriteLine("Panel1_Paint STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Panel1.AutoScroll = True
        'Trace.WriteLine("Panel1_Paint FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnInv_Click(sender As Object, e As EventArgs) Handles btnInv.Click
        ListBox4.Items.Clear()
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then

            conn.Open()
            Dim str As String = "SELECT Visit_no FROM Visits " &
                "WHERE Patient_no =" & txtVisPatNo.Text & " " &
                "ORDER BY Visit_no"
            Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
            'cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txt1.Text))

            Dim reader As OleDbDataReader = cmd.ExecuteReader
            While reader.Read
                Dim VisitNo As String = CStr(reader("Visit_no"))
                Dim item As String = String.Format("{0}", VisitNo)
                ListBox4.Items.Add(item).ToString()
            End While
            reader.Close()
            conn.Close()
        End If
        If txtVisNo.Text = GetAutonumber("Visits", "Visit_no") Then
            Exit Sub
        End If
        InvAndAttEnabled()
        DrugEnabled()
        ShowInvVisTable()
        ShowAttachVisTable()
        ShowVisDPVisTable()
    End Sub

    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        txt1.Text = ""
        If ListBox4.SelectedIndex > -1 Then
            txt1.Text = CType(ListBox4.SelectedItem, String)
        End If
        InvAndAttEnabled()
        DrugEnabled()
    End Sub

    Private Sub txt1_TextChanged(sender As Object, e As EventArgs) Handles txt1.TextChanged

        ShowVisitsTable()
        ShowVisDPTable()
        ShowInvTable()
        ShowAttachTable()
    End Sub

    Private Sub cbxVisSearch_TextChanged(sender As Object, e As EventArgs) Handles cbxVisSearch.TextChanged
        '    txt1.Text = ""
        '    ListBox1.Items.Clear()
    End Sub

    Private Sub InvAndAttEnabled()

        cbxInvest.Enabled = True
        DTPickerInv.Enabled = True
        cbxResult.Enabled = True
        cbxInvest1.Enabled = True
        cbxInvest2.Enabled = True
        cbxInvest3.Enabled = True
        cbxInvest4.Enabled = True
        cbxInvest5.Enabled = True
        DTPickerInv1.Enabled = True
        DTPickerInv2.Enabled = True
        DTPickerInv3.Enabled = True
        DTPickerInv4.Enabled = True
        DTPickerInv5.Enabled = True
        cbxResult1.Enabled = True
        cbxResult2.Enabled = True
        cbxResult3.Enabled = True
        cbxResult4.Enabled = True
        cbxResult5.Enabled = True

        txtCo1.Enabled = True
        txtCo2.Enabled = True
        txtCo3.Enabled = True
        txtCo4.Enabled = True
        txtCo5.Enabled = True
        DTPickerAtt.Enabled = True
        btnOpen1.Enabled = True
        btnOpen2.Enabled = True
        btnOpen3.Enabled = True
        btnOpen4.Enabled = True
        btnOpen5.Enabled = True

    End Sub

    Private Sub InvAndAttDisabled()

        cbxInvest.Enabled = False
        DTPickerInv.Enabled = False
        cbxResult.Enabled = False
        cbxInvest1.Enabled = False
        cbxInvest2.Enabled = False
        cbxInvest3.Enabled = False
        cbxInvest4.Enabled = False
        cbxInvest5.Enabled = False
        DTPickerInv1.Enabled = False
        DTPickerInv2.Enabled = False
        DTPickerInv3.Enabled = False
        DTPickerInv4.Enabled = False
        DTPickerInv5.Enabled = False
        cbxResult1.Enabled = False
        cbxResult2.Enabled = False
        cbxResult3.Enabled = False
        cbxResult4.Enabled = False
        cbxResult5.Enabled = False

        txtCo1.Enabled = False
        txtCo2.Enabled = False
        txtCo3.Enabled = False
        txtCo4.Enabled = False
        txtCo5.Enabled = False
        DTPickerAtt.Enabled = False
        btnOpen1.Enabled = False
        btnOpen2.Enabled = False
        btnOpen3.Enabled = False
        btnOpen4.Enabled = False
        btnOpen5.Enabled = False

    End Sub

    Sub DrugEnabled()
        cbxDrug1.Enabled = True
        cbxDrug2.Enabled = True
        cbxDrug3.Enabled = True
        cbxDrug4.Enabled = True
        cbxDrug5.Enabled = True
        cbxDrug6.Enabled = True
        cbxDrug7.Enabled = True
        cbxDrug8.Enabled = True
        cbxPlan1.Enabled = True
        cbxPlan2.Enabled = True
        cbxPlan3.Enabled = True
        cbxPlan4.Enabled = True
        cbxPlan5.Enabled = True
        cbxPlan6.Enabled = True
        cbxPlan7.Enabled = True
        cbxPlan8.Enabled = True
    End Sub

    Sub DrugDisabled()
        cbxDrug1.Enabled = False
        cbxDrug2.Enabled = False
        cbxDrug3.Enabled = False
        cbxDrug4.Enabled = False
        cbxDrug5.Enabled = False
        cbxDrug6.Enabled = False
        cbxDrug7.Enabled = False
        cbxDrug8.Enabled = False
        cbxPlan1.Enabled = False
        cbxPlan2.Enabled = False
        cbxPlan3.Enabled = False
        cbxPlan4.Enabled = False
        cbxPlan5.Enabled = False
        cbxPlan6.Enabled = False
        cbxPlan7.Enabled = False
        cbxPlan8.Enabled = False
    End Sub

    Private Sub cbxVisSearch_Validating(sender As Object, e As CancelEventArgs) Handles cbxVisSearch.Validating
        ListBox4.Items.Clear()
        'MessageBox.Show("hi what's wrong")
        If rdoVisit.Checked Then
            SearchVisitNo()
        ElseIf rdoD.Checked Then
            ClearData()
            ClearInv()
            ClearDrug()
            SearchDiagnosis()
            'MessageBox.Show("dia works well")
        ElseIf rdoVisName.Checked Then
            SearchVisName()
        ElseIf rdoID.Checked Then
            SearchVisID()
        End If

        'ShowVisDPVisTable()
        'ShowInvVisTable()
        'ShowAttachVisTable()
        'Panel4.Visible = False
        InvAndAttDisabled()
        DrugDisabled()
        btnNewVisit.Enabled = True
    End Sub

    Private Sub DTPAtt_ValueChanged(sender As Object, e As EventArgs) Handles DTPAtt.ValueChanged
        lblcurTime.Text = DTPAtt.Text
    End Sub

    Private Sub cbxVisSearch_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxVisSearch.MouseClick
        ListBox4.Items.Clear()
        If rdoVisName.Checked Then
            InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(1)
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\PatNames.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Names>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxVisSearch.Items.AddRange(cbElements)
        ElseIf rdoD.Checked Then
            InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
            '' Read the XML file from disk only once
            Dim xDoc1 = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements1 = xDoc1.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxVisSearch.Items.AddRange(cbElements1)
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            If txtPatName.Text <> "" Then

                'TabControl1.SelectedTab = Me.TabPage2
                Label48.Text = "Previous Visits"
                Dim con As New OleDbConnection(cs)
                con.Open()
                cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(MnsDate)AS[Visit Date],(NVD)AS[NVD],(CS)AS[CS],
                                        (G)AS[G],(P)AS[P],(A)AS[A],(HPOC)AS[Previous Obstetric Complications],(LD)AS[LD],
                                        (LC)AS[LC],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],
                                        (ElapW)AS[Gestational age],(GAW)AS[Remaining],(MedH1)AS[Medical History1],(MedH2)AS[Medical History2],(MedH3)AS[Medical History3],
                                        (SurH1)AS[Surgical History1],(SurH2)AS[Surgical History2],(SurH3)AS[Surgical History3],(GynH1)AS[Gynecological History1],
                                        (GynH2)AS[Gynecological History2],(GynH3)AS[Gynecological History3],(DrugH1)AS[Drug History1],(DrugH2)AS[Drug History2],(DrugH3)AS[Drug History3],(Gyn)AS[Gyna]
                                        FROM Gyn WHERE Patient_no = @Patient_no ORDER BY Vis_no DESC", con)
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
                Dim da As New OleDbDataAdapter(cmd)
                Dim ds As New DataSet
                da.Fill(ds, "Gyn")
                DataGridView1.DataSource = ds.Tables("Gyn").DefaultView

                con.Close()
                Label47.Text = (DataGridView1.Rows.Count).ToString()

            End If
            Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
            Dim date2 As Date = DTPickerLMP.Value
            Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
            If DTPickerEDD.Value = DTPickerLMP.Value Then
                Exit Sub
            End If
            txtElapsed.Text = CStr(weeks) '& "  Weeks"
            txtGA.Text = CStr(40 - weeks)
            DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
        Catch ex As Exception
            MsgBox(ErrorToString)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If txtPatName.Text <> "" Then

            TabControl1.SelectedTab = Me.TabPage2
            Label49.Text = "U/S Visits"
            Dim con As New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(AttDt)AS[Visit Date],(GL)AS[General Look],(Pls)AS[Puls],
                               (BP)AS[Blood Pressure],(Wt)AS[Weight],(BdBt)AS[Body Built],(ChtH)AS[Chest and Heart],
                               (HdNe)AS[Head and Neck],(Ext)AS[Extremities],(FunL)AS[Fundal Level],(Scrs)AS[Scars],
                               (Edm)AS[Edema],(US)AS[Ultra Sound],(Amount)AS[Amount]
                               FROM Gyn2 WHERE Patient_no = @Patient_no
                               ORDER BY Vis_no DESC", con)
            cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
            'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
            Dim da As New OleDbDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds, "Gyn2")
            DataGridView2.DataSource = ds.Tables("Gyn2").DefaultView

            con.Close()
            Label50.Text = (DataGridView2.Rows.Count).ToString()

        End If

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ListBox4.Items.Clear()
        cbxVisSearch.Text = ""
        Button8.BackColor = Color.LightSeaGreen
        Label81.Text = "Investigation Visits"
        If Button8.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button6.BackColor = Color.SeaGreen
            Button17.BackColor = Color.SeaGreen
        End If
        'If txtPatName.Text = "" Then
        '    TabControl1.SelectedTab = Me.TabPage3
        '    Exit Sub

        'End If
        ClearData()
        txtVisName.Text = txtPatName.Text
        txtVisPatNo.Text = txtNo.Text
        ShowVisitsPatTable()

        btnNewVisit.Enabled = True
        InvAndAttDisabled()
        DrugDisabled()
        TabControl1.SelectedTab = Me.TabPage3


    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs)
        Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        Dim date2 As Date = DTPickerLMP.Value
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        If DTPickerEDD.Value = DTPickerLMP.Value Then
            Exit Sub
        End If
        txtElapsed.Text = CStr(weeks) '& "  Weeks"
        txtGA.Text = CStr(40 - weeks)
        DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs)
        Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        Dim date2 As Date = DTPickerLMP.Value
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        If DTPickerEDD.Value = DTPickerLMP.Value Then
            Exit Sub
        End If
        txtElapsed.Text = CStr(weeks) '& "  Weeks"
        txtGA.Text = CStr(40 - weeks)
        DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Label7.Text = "Expected Date Of Delivery"
        Dim con As New OleDbConnection(cs)
        con.Open()
        'DateAdd('d',280,[LMPDate]) AND (EDDDate >= ?) AND (? >= EDDDate) 
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],(ElapW)AS[Gestational age],
                              (GAW)AS[Remaining] FROM Gyn WHERE (EDDDate >= ?) AND (? >= EDDDate) AND
                                (EDDDate <> LMPDate) AND (Gyn = 0)
                               ORDER BY EDDDate DESC", con)

        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker5.Value
        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker6.Value
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn")
        DataGridView3.DataSource = ds.Tables("Gyn").DefaultView

        con.Close()
        Label53.Text = (DataGridView3.Rows.Count).ToString()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Label7.Text = "Expected Date Of Delivery"
        Dim con As New OleDbConnection(cs)
        con.Open()
        'DateAdd('d',280,[LMPDate]) AND (EDDDate >= ?) AND (? >= EDDDate) 
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],(ElapW)AS[Gestational age],
                              (GAW)AS[Remaining] FROM Gyn WHERE (DateDiff('d',Date(),[EDDDate]) <= 30) AND (DateDiff('d',Date(),[EDDDate]) >= 0) AND
                                (EDDDate <> LMPDate) AND (Gyn = 0)
                               ORDER BY EDDDate ASC", con)

        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker5.Value
        cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker6.Value
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn")
        DataGridView3.DataSource = ds.Tables("Gyn").DefaultView

        con.Close()
        Label53.Text = (DataGridView3.Rows.Count).ToString()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'If txtPatName.Text <> "" Then

        'TabControl1.SelectedTab = Me.TabPage2
        Label79.Text = "Previous Visits"
        Dim con As New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(MnsDate)AS[Visit Date],(NVD)AS[NVD],(CS)AS[CS],
                                        (G)AS[G],(P)AS[P],(A)AS[A],(HPOC)AS[Previous Obstetric Complications],(LD)AS[LD],
                                        (LC)AS[LC],(LMPDate)AS[LMPDate],(EDDDate)AS[Expected Date Of Delivery],
                                        (ElapW)AS[Gestational age],(GAW)AS[Remaining],(MedH1)AS[Medical History1],(MedH2)AS[Medical History2],(MedH3)AS[Medical History3],
                                        (SurH1)AS[Surgical History1],(SurH2)AS[Surgical History2],(SurH3)AS[Surgical History3],(GynH1)AS[Gynecological History1],
                                        (GynH2)AS[Gynecological History2],(GynH3)AS[Gynecological History3],(DrugH1)AS[Drug History1],(DrugH2)AS[Drug History2],(DrugH3)AS[Drug History3],(Gyn)AS[Gyna]
                                        FROM Gyn WHERE Patient_no = @Patient_no ORDER BY Vis_no DESC", con)
            cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
            'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
            Dim da As New OleDbDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds, "Gyn")
            DataGridView6.DataSource = ds.Tables("Gyn").DefaultView

            con.Close()
        Label80.Text = (DataGridView6.Rows.Count).ToString()

        'End If
        'Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        'Dim date2 As Date = DTPickerLMP.Value
        'Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        'If DTPickerEDD.Value = DTPickerLMP.Value Then
        '    Exit Sub
        'End If
        'txtElapsed.Text = CStr(weeks) '& "  Weeks"
        'txtGA.Text = CStr(40 - weeks)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
    End Sub

    Private Sub DataGridView3_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick

        Dim dgv As DataGridViewRow = DataGridView3.SelectedRows(0)

        txtNo.Text = dgv.Cells(0).Value.ToString
        'txtVis1.Text = dgv.Cells(1).Value.ToString
        'TextBox16.Text = dgv.Cells(1).Value.ToString
        'DTPickerLMP.Value = CDate(dgv.Cells(2).Value.ToString)
        'DTPickerEDD.Value = CDate(dgv.Cells(3).Value.ToString)
        'txtElapsed.Text = dgv.Cells(4).Value.ToString
        'txtGA.Text = dgv.Cells(5).Value.ToString
        ShowPatTable()
        DGVPatients()

        'TextBox6.Text = txtNo.Text
        'TextBox7.Text = txtNo.Text
        'TextBox5.Text = txtPatName.Text
        TextBox9.Text = txtPatName.Text

        DataGridView6.DataSource = Nothing
        Label80.Text = "0"
        DataGridView5.DataSource = Nothing
        Label78.Text = "0"
        ClearData()

    End Sub

    Sub DGVPatients()
        If txtPatName.Text <> "" Then

            'TabControl1.SelectedTab = Me.TabPage2
            'Label48.Text = "Previous Visits"
            Dim con As New OleDbConnection(cs)
            con.Open()
            cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Name) AS [Patient Name],(Address) AS [Address],
                                   (Birthdate) AS [Birth Date],(Age) AS [Age],(Phone) AS [Phone],(HusName) AS [Husband Name]
                                    FROM Pat WHERE Patient_no = @Patient_no", con)
            cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
            'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
            Dim da As New OleDbDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds, "Pat")
            DataGridView7.DataSource = ds.Tables("Pat").DefaultView

            con.Close()
            'Label47.Text = (DataGridView6.Rows.Count).ToString()
        End If
    End Sub

    Private Sub DataGridView6_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView6.RowHeaderMouseClick
        TextBox16.Text = ""
        'ClearForDGV2()
        Button6.BackColor = Color.LightSeaGreen
        Label81.Text = "History"
        If Button6.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button8.BackColor = Color.SeaGreen
            Button17.BackColor = Color.SeaGreen
        End If

        Dim dgv As DataGridViewRow = DataGridView6.SelectedRows(0)
        txtNo.Text = dgv.Cells(0).Value.ToString
        txtVis1.Text = dgv.Cells(1).Value.ToString
        TextBox16.Text = dgv.Cells(1).Value.ToString
        DTPickerMns.Value = CDate(dgv.Cells(2).Value.ToString)
        chbxNVD.Checked = CBool(dgv.Cells(3).Value.ToString)
        chbxCS.Checked = CBool(dgv.Cells(4).Value.ToString)
        txtG.Text = dgv.Cells(5).Value.ToString
        txtP.Text = dgv.Cells(6).Value.ToString
        txtA.Text = dgv.Cells(7).Value.ToString
        'chbxNVD.Checked = CBool(dgv.Cells(5).Value.ToString)
        'chbxCS.Checked = CBool(dgv.Cells(6).Value.ToString)
        cbxHPOC.Text = dgv.Cells(8).Value.ToString
        cbxLD.Text = dgv.Cells(9).Value.ToString
        cbxLC.Text = dgv.Cells(10).Value.ToString
        'DTPickerMns.Value = CDate(dgv.Cells(10).Value.ToString)
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

        ShowPatTable()

        TabControl1.SelectedTab = Me.TabPage1
        TextBox5.Text = txtPatName.Text
        TextBox6.Text = txtNo.Text
        GynEnabled()
        'ClearForDGV2()
        Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        Dim date2 As Date = DTPickerLMP.Value
        Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        If DTPickerEDD.Value = DTPickerLMP.Value Then
            txtElapsed.Text = "0"
            txtGA.Text = "0"
            Exit Sub
        End If
        txtElapsed.Text = CStr(weeks) '& "  Weeks"
        txtGA.Text = CStr(40 - weeks)
        'DTPickerEDD.Value = DTPickerLMP.Value.AddDays(280)
    End Sub

    Private Sub DataGridView5_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView5.RowHeaderMouseClick
        TextBox17.Text = ""
        Button6.BackColor = Color.LightSeaGreen
        Label81.Text = "History"
        If Button6.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button8.BackColor = Color.SeaGreen
        End If

        Dim dgv As DataGridViewRow = DataGridView5.SelectedRows(0)

        'TextBox8.Text = txtPatName.Text
        'TextBox7.Text = txtNo.Text

        'ClearForDGV1()

        'txtVis.Text = dgv.Cells(0).Value.ToString
        txtNo.Text = dgv.Cells(0).Value.ToString
        txtVis.Text = dgv.Cells(1).Value.ToString
        TextBox17.Text = dgv.Cells(1).Value.ToString
        DTPickerAtt.Value = CDate(dgv.Cells(2).Value.ToString)
        cbxGL.Text = dgv.Cells(3).Value.ToString
        cbxPuls.Text = dgv.Cells(4).Value.ToString
        cbxBP.Text = dgv.Cells(5).Value.ToString
        cbxWeight.Text = dgv.Cells(6).Value.ToString
        cbxBodyBuilt.Text = dgv.Cells(7).Value.ToString
        cbxChtH.Text = dgv.Cells(8).Value.ToString
        cbxHdNe.Text = dgv.Cells(9).Value.ToString
        cbxExt.Text = dgv.Cells(10).Value.ToString
        cbxFunL.Text = dgv.Cells(11).Value.ToString
        cbxScars.Text = dgv.Cells(12).Value.ToString
        cbxEdema.Text = dgv.Cells(13).Value.ToString
        cbxUS.Text = dgv.Cells(14).Value.ToString
        txtAmount.Text = dgv.Cells(15).Value.ToString
        'cbxGL.Text = dgv.Cells(15).Value.ToString
        'DTPickerAtt.Value = CDate(dgv.Cells(15).Value.ToString)
        TextBox8.Text = txtPatName.Text
        TextBox7.Text = txtNo.Text

        ShowPatTable()
        'ClearForDGV1()

        TabControl1.SelectedTab = Me.TabPage1
        'TextBox5.Text = txtPatName.Text
        'TextBox6.Text = txtNo.Text
        'GynEnabled()


        Gyn2Enabled()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Try
            If txtPatName.Text <> "" Then

                'TabControl1.SelectedTab = Me.TabPage2
                'Label48.Text = "Previous Visits"
                Dim con As New OleDbConnection(cs)
                con.Open()
                cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Visit_no)AS[Visit No],(Complain)AS[Complain],(Sign)AS[Sign],
                                        (Diagnosis)AS[Diagnosis],(Intervention)AS[Intervention],(Amount)AS[Amount],(VisDate)AS[Visit Date] 
                                        FROM Visits WHERE Patient_no = @Patient_no ORDER BY Visit_no DESC", con)
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
                'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
                Dim da As New OleDbDataAdapter(cmd)
                Dim ds As New DataSet
                da.Fill(ds, "Visits")
                DataGridView8.DataSource = ds.Tables("Visits").DefaultView

                con.Close()
                Label82.Text = (DataGridView8.Rows.Count).ToString()
            End If
        Catch ex As Exception
            MsgBox(ErrorToString)
        End Try
    End Sub

    Private Sub DataGridView8_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView8.RowHeaderMouseClick
        TabControl1.SelectedTab = Me.TabPage3
        Button8.BackColor = Color.LightSeaGreen
        Label81.Text = "Visits"
        If Button8.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button6.BackColor = Color.SeaGreen
        End If
        ClearDrug()
        ClearInv()
        Dim dgv As DataGridViewRow = DataGridView8.SelectedRows(0)
        txtVisPatNo.Text = dgv.Cells(0).Value.ToString
        txtVisNo.Text = dgv.Cells(1).Value.ToString

        txtVisPatNo.Text = txtNo.Text
        txtVisName.Text = txtPatName.Text
        ShowVisits()
        InvAndAttDisabled()
        DrugDisabled()

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        'ListBox5.Items.Clear()
        'conn.Open()
        'Dim str As String = "SELECT Vis_no,Amount,AttDt
        '                     FROM Gyn2 WHERE (? >= AttDt AND AttDt >= ?)"

        'Dim cmd As OleDbCommand = New OleDbCommand(str, conn)
        'cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
        'cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
        'Dim reader As OleDbDataReader = cmd.ExecuteReader
        'While reader.Read
        '    Dim VisitNo As String = CStr(reader("Vis_no"))
        '    Dim Amo As String = CStr(reader("Amount"))
        '    Dim vdate As String = CStr(reader("AttDt"))
        '    Dim item As String = String.Format("{0}): {1}) : {2}", VisitNo, Amo, vdate)
        '    ListBox5.Items.Add(item).ToString()
        'End While
        'reader.Close()
        'conn.Close()
        TextBox13.Text = "0"
        Using cn As New OleDbConnection(cs)
            cn.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = cn
                cmd.CommandText = "SELECT (Patient_no)AS[ID], (Vis_no)AS[Visit No],Amount,(AttDt)AS[Visit Date] FROM Gyn2 WHERE 
                                    (? >= AttDt AND AttDt >= ?)"
                cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
                cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
                Using da As New OleDbDataAdapter(cmd), ds As New DataSet
                    da.Fill(ds, "Gyn2")
                    DataGridView10.DataSource = ds.Tables("Gyn2").DefaultView
                    Label85.Text = DataGridView10.Rows.Count.ToString
                    'Using reader As OleDbDataReader = cmd.ExecuteReader
                    'While reader.Read
                    '    Dim VisitNo As String = CStr(reader("Vis_no"))
                    '    Dim Amo As String = CStr(reader("Amount"))
                    '    Dim vdate As String = CStr(reader("AttDt"))
                    '    Dim item As String = String.Format("{0}): {1}): {2}", VisitNo, Amo, vdate)
                    '    ListBox5.Items.Add(item).ToString()
                    'End While
                    'reader.Close()
                    'End Using
                End Using
            End Using
        End Using
        SumAmountUS()
    End Sub

    Sub SumAmountUS()
        Using con As New OleDbConnection(cs) '("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con

                cmd.CommandText = "SELECT SUM(VAL(Amount)) AS total FROM Gyn2
                                    WHERE (? >= AttDt AND AttDt >= ?)"
                'ORDER BY Vis_no ASC"
                cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
                cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        TextBox11.Text = dt.Rows(0).Item("total").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'Dim con As New OleDbConnection(cs)
        'con.Open()
        'cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Visit_no)AS[Visit No],Amount,
        '                           (VisDate)AS[Visit Date] FROM Visits WHERE (? >= VisDate AND VisDate >= ?)", con)
        'cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
        'cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
        ''cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        'Dim da As New OleDbDataAdapter(cmd)
        'Dim ds As New DataSet
        'da.Fill(ds, "Visits")
        'DataGridView9.DataSource = ds.Tables("Visits").DefaultView

        'con.Close()
        TextBox13.Text = "0"
        Using cn As New OleDbConnection(cs)
            cn.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = cn
                cmd.CommandText = "SELECT (Patient_no)AS[ID],(Visit_no)AS[Visit No],Amount,
                                   (VisDate)AS[Visit Date] FROM Visits WHERE (? >= VisDate AND VisDate >= ?)"
                cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
                cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
                Using ds As New DataSet, da As New OleDbDataAdapter(cmd)
                    'Using da As New OleDbDataAdapter(cmd)
                    da.Fill(ds, "Visits")
                    DataGridView9.DataSource = ds.Tables("Visits").DefaultView
                    Label86.Text = DataGridView9.Rows.Count
                    'End Using
                End Using
            End Using
        End Using
        SumAmountVisits()

    End Sub

    Sub SumAmountVisits()
        Using con As New OleDbConnection(cs)
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT SUM(VAL(Amount)) AS total FROM Visits 
                                   WHERE (? >= VisDate AND VisDate >= ?)"
                cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker3.Value
                cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
                Using dt As New DataTable
                    dt.Load(cmd.ExecuteReader)
                    If dt.Rows.Count > 0 Then
                        TextBox12.Text = dt.Rows(0).Item("total").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub FillDGV9Empty()
        Using cn As New OleDbConnection(cs)
            cn.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = cn
                cmd.CommandText = "SELECT (Patient_no)AS[ID],(Visit_no)AS[Visit No],Amount,
                                   (VisDate)AS[Visit Date] FROM Visits WHERE Patient_no=?"
                cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(Val(TextBox9.Text))
                'cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
                Using ds As New DataSet, da As New OleDbDataAdapter(cmd)
                    'Using da As New OleDbDataAdapter(cmd)
                    da.Fill(ds, "Visits")
                    DataGridView9.DataSource = ds.Tables("Visits").DefaultView
                    Label86.Text = DataGridView9.Rows.Count
                    'End Using
                End Using
            End Using
        End Using
    End Sub

    Private Sub FillDGV10Empty()
        Using cn As New OleDbConnection(cs)
            cn.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = cn
                cmd.CommandText = "SELECT (Patient_no)AS[ID], (Vis_no)AS[Visit No],Amount,(AttDt)AS[Visit Date] FROM Gyn2 WHERE 
                                   Patient_no = ?"
                cmd.Parameters.Add("?", OleDbType.Integer).Value = CInt(Val(TextBox9.Text))
                'cmd.Parameters.Add("?", OleDbType.DBDate).Value = DateTimePicker2.Value
                Using da As New OleDbDataAdapter(cmd), ds As New DataSet
                    da.Fill(ds, "Gyn2")
                    DataGridView10.DataSource = ds.Tables("Gyn2").DefaultView
                    Label85.Text = DataGridView10.Rows.Count.ToString
                End Using
            End Using
        End Using
    End Sub

    Private Sub DTPickerEDD_VisibleChanged(sender As Object, e As EventArgs) Handles DTPickerEDD.VisibleChanged
        UpdateGyn()
    End Sub

    Private Sub DTPickerEDD_ValueChanged(sender As Object, e As EventArgs) Handles DTPickerEDD.ValueChanged
        UpdateGyn()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Label77.Text = "US Visits"
        Dim con As New OleDbConnection(cs)
        con.Open()
        cmd = New OleDbCommand("SELECT (Patient_no)AS[ID],(Vis_no)AS[Visit No],(AttDt)AS[Visit Date],(GL)AS[General Look],(Pls)AS[Puls],
                               (BP)AS[Blood Pressure],(Wt)AS[Weight],(BdBt)AS[Body Built],(ChtH)AS[Chest and Heart],
                               (HdNe)AS[Head and Neck],(Ext)AS[Extremities],(FunL)AS[Fundal Level],(Scrs)AS[Scars],
                               (Edm)AS[Edema],(US)AS[Ultra Sound],(Amount)AS[Amount]
                               FROM Gyn2 WHERE Patient_no = @Patient_no
                               ORDER BY Vis_no DESC", con)
        cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtNo.Text))
        'cmd.Parameters.Add("@Client_no", OleDbType.Integer).Value = CInt(Val(txtID.Text))
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds, "Gyn2")
        DataGridView5.DataSource = ds.Tables("Gyn2").DefaultView

        con.Close()
        Label78.Text = (DataGridView5.Rows.Count).ToString()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        'Dim sum As Integer
        If TextBox11.Text = "" And TextBox12.Text = "" Then
            Exit Sub
        End If
        TextBox13.Text = (Val(TextBox11.Text) + Val(TextBox12.Text))
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TabControl1.SelectedTab = Me.TabPage5
        Button17.BackColor = Color.LightSeaGreen
        Label81.Text = "Income"
        If Button17.BackColor = Color.LightSeaGreen Then
            Button8.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button6.BackColor = Color.SeaGreen
            Button2.BackColor = Color.SeaGreen
        End If
        DateTimePicker2.Select()
    End Sub

    Private Sub DataGridView9_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView9.RowHeaderMouseClick
        Button8.BackColor = Color.LightSeaGreen
        If Button8.BackColor = Color.LightSeaGreen Then
            Button17.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button6.BackColor = Color.SeaGreen
            Button2.BackColor = Color.SeaGreen
        End If
        ListBox4.Items.Clear()

        Dim dgv As DataGridViewRow = DataGridView9.SelectedRows(0)
        txtVisPatNo.Text = dgv.Cells(0).Value.ToString
        txtVisNo.Text = dgv.Cells(1).Value.ToString

        txtNo.Text = txtVisPatNo.Text

        ShowPatTable()
        ShowVisits()

        InvAndAttDisabled()
        DrugDisabled()
        ClearInv()
        ClearDrug()
        TabControl1.SelectedTab = Me.TabPage3

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        ClearGyn()
        TextBox16.Text = ""
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        ClearGyn2()
        TextBox17.Text = ""
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        cbxSearch.Text = ""
        txtPatName.Text = ""
        cbxJob.Text = ""
        cbxAddress.Text = ""
        DTPicker.Value = Now
        txtAge.Text = ""
        txtPhone.Text = ""
        txtHusband.Text = ""
        cbxHusJob.Text = ""
        txtNo.Text = GetAutonumber("Pat", "Patient_no")
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        TextBox9.Text = ""
        'FillDGV3()
        DataGridView3.DataSource = Nothing
        Label53.Text = "0"
        DataGridView5.DataSource = Nothing
        Label78.Text = "0"
        DataGridView6.DataSource = Nothing
        Label80.Text = "0"
        DataGridView7.DataSource = Nothing

        'FillDGV5Empty()
        'FillDGV6Empty()
        'FillDGV7Empty()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        DataGridView9.DataSource = Nothing
        Label86.Text = "0"
        DataGridView10.DataSource = Nothing
        Label85.Text = "0"
        'FillDGV9Empty()
        'FillDGV10Empty()
    End Sub

    Private Sub DataGridView10_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView10.RowHeaderMouseClick
        TextBox17.Text = ""
        Button6.BackColor = Color.LightSeaGreen
        Label81.Text = "History"
        If Button6.BackColor = Color.LightSeaGreen Then
            Button8.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button17.BackColor = Color.SeaGreen
            Button2.BackColor = Color.SeaGreen
        End If

        Dim dgv As DataGridViewRow = DataGridView10.SelectedRows(0)
        txtNo.Text = dgv.Cells(0).Value.ToString
        txtVis.Text = dgv.Cells(1).Value.ToString
        TextBox17.Text = dgv.Cells(1).Value.ToString

        ShowPatTable()
        ShowGyn2Table()

        'TextBox6.Text = txtNo.Text
        TextBox7.Text = txtNo.Text
        'TextBox5.Text = txtPatName.Text
        TextBox8.Text = txtPatName.Text
        Gyn2Enabled()
        'ClearForDGV1()

        TabControl1.SelectedTab = Me.TabPage1

    End Sub

    Private Sub txtVisNo_TextChanged(sender As Object, e As EventArgs) Handles txtVisNo.TextChanged
        If txtVisNo.Text = GetAutonumber("Visits", "Visit_no") Then
            txtVisNo.BackColor = Color.SeaGreen
            txtVisNo.ForeColor = Color.White
        ElseIf txtVisNo.Text <> GetAutonumber("Visits", "Visit_no") Then
            txtVisNo.BackColor = Color.LightSeaGreen
            txtVisNo.ForeColor = Color.White
        End If
    End Sub

    Private Sub txtVis1_TextChanged(sender As Object, e As EventArgs) Handles txtVis1.TextChanged
        'ClearForDGV2()

        If txtVis1.Text = GetAutonumber("Gyn", "Vis_no") Then
            txtVis1.BackColor = Color.SeaGreen
            txtVis1.ForeColor = Color.White
        ElseIf txtVis1.Text <> GetAutonumber("Gyn", "Vis_no") Then
            txtVis1.BackColor = Color.LightSeaGreen
            txtVis1.ForeColor = Color.White
        End If
        'TextBox16.Text = txtVis1.Text
    End Sub

    Private Sub txtVis_TextChanged(sender As Object, e As EventArgs) Handles txtVis.TextChanged
        'TextBox17.Text = txtVis.Text
        If txtVis.Text = GetAutonumber("Gyn2", "Vis_no") Then
            txtVis.BackColor = Color.SeaGreen
            txtVis.ForeColor = Color.White
        ElseIf txtVis.Text <> GetAutonumber("Gyn2", "Vis_no") Then
            txtVis.BackColor = Color.LightSeaGreen
            txtVis.ForeColor = Color.White
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = GetAutonumber("Pat", "Patient_no") Then
            TextBox6.BackColor = Color.Teal
            TextBox6.ForeColor = Color.White
        ElseIf TextBox6.Text <> GetAutonumber("Pat", "Patient_no") Then
            TextBox6.BackColor = Color.MediumTurquoise
            TextBox6.ForeColor = Color.White
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = GetAutonumber("Pat", "Patient_no") Then
            TextBox7.BackColor = Color.Teal
            TextBox7.ForeColor = Color.White
        ElseIf TextBox7.Text <> GetAutonumber("Pat", "Patient_no") Then
            TextBox7.BackColor = Color.MediumTurquoise
            TextBox7.ForeColor = Color.White
        End If
    End Sub

    Private Sub cbxSearch_Validated(sender As Object, e As EventArgs) Handles cbxSearch.Validated
        If txtVis1.Text <> GetAutonumber("Gyn", "Vis_no") Then
            ClearGyn()
        ElseIf txtVis.Text <> GetAutonumber("Gyn2", "Vis_no") Then
            ClearGyn2()
        End If

    End Sub

    Private Sub txtVisPatNo_TextChanged(sender As Object, e As EventArgs) Handles txtVisPatNo.TextChanged
        If txtVisPatNo.Text = GetAutonumber("Pat", "Patient_no") Then
            txtVisPatNo.BackColor = Color.Teal
            txtVisPatNo.ForeColor = Color.White
        ElseIf txtVisPatNo.Text <> GetAutonumber("Pat", "Patient_no") Then
            txtVisPatNo.BackColor = Color.MediumTurquoise
            txtVisPatNo.ForeColor = Color.White
        End If
    End Sub

    Private Sub txtComplain_Click(sender As Object, e As EventArgs) Handles txtComplain.Click
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
    End Sub

    Private Sub txtSign_Click(sender As Object, e As EventArgs) Handles txtSign.Click
        InputLanguage.CurrentInputLanguage = InputLanguage.InstalledInputLanguages(0)
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        Button6.BackColor = Color.LightSeaGreen
        Label81.Text = "History"
        If Button6.BackColor = Color.LightSeaGreen Then
            Button2.BackColor = Color.SeaGreen
            Button1.BackColor = Color.SeaGreen
            Button8.BackColor = Color.SeaGreen
        End If

        GynEnabled()
        ClearForDGV2()

        'Dim date1 As Date = DateTimePicker1.Value   'Now 'DTPickerMns.Value
        'Dim date2 As Date = DTPickerLMP.Value
        'Dim weeks As Integer = CInt((date1 - date2).TotalDays / 7)
        'If DTPickerEDD.Value = DTPickerLMP.Value Then
        '    txtElapsed.Text = "0"
        '    txtGA.Text = "0"
        '    Exit Sub
        'End If
        'txtElapsed.Text = CStr(weeks) '& "  Weeks"
        'txtGA.Text = CStr(40 - weeks)
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        'Button6.BackColor = Color.LightSeaGreen
        'Label81.Text = "History"
        'If Button6.BackColor = Color.LightSeaGreen Then
        '    Button2.BackColor = Color.SeaGreen
        '    Button1.BackColor = Color.SeaGreen
        '    Button8.BackColor = Color.SeaGreen
        '    Button17.BackColor = Color.SeaGreen

        'End If

        ClearForDGV1()
        'Gyn2Enabled()
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