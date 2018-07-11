Option Strict On
Option Explicit On

Imports System.Data.OleDb
Imports System.IO
Imports System.Xml
Imports System.DateTime
Imports System.Linq
Imports System.Windows.Forms
Imports System.Management
Imports System.Threading

''' <summary>
''' Author Amr Aly
''' Date of the beginning of this Application (13/6/2017) 
''' </summary>


Public Class Form2

    Inherits System.Windows.Forms.Form

    Dim conn As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader

    Dim f2 As Form2
    Public f1 As Form1

    Public Sub New(S As String, N As String)

        Trace.WriteLine("New STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call

        'ReleaseMemory()
        Me.txtVisName.Text = S
        Me.txtVisPatNo.Text = N

        txtVisNo.Text = GetAutonumber("Visits", "Visit_no")
        DTPickerNow()
        lblcurTime.Text = Now.ToShortDateString
        btnNewVisit.Enabled = False
        txtVisNo.Select()

        Trace.WriteLine("New FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
    ''##Very important code for check null value when saving in database
    Public Function CheckNull(ByVal fieldValue As String) As String
        Trace.WriteLine("CheckNull STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If fieldValue.Equals(DBNull.Value) Then Return "" Else
        If fieldValue = "N/A" Then
            Return "value for N/A"
        Else
            Return ""
        End If
        Trace.WriteLine("CheckNull FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Function
    ''#This code for load data in run time,But we have to find another one  
    Private Sub loaddata()
        Trace.WriteLine("loaddata STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ''#This line very important when you use this code before (in the form1)
        Me.DataBindings.Clear()
        Me.Controls.Clear()
        InitializeComponent()
        Form2_Load(Nothing, Nothing)

        Me.AutoScroll = True
        Me.Location = New Point(0, 0)
        Me.Size = Screen.PrimaryScreen.WorkingArea.Size

        DTPickerNow()
        InvAndAttDisabled()
        txtVisNo.Text = GetAutonumber("Visits", "Visit_no")
        lblcurTime.Text = Now.ToShortDateString
        Trace.WriteLine("loaddata FINSHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Trace.WriteLine("frm2_Load STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Trace.WriteLine("frm2_Load FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub Form2_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Trace.WriteLine("frm2_Shown STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Me.AutoScroll = True
        Me.Location = New Point(0, 0)
        Me.Size = Screen.PrimaryScreen.WorkingArea.Size

        UpdateVisName()
        UpdateInvName()

        If txtVisName.Text <> "" Then

            ShowVisDPPatTable()
            ShowVisitsPatTable()
            btnNewVisit.Enabled = True
        End If
        txtVisNo.Select()
        'Panel4.Visible = False
        InvAndAttDisabled()
        Trace.WriteLine("frm2_Shown FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Trace.WriteLine("frm2_FormClosing STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        RDXmlDiaInter()
        RDXmlInv()
        RDXmlInvRes()
        RDXmlPlan()
        RemoveDuplicateXml()
        Form1.Show()
        Form1.Close()
        End
        Trace.WriteLine("frm2_FormClosing FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal hProcess As IntPtr, ByVal dwMinimumWorkingSetSize As Int32, ByVal dwMaximumWorkingSetSize As Int32) As Int32
    Friend Sub ReleaseMemory()
        Try
            GC.Collect()
            GC.WaitForPendingFinalizers()
            If Environment.OSVersion.Platform = PlatformID.Win32NT Then
                SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Function GetTable(SelectCommand As String) As DataTable
        Trace.WriteLine("GatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim cmd As New OleDbCommand("", conn)
        Try
            Dim Data_table3 As New DataTable
            If conn.State = ConnectionState.Closed Then conn.Open()
            cmd.CommandText = SelectCommand
            Data_table3.Load(cmd.ExecuteReader())
            Return Data_table3
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
        Dim Str1 As String
        Str1 = "select max ( " & ColumnName & " ) + 1 from " & TableName
        Dim Data_table4 As New DataTable
        Data_table4 = GetTable(Str1)
        Dim AutoNum As String
        If Data_table4.Rows(0)(0) Is DBNull.Value Then
            AutoNum = "1"
        Else
            AutoNum = CType(Data_table4.Rows(0)(0), String)
        End If
        Return AutoNum
        Trace.WriteLine("GetAutonumber FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Function

    Sub ClearData()
        Trace.WriteLine("ClearData STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        txtSearch.Text = ""
        txtVisName.Text = ""
        txtComplain.Text = ""
        txtSign.Text = ""
        cbxDia.Text = ""
        cbxInter.Text = ""
        txtAmount.Text = ""
        DTPickerNow()

        txtVisNo.Text = GetAutonumber("Visits", "Visit_no")
        txtVisPatNo.Text = ""
        Trace.WriteLine("ClearData FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ClearDrug()
        Trace.WriteLine("ClearDrug STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        For Each i As Control In Panel1.Controls
            If TypeOf i Is ComboBox Then
                i.Text = ""
            End If
        Next
        Trace.WriteLine("ClearDrug FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ClearInv()
        Trace.WriteLine("ClearInv STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        For Each i As Control In Panel4.Controls
            If TypeOf i Is TextBox Then
                i.Text = ""
            End If
        Next
        For Each combo As Control In Panel4.Controls
            If TypeOf combo Is ComboBox Then
                combo.Text = ""
            End If
        Next
        DTPickerNow()
        Trace.WriteLine("ClearInv FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub DTPickerNow()
        Trace.WriteLine("DTPickerNow STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        For Each L As Control In Panel4.Controls
            If TypeOf L Is DateTimePicker Then
                L.Text = CType(Now, String)
            End If
        Next
        Trace.WriteLine("DTPickerNow FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub SaveInves()
        Trace.WriteLine("SaveInves STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("SaveInves FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub UpdateInves()
        Trace.WriteLine("UpdateInves STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("UpdateInves FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub UpdateVisitDP()
        Trace.WriteLine("UpdateDrugs STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("UpdateDrugs FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub SaveVisitDP()
        Trace.WriteLine("SaveDrugs STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("SaveDrugs FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub UpdateAttach()
        Trace.WriteLine("UpdateAttach STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
                .Add("@AttDate", OleDbType.DBDate).Value = CDate(DTPickerAtt.Value)
                .Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtVisNo.Text))
            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If
        Trace.WriteLine("UpdateAttach FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub SaveAttach()
        Trace.WriteLine("SaveAttach STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
                .Add("@AttDate", OleDbType.DBDate).Value = CDate(DTPickerAtt.Value)

            End With

            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If

            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
        End If

        Trace.WriteLine("SaveAttach FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
                .Add("@Amount", OleDbType.VarChar).Value = txtAmount.Text
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
                .Add("@Amount", OleDbType.VarChar).Value = txtAmount.Text
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
        Trace.WriteLine("ShowAttachPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

                        DTPickerAtt.Text = dt.Rows(0).Item("AttDate").ToString
                    End If

                End Using
            End Using
        End Using
        Trace.WriteLine("ShowAttachPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowAttachVisTable()
        Trace.WriteLine("ShowAttachPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

                        DTPickerAtt.Text = dt.Rows(0).Item("AttDate").ToString
                    End If

                End Using
            End Using
        End Using
        Trace.WriteLine("ShowAttachPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowAttachTable()
        Trace.WriteLine("ShowAttachTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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

                        DTPickerAtt.Text = dt.Rows(0).Item("AttDate").ToString
                    End If

                End Using
            End Using
        End Using
        Trace.WriteLine("ShowAttachTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowInvPatTable()
        Trace.WriteLine("ShowInvPatTable STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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
        Trace.WriteLine("ShowInvTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowInvVisTable()
        Trace.WriteLine("ShowInvPatTable STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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
        Trace.WriteLine("ShowInvTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowInvTable()
        Trace.WriteLine("ShowInvTable STARED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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
        Trace.WriteLine("ShowInvTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisDPPatTable()
        Trace.WriteLine("ShowVisDPPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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
        Trace.WriteLine("ShowVisDPPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisDPVisTable()
        Trace.WriteLine("ShowVisDPPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()

            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM VisitDP WHERE Visit_no=@Visit_no " &
                                  "ORDER BY Visit_no"
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
        Trace.WriteLine("ShowVisDPPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisDPTable()
        Trace.WriteLine("ShowVisDPTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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
        Trace.WriteLine("ShowVisDPTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisitsPatTable()
        Trace.WriteLine("ShowVisitsPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits WHERE Patient_no=@Patient_no " '&
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
                        txtAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        Trace.WriteLine("ShowVisitsPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisitsVisTable()
        Trace.WriteLine("ShowVisitsPatTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits WHERE Visit_no=@Visit_no " &
                                  "ORDER BY Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtSearch.Text))
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
                        txtAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        Trace.WriteLine("ShowVisitsPatTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub ShowVisitsTable()
        Trace.WriteLine("ShowVisitsTable STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
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
                        txtAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
        Trace.WriteLine("ShowVisitsTable FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnVisSave_Click(sender As Object, e As EventArgs) Handles btnVisSave.Click
        Trace.WriteLine("btnVisSave_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        CheckNull("")

        SaveVisits()
        SaveVisitDP()
        SaveInves()
        SaveAttach()
        btnNewVisit.Enabled = True

        Trace.WriteLine("btnVisSave_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnPatient_Click(sender As Object, e As EventArgs) Handles btnPatient.Click
        Trace.WriteLine("btnPatient_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisPatNo.Text = "" And txtVisName.Text = "" Then
            GoToPatients()
            f1.txtNo.Text = GetAutonumber("Pat", "Patient_no")
            f1.txtNo.Select()
        ElseIf txtVisNo.Text <> GetAutonumber("Visits", "Visit_no") Then
            GoToPatients()
            f1.txtNo.Select()

        End If
        Trace.WriteLine("btnPatient_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub GoToPatients()
        Trace.WriteLine("GoToPatient STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        f1 = New Form1
        f1.Show()
        Me.Hide()

        f1.txtNo.Text = txtVisPatNo.Text
        f1.cbxPatName.Text = txtVisName.Text

        Trace.WriteLine("GoToPatient FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Trace.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If MsgBox("You Will Exit The Clinic" + vbCrLf +
                  "Are you sure ?", MsgBoxStyle.YesNo,
                  "Confirm Message") = vbNo Then
            Exit Sub
        End If
        Form1.Show()
        Form1.Close()
        End
        Trace.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnVisClear_Click(sender As Object, e As EventArgs) Handles btnVisClear.Click
        Trace.WriteLine("btnVisClear_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ListBox1.Items.Clear()
        txt1.Text = ""

        ClearData()
        ClearDrug()
        ClearInv()
        btnNewVisit.Enabled = False
        rdoVisit.Checked = True

        InvAndAttDisabled()
        lblcurTime.Text = Now.ToShortDateString

        Trace.WriteLine("btnVisClear_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub


    Private Sub btnNewVisit_Click(sender As Object, e As EventArgs) Handles btnNewVisit.Click
        Trace.WriteLine("btnNewVisit_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        txtComplain.Text = ""
        txtSign.Text = ""
        cbxDia.Text = ""
        cbxInter.Text = ""
        txtAmount.Text = ""
        ClearDrug()
        ClearInv()
        DTPickerNow()

        If txtVisNo.Text <> GetAutonumber("Visits", "Visit_no") Then

            txtVisNo.Text = GetAutonumber("Visits", "Visit_no")

        End If
        lblcurTime.Text = Now.ToShortDateString
        btnNewVisit.Enabled = False

        btnVisSave_Click(Nothing, Nothing)
        txtVisNo.Select()

        Trace.WriteLine("btnNewVisit_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private PrescriptionStr As String
    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Trace.WriteLine("btnPrint_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If chbxR.Checked = True Then

            PrescriptionStr = "                " & Label1.Text & "                                                  " & Label7.Text & vbNewLine &
                              "          " & Label2.Text & "                                        " & Label8.Text & vbNewLine &
                              "          " & Label3.Text & "                                                  " & Label9.Text & vbCrLf &
                              "          " & Label4.Text & "                                             " & Label10.Text & vbNewLine &
                              "          " & Label5.Text & "                                                 " & Label11.Text & vbCrLf &
                              "          " & Label6.Text & "                                             " & Label12.Text & vbCrLf & vbCrLf &
                              "          " & "Name  : " & txtVisName.Text & "                    " & "ID : " & txtVisPatNo.Text & vbCrLf &
                              "          " & "Diagnosis : " & cbxDia.Text & "     " & lblcurTime.Text & "        " & "No.: " & txtVisNo.Text & vbCrLf &
                              "          " & "R/: " & vbNewLine &
                              "          " & cbxDrug1.Text & vbNewLine & "                                                   " & cbxPlan1.Text & vbCrLf & "        " & cbxDrug2.Text & vbCrLf & "                                                  " & cbxPlan2.Text & vbCrLf &
                              "          " & cbxDrug3.Text & vbCrLf & "                                                  " & cbxPlan3.Text & vbCrLf & "        " & cbxDrug4.Text & vbCrLf & "                                                  " & cbxPlan4.Text & vbCrLf &
                              "          " & cbxDrug5.Text & vbCrLf & "                                                  " & cbxPlan5.Text & vbCrLf & "        " & cbxDrug6.Text & vbCrLf & "                                                  " & cbxPlan6.Text & vbCrLf &
                              "          " & cbxDrug7.Text & vbCrLf & "                                                  " & cbxPlan7.Text & vbCrLf & "        " & cbxDrug8.Text & vbCrLf & "                                                  " & cbxPlan8.Text & vbCrLf &
                              "          " & cbxDrug9.Text & vbCrLf & "                                                  " & cbxPlan9.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf &
                              "               " & Label13.Text
            'vbCrLf & "        " & cbxDrug10.Text & vbCrLf & "                                                  " & cbxPlan10.Text & vbCrLf
        ElseIf chbxR.Checked = False Then
            PrescriptionStr = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine &
                              "          " & "     " & txtVisName.Text & "                    " & "ID : " & txtVisPatNo.Text & vbCrLf &
                              "          " & "     " & lblcurTime.Text & "        " & "No.: " & txtVisNo.Text & vbCrLf & vbNewLine & vbNewLine &     '"    " & "R/: " & vbNewLine &
                              "               " & cbxDrug1.Text & vbNewLine & "                                              " & cbxPlan1.Text & vbCrLf & vbCrLf & "               " & cbxDrug2.Text & vbCrLf & "                                             " & cbxPlan2.Text & vbCrLf & vbCrLf &
                              "               " & cbxDrug3.Text & vbCrLf & "                                             " & cbxPlan3.Text & vbCrLf & vbCrLf & "               " & cbxDrug4.Text & vbCrLf & "                                             " & cbxPlan4.Text & vbCrLf & vbCrLf &
                              "               " & cbxDrug5.Text & vbCrLf & "                                             " & cbxPlan5.Text & vbCrLf & vbCrLf & "               " & cbxDrug6.Text & vbCrLf & "                                             " & cbxPlan6.Text & vbCrLf & vbCrLf &
                              "               " & cbxDrug7.Text & vbCrLf & "                                             " & cbxPlan7.Text & vbCrLf & vbCrLf & "               " & cbxDrug8.Text & vbCrLf & "                                             " & cbxPlan8.Text & vbCrLf & vbCrLf &
                              "               " & cbxDrug9.Text & vbCrLf & "                                             " & cbxPlan9.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        End If

        If cbxDrug1.Text <> "" And txtVisName.Text <> "" Then
            PrintDocument1.Print()
        End If
        Trace.WriteLine("btnPrint_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnPrintInv_Click(sender As Object, e As EventArgs) Handles btnPrintInv.Click
        ''##This condition for investigation's prescription with report if checkbox7 is "checked"
        If chbxRI.Checked = True Then

            PrescriptionStr = "                " & Label1.Text & "                                                  " & Label7.Text & vbNewLine &
                              "          " & Label2.Text & "                                        " & Label8.Text & vbNewLine &
                              "          " & Label3.Text & "                                                  " & Label9.Text & vbCrLf &
                              "          " & Label4.Text & "                                             " & Label10.Text & vbNewLine &
                              "          " & Label5.Text & "                                                 " & Label11.Text & vbCrLf &
                              "          " & Label6.Text & "                                             " & Label12.Text & vbCrLf & vbCrLf &
                              "          " & "Name : " & txtVisName.Text & vbCrLf &
                              "          " & "Patient ID : " & txtVisPatNo.Text & "          " & "Visit No. : " & txtVisNo.Text &
                              "          " & lblcurTime.Text & vbCrLf & vbCrLf & "                              " & " PLEASE FOR " & vbCrLf & vbCrLf &
                              "                    " & cbxInvest.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest1.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest2.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest3.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest4.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest5.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf &
                              "               " & Label13.Text
        ElseIf chbxRI.Checked = False Then
            PrescriptionStr = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine &
                              "               " & "     " & txtVisName.Text & vbCrLf &
                              "               " & "Visit No. : " & txtVisNo.Text & "     " & "Patient ID : " & txtVisPatNo.Text & "          " & lblcurTime.Text & vbCrLf & vbCrLf &
                              "                                        " & " PLEASE FOR " & vbCrLf & vbCrLf &
                              "                    " & cbxInvest.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest1.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest2.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest3.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest4.Text & vbCrLf & vbCrLf &
                              "                    " & cbxInvest5.Text & vbCrLf & vbCrLf
        End If
        If txtVisName.Text <> "" And cbxInvest.Text <> "" Then
            PrintDocument1.Print()
        End If

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Trace.WriteLine("PrintDocument1_PrintPage STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("PrintDocument1_PrintPage FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub SearchVisitNo()

        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM Visits WHERE Visit_no=@Visit_no"
                cmd.Parameters.Add("@Visit_no", OleDbType.Integer).Value = CInt(Val(txtSearch.Text))
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
                        txtAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub SearchName()

        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM [Visits]" &
                                  " WHERE Name LIKE '%" & txtSearch.Text & "%' " & 'LIKE '%" & txtSearch.Text & "%'
                                  "ORDER BY Visit_no"
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
                        txtAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using

    End Sub

    Private Sub SearchID()

        Using con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=TestDB.accdb;jet oledb:database password=mero1981923")
            con.Open()
            Using cmd As New OleDbCommand
                cmd.Connection = con
                cmd.CommandText = "SELECT * FROM [Visits]" &
                                  " WHERE Patient_no LIKE '%" & txtSearch.Text & "%' " & 'LIKE '%" & txtSearch.Text & "%'
                                  "ORDER BY Visit_no"
                cmd.Parameters.Add("@Patient_no", OleDbType.Integer).Value = CInt(Val(txtSearch.Text))
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
                        txtAmount.Text = dt1.Rows(0).Item("Amount").ToString
                        lblcurTime.Text = dt1.Rows(0).Item("VisDate").ToString
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub SearchtxtVisPatNo()
        'Dim con As New OleDbConnection("provider=microsoft.ace.oledb.12.0; data source=Dr_T.accdb;jet oledb:database password=mero1981923")
        'Dim cmd As New OleDbCommand
        'Dim da As New OleDbDataAdapter
        'Dim dt As New DataTable
        'Dim sSQL As String = String.Empty

        'Try
        '    con.Open()
        '    cmd.Connection = con
        '    cmd.CommandType = CommandType.Text
        '    sSQL = "SELECT * FROM Visits " &
        '           "WHERE Patient_no =" & Me.txtVisPatNo.Text & " " &
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

    Private Sub txtSearch_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtSearch.Validating

        If rdoVisit.Checked Then
            SearchVisitNo()
        ElseIf rdoName.Checked Then
            SearchName()
        ElseIf rdoID.Checked Then
            SearchID()
        End If

        ShowVisDPVisTable()
        'ShowInvVisTable()
        'ShowAttachVisTable()
        'Panel4.Visible = False
        InvAndAttDisabled()
        btnNewVisit.Enabled = True
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

    Sub SaveInXmlDiaInter()
        ''Writing XML content...
        Dim xmldoc As XmlDocument = New XmlDocument()
        xmldoc.Load(Directory.GetCurrentDirectory & "\DiaInter.xml")

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

    Sub SaveInXmlInv()
        Trace.WriteLine("SaveInXmlInv STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("SaveInXmlInv FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
    Sub SaveInXml()
        ''## These codes from https://stackoverflow.com/questions/33237851/write-combobox-selected-item-to-xml-file?answertab=votes#tab-top
        ''Writing XML content...
        Trace.WriteLine("SaveInXml STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("SaveInXml FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
    ''##For saving a new item in the Plan Xml file in the run time
    Sub SaveInXmlPlan()
        Trace.WriteLine("SaveInXmlPlan STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("SaveInXmlPlan FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
    ''##Fernado Soto from expert-exchange  https://www.experts-exchange.com/questions/29061270/In-which-event-of-a-combo-box-can-i-Load-XML-File.html?anchor=a42323104&notificationFollowed=198557353#a42323104
    Sub LoadXmlInv()
        Trace.WriteLine("LoadXmlInv STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("LoadXmlInv FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("LoadXmlDrug STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("LoadXmlDrug FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Sub RemoveDuplicateXml()
        ''##By Fernando Soto in this site: https://www.experts-exchange.com/questions/28454848/Help-with-removing-duplicates-in-an-xml-file-using-vb-net.html
        Trace.WriteLine("RemoveDuplicateXml STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("RemoveDuplicateXml FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
    ''##Removing Dplicates from Plan Xml file 
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
        Dim results = (From n In xdoc2.Descendants("Invset")
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

    Sub RDXmlDiaInter()
        Trace.WriteLine("RDXmlDiaInter STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Dim fileName As String = "DiaInter.xml"
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
        Trace.WriteLine("RDXmlDiaInter FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtVisNo_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtVisNo.Validating
        If txtVisNo.Text = GetAutonumber("Visits", "Visit_no") And txtVisName.Text <> "" Then

            btnVisSave_Click(Nothing, Nothing)

        End If
    End Sub

    Private Sub txtComplain_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtComplain.Validating
        Trace.WriteLine("txtComplain_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then

            UpdateVisits()

        End If

        Trace.WriteLine("txtComplain_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtSign_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtSign.Validating
        Trace.WriteLine("txtSign_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisits()

            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            ' Now fill the ComboBox's 
            cbxDia.Items.AddRange(cbElements)
        End If
        'txtSign.Refresh()
        Trace.WriteLine("txtSign_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDia_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDia.Validating
        Trace.WriteLine("cbxDia_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxDia_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDia_MouseEnter(sender As Object, e As EventArgs) Handles cbxDia.MouseEnter
        'RDXmlDiaInter()
    End Sub

    Private Sub cbxDia_Click(sender As Object, e As EventArgs) Handles cbxDia.Click
        Trace.WriteLine("txtDia_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDia.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("txtDia_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInter_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInter.Validating
        Trace.WriteLine("cbxInter_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxInter_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInter_Click(sender As Object, e As EventArgs) Handles cbxInter.Click
        Trace.WriteLine("cbxInter_Click STRTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\DiaInter.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInter.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxInter_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAmount_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtAmount.Validating
        Trace.WriteLine("txtAmount_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then

            UpdateVisits()
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug1.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("txtAmount_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug1.Validating
        Trace.WriteLine("cbxDrug1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ''##This condition i made it to prevent the error when moving between Drugs comboboxes by Tab or by mouse
        ''##And to avoid "try and catch" statement

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()

            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan1.Items.AddRange(PlElements)
        End If
        Trace.WriteLine("cbxDrug1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug1_Click(sender As Object, e As EventArgs) Handles cbxDrug1.Click
        Trace.WriteLine("cbxDrug1_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug1.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug1_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug2.Validating
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))


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
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug2_Click(sender As Object, e As EventArgs) Handles cbxDrug2.Click
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug2.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug3.Validating
        Trace.WriteLine("cbxDrug3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXml()
            Dim PlDoc = XElement.Load(Application.StartupPath + "\Plans.xml")
            '' Parse the XML document only once
            Dim PlElements = PlDoc.<Plans>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxPlan3.Items.AddRange(PlElements)
        End If
        Trace.WriteLine("cbxDrug3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug3_Click(sender As Object, e As EventArgs) Handles cbxDrug3.Click
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug3.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug4.Validating
        Trace.WriteLine("cbxDrug4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxDrug4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug4_Click(sender As Object, e As EventArgs) Handles cbxDrug4.Click
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug4.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug5.Validating
        Trace.WriteLine("cbxDrug5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxDrug5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug5_Click(sender As Object, e As EventArgs) Handles cbxDrug5.Click
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug5.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug6_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug6.Validating
        Trace.WriteLine("cbxDrug6_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxDrug6_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug6_Click(sender As Object, e As EventArgs) Handles cbxDrug6.Click
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug6.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug7_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug7.Validating
        Trace.WriteLine("cbxDrug7_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxDrug7_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug7_Click(sender As Object, e As EventArgs) Handles cbxDrug7.Click
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug7.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug8_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug8.Validating
        Trace.WriteLine("cbxDrug8_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxDrug8_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug8_Click(sender As Object, e As EventArgs) Handles cbxDrug8.Click
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug8.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug9_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug9.Validating
        Trace.WriteLine("cbxDrug9_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxDrug9_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug9_Click(sender As Object, e As EventArgs) Handles cbxDrug9.Click
        Trace.WriteLine("cbxDrug2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Drugs1.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Drugs>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxDrug9.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxDrug2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxDrug10_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxDrug10.Validating
        Trace.WriteLine("cbxDrug10_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxDrug10_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan1.Validating
        Trace.WriteLine("cbxPlan1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxPlan1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan2.Validating
        Trace.WriteLine("cbxPlan2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxPlan2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan3.Validating
        Trace.WriteLine("cbxPlan3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxPlan3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan4.Validating
        Trace.WriteLine("cbxPlan4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxPlan4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan5.Validating
        Trace.WriteLine("cbxPlan5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxPlan5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan6_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan6.Validating
        Trace.WriteLine("cbxPlan6_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxPlan6_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan7_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan7.Validating
        Trace.WriteLine("cbxPlan7_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxPlan7_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan8_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan8.Validating
        Trace.WriteLine("cbxPlan8_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxPlan8_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan9_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan9.Validating
        Trace.WriteLine("cbxPlan9_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("cbxPlan9_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan10_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxPlan10.Validating
        Trace.WriteLine("cbxPlan10_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateVisitDP()
            SaveInXmlPlan()
        End If
        Trace.WriteLine("cbxPlan10_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Trace.WriteLine("btnRefresh_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        RemoveDuplicateXml()
        RDXmlInv()
        RDXmlPlan()
        RDXmlDiaInter()
        RDXmlInvRes()

        loaddata()
        Trace.WriteLine("btnRefresh_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxPlan1_MouseClick(sender As Object, e As MouseEventArgs) Handles cbxPlan1.MouseClick
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
        Trace.WriteLine("Panel4_Paint STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Panel4.AutoScroll = True
        Trace.WriteLine("Panel4_Paint FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest.Validating
        Trace.WriteLine("cbxInvest_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxInvest_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest_Click(sender As Object, e As EventArgs) Handles cbxInvest.Click
        Trace.WriteLine("cbxInvest_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxInvest_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest1_Validating(sender As Object, e As EventArgs) Handles cbxInvest1.Validating
        Trace.WriteLine("cbxInvest1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxInvest1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest1_Click(sender As Object, e As EventArgs) Handles cbxInvest1.Click
        Trace.WriteLine("cbxInvest1_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest1.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxInvest1_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest2.Validating
        Trace.WriteLine("cbxInvest2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxInvest2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest2_Click(sender As Object, e As EventArgs) Handles cbxInvest2.Click
        Trace.WriteLine("cbxInvest2_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest2.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxInvest2_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest3.Validating
        Trace.WriteLine("cbxInvest3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxInvest3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest3_Click(sender As Object, e As EventArgs) Handles cbxInvest3.Click
        Trace.WriteLine("cbxInvest3_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest3.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxInvest3_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest4.Validating
        Trace.WriteLine("cbxInvest4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxInvest4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest4_Click(sender As Object, e As EventArgs) Handles cbxInvest4.Click
        Trace.WriteLine("cbxInvest4_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest4.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxInvest4_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxInvest5.Validating
        Trace.WriteLine("cbxInvest5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
            SaveInXmlInv()
        End If
        Trace.WriteLine("cbxInvest5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxInvest5_Click(sender As Object, e As EventArgs) Handles cbxInvest5.Click
        Trace.WriteLine("cbxInvest5_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\Investigations.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxInvest5.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxInvest5_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv.Validating
        Trace.WriteLine("DTPickerInv_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        Trace.WriteLine("DTPickerInv_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv1.Validating
        Trace.WriteLine("DTPickerInv1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        Trace.WriteLine("DTPickerInv1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv2.Validating
        Trace.WriteLine("DTPickerInv2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        Trace.WriteLine("DTPickerInv2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv3.Validating
        Trace.WriteLine("DTPickerInv3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        Trace.WriteLine("DTPickerInv3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv4.Validating
        Trace.WriteLine("DTPickerInv4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        Trace.WriteLine("DTPickerInv4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerInv5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerInv5.Validating
        Trace.WriteLine("DTPickerInv5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        Trace.WriteLine("DTPickerInv5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult.Validating
        Trace.WriteLine("cbxResult_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxResult_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult_Click(sender As Object, e As EventArgs) Handles cbxResult.Click
        Trace.WriteLine("cbxResult_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxResult_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult_MouseEnter(sender As Object, e As EventArgs) Handles cbxResult.MouseEnter
        Trace.WriteLine("cbxResult_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxResult_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult1.Validating
        Trace.WriteLine("cbxResult1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxResult1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult1_Click(sender As Object, e As EventArgs) Handles cbxResult1.Click
        Trace.WriteLine("cbxResult1_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult1.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxResult1_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult2.Validating
        Trace.WriteLine("cbxResult2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxResult2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult2_Click(sender As Object, e As EventArgs) Handles cbxResult2.Click
        Trace.WriteLine("cbxResult2_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult2.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxResult2_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult3.Validating
        Trace.WriteLine("cbxResult3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxResult3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult3_Click(sender As Object, e As EventArgs) Handles cbxResult3.Click
        Trace.WriteLine("cbxResult3_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult3.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxResult3_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult4.Validating
        Trace.WriteLine("cbxResult4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

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
        Trace.WriteLine("cbxResult4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult4_Click(sender As Object, e As EventArgs) Handles cbxResult4.Click
        Trace.WriteLine("cbxResult4_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult4.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxResult4_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cbxResult5.Validating
        Trace.WriteLine("cbxResult5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateInves()
        End If
        Trace.WriteLine("cbxResult5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub cbxResult5_Click(sender As Object, e As EventArgs) Handles cbxResult5.Click
        Trace.WriteLine("cbxResult5_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            '' Read the XML file from disk only once
            Dim xDoc = XElement.Load(Application.StartupPath + "\InvRes.xml")
            '' Parse the XML document only once
            Dim cbElements = xDoc.<Invest>.Select(Function(n) n.Value).ToArray()
            '' Now fill the ComboBox's 
            cbxResult5.Items.AddRange(cbElements)
        End If
        Trace.WriteLine("cbxResult5_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo1_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo1.Validating
        Trace.WriteLine("txtCo1_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        Trace.WriteLine("txtCo1_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo2_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo2.Validating
        Trace.WriteLine("txtCo2_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        Trace.WriteLine("txtCo2_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo3_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo3.Validating
        Trace.WriteLine("txtCo3_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
            'btnVisSave_Click(Nothing, Nothing)
        End If
        Trace.WriteLine("txtCo3_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo4_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo4.Validating
        Trace.WriteLine("txtCo4_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
            'btnVisSave_Click(Nothing, Nothing)
        End If
        Trace.WriteLine("txtCo4_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtCo5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtCo5.Validating
        Trace.WriteLine("txtCo5_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
            'btnVisSave_Click(Nothing, Nothing)
        End If
        Trace.WriteLine("txtCo5_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt1.MouseDoubleClick
        Trace.WriteLine("txtAtt1_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        ''##This condition for prevent error when "txtAtt textbox" is empty
        If txtAtt1.Text <> "" Then
            Process.Start(Me.txtAtt1.Text)
        End If
        Trace.WriteLine("txtAtt1_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt2_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt2.MouseDoubleClick
        Trace.WriteLine("txtAtt2_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtAtt2.Text <> "" Then
            Process.Start(Me.txtAtt2.Text)
        End If
        Trace.WriteLine("txtAtt2_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt3_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt3.MouseDoubleClick
        Trace.WriteLine("txtAtt3_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtAtt3.Text <> "" Then
            Process.Start(Me.txtAtt3.Text)
        End If
        Trace.WriteLine("txtAtt3_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt4_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt4.MouseDoubleClick
        Trace.WriteLine("txtAtt4_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtAtt4.Text <> "" Then
            Process.Start(Me.txtAtt4.Text)
        End If
        Trace.WriteLine("txtAtt4_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub txtAtt5_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles txtAtt5.MouseDoubleClick
        Trace.WriteLine("txtAtt5_MouseDoubleClick STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtAtt5.Text <> "" Then
            Process.Start(Me.txtAtt5.Text)
        End If
        Trace.WriteLine("txtAtt5_MouseDoubleClick FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub DTPickerAtt_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles DTPickerAtt.Validating
        Trace.WriteLine("DTPickerAtt_Validating STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        If txtVisName.Text <> "" And txtVisPatNo.Text <> "" Then
            UpdateAttach()
        End If
        Trace.WriteLine("DTPickerAtt_Validating FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Dim f As New OpenFileDialog
    Private Sub btnOpen1_Click(sender As Object, e As EventArgs) Handles btnOpen1.Click
        ''##from https://answers.microsoft.com/en-us/windows/forum/windows8_1-winapps-appother/storing-and-retrieving-a-file-path-using-access/374b9f15-77c3-4348-bf75-676658c9bb6b?tm=1506765470714&rtAction=1506810163404
        Trace.WriteLine("btnOpen1_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("btnOpen1_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnopen2_Click(sender As Object, e As EventArgs) Handles btnOpen2.Click
        Trace.WriteLine("btnOpen2_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("btnOpen2_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnopen3_Click(sender As Object, e As EventArgs) Handles btnOpen3.Click
        Trace.WriteLine("btnOpen3_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("btnOpen3_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnopen4_Click(sender As Object, e As EventArgs) Handles btnOpen4.Click
        Trace.WriteLine("btnOpen4_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("btnOpen4_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnopen5_Click(sender As Object, e As EventArgs) Handles btnOpen5.Click
        Trace.WriteLine("btnOpen5_Click STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
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
        Trace.WriteLine("btnOpen5_Click FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint
        Trace.WriteLine("Panel1_Paint STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        Panel1.AutoScroll = True
        Trace.WriteLine("Panel1_Paint FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub Panel1_MouseHover(sender As Object, e As EventArgs) Handles Panel1.MouseHover
        Trace.WriteLine("Panel1_MouseHover STARTED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        'Panel1.Select()
        Trace.WriteLine("Panel1_MouseHover FINISHED @ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Private Sub btnInv_Click(sender As Object, e As EventArgs) Handles btnInv.Click
        ListBox1.Items.Clear()
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
                ListBox1.Items.Add(item).ToString()
            End While
            reader.Close()
            conn.Close()
        End If

        InvAndAttEnabled()
        'Panel4.Visible = True
        ShowInvVisTable()
        ShowAttachVisTable()

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        If ListBox1.SelectedIndex > -1 Then
            txt1.Text = CType(ListBox1.SelectedItem, String)
        End If
        InvAndAttEnabled()

    End Sub

    Private Sub txt1_TextChanged(sender As Object, e As EventArgs) Handles txt1.TextChanged

        ShowVisitsTable()
        ShowVisDPTable()
        ShowInvTable()
        ShowAttachTable()
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        txt1.Text = ""
        ListBox1.Items.Clear()
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

    End Sub


End Class