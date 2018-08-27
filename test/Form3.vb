Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.IO

Module Hello
    Sub Main()
        ' Create an instance of the licensed application
        Dim app As frmMain = Nothing
        Try
            ' This will throw a LicenseException if the 
            ' license is invalid... If we get an exception,
            ' "app" will remain null and the Run() method
            ' (below) will not be executed...
            app = New frmMain
        Catch ex As Exception
            ' Catch any error, but especially licensing errors...
            Dim strErr As String = String.Format("Error executing application: '{0}'", ex.Message)
            MessageBox.Show(strErr, "VB RegistryLicensedApplication Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        If Not app Is Nothing Then
            Application.Run(app)
        End If
    End Sub
End Module

<LicenseProviderAttribute(GetType(RegistryLicenseProvider)),
 GuidAttribute("2de915e1-df71-3443-9f4d-32259c92ced2")>
Public Class frmMain
    Inherits System.Windows.Forms.Form

    Private _license As License = Nothing

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Obtain the license
        Me._license = LicenseManager.Validate(GetType(frmMain), Me)

        Dim f1 As New Form1
        f1.Show()
        Me.Hide()

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If

            If Not _license Is Nothing Then
                Me._license.Dispose()
                Me._license = Nothing
            End If

        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(8, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 36)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "A Licensed KMAClinic Application." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'cmdExit
        '
        Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdExit.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        Me.cmdExit.Location = New System.Drawing.Point(230, 12)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(81, 38)
        Me.cmdExit.TabIndex = 1
        Me.cmdExit.Text = "My Clinic"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(321, 60)
        Me.ControlBox = False
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "Licensed Application"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        ' Just Open a Form1...
        'Dim f1 As New Form1
        'f1.Hide()
        Me.Hide()
        'f1.Show()
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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

        'CompactAccessDatabase()
        'BackupXML()
    End Sub
    '## Paul https://social.msdn.microsoft.com/Forums/vstudio/en-US/35b9de93-e5fd-4e3f-a8f6-97516184d4c7/what-is-the-best-solution-for-compact-and-repair-access-2013-database?forum=vbgeneral
    Sub CompactAccessDatabase()

        '##Path For Real Projects by "Application.StartupPath"
        'Dim DatabasePath As String = "D:KMAClinic\bin\Release\Dr_T.accdb"
        Dim DatabasePath As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\TestDB.accdb")

        '##For Real Projects
        Dim DatabasePathCompacted As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\_" & Format(Now(), "ddMMyyyy_hhmmss") & ".accdb")

        Dim CompactDB As New Microsoft.Office.Interop.Access.Dao.DBEngine

        '##Here you can write your database password with this method (DatabasePath, DatabasePathCompacted, , , ";pwd=mero1981923")
        CompactDB.CompactDatabase(DatabasePath, DatabasePathCompacted, , , ";pwd=mero1981923")
        CompactDB = Nothing

        Dim backuppath As String = Path.Combine(Application.StartupPath, Directory.GetCurrentDirectory + "\Backups\Docs\TestDB_" &
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

End Class