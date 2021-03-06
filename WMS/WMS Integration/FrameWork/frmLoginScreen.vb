Imports System.IO
Imports System.IO.Directory

Imports System.Diagnostics.Process
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports SAPbobsCOM

Public Class frmLoginScreen
    Inherits System.Windows.Forms.Form

#Region "Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If

        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnReferesh As System.Windows.Forms.Button
    Friend WithEvents txtServerPwd As System.Windows.Forms.TextBox
    Friend WithEvents txtDBUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtServerName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents cmbCompanyDB As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCompanyPWD As System.Windows.Forms.TextBox
    Friend WithEvents txtSBOUserName As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    '  Friend WithEvents Show_Click As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents fldFolderBrowse As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbMainSAPCompany As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents MainSAPPWD As System.Windows.Forms.TextBox
    Friend WithEvents MainSAPUID As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents MainSQLPWD As System.Windows.Forms.TextBox
    Friend WithEvents MainSQLUID As System.Windows.Forms.TextBox
    Friend WithEvents MainSQLServer As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents LocalLicenseServer As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtDirectory As System.Windows.Forms.TextBox
    Friend WithEvents mainLicenseServer As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cmbServertype As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cmbMainServertype As System.Windows.Forms.ComboBox
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.btnReferesh = New System.Windows.Forms.Button
        Me.txtServerPwd = New System.Windows.Forms.TextBox
        Me.txtDBUserName = New System.Windows.Forms.TextBox
        Me.txtServerName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LocalLicenseServer = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtDirectory = New System.Windows.Forms.TextBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.btnConnect = New System.Windows.Forms.Button
        Me.cmbCompanyDB = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtCompanyPWD = New System.Windows.Forms.TextBox
        Me.txtSBOUserName = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.fldFolderBrowse = New System.Windows.Forms.FolderBrowserDialog
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.mainLicenseServer = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmbMainSAPCompany = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.MainSAPPWD = New System.Windows.Forms.TextBox
        Me.MainSAPUID = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.MainSQLPWD = New System.Windows.Forms.TextBox
        Me.MainSQLUID = New System.Windows.Forms.TextBox
        Me.MainSQLServer = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbServertype = New System.Windows.Forms.ComboBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.cmbMainServertype = New System.Windows.Forms.ComboBox
        Me.Frame1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Label16)
        Me.Frame1.Controls.Add(Me.cmbServertype)
        Me.Frame1.Controls.Add(Me.btnReferesh)
        Me.Frame1.Controls.Add(Me.txtServerPwd)
        Me.Frame1.Controls.Add(Me.txtDBUserName)
        Me.Frame1.Controls.Add(Me.txtServerName)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(12, 12)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(327, 158)
        Me.Frame1.TabIndex = 9
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Local Server Details"
        '
        'btnReferesh
        '
        Me.btnReferesh.Location = New System.Drawing.Point(121, 127)
        Me.btnReferesh.Name = "btnReferesh"
        Me.btnReferesh.Size = New System.Drawing.Size(163, 23)
        Me.btnReferesh.TabIndex = 4
        Me.btnReferesh.Text = "Refresh/Update Company List"
        '
        'txtServerPwd
        '
        Me.txtServerPwd.Location = New System.Drawing.Point(132, 72)
        Me.txtServerPwd.Name = "txtServerPwd"
        Me.txtServerPwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtServerPwd.Size = New System.Drawing.Size(152, 20)
        Me.txtServerPwd.TabIndex = 3
        '
        'txtDBUserName
        '
        Me.txtDBUserName.Location = New System.Drawing.Point(132, 48)
        Me.txtDBUserName.Name = "txtDBUserName"
        Me.txtDBUserName.Size = New System.Drawing.Size(152, 20)
        Me.txtDBUserName.TabIndex = 2
        '
        'txtServerName
        '
        Me.txtServerName.Location = New System.Drawing.Point(132, 24)
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.Size = New System.Drawing.Size(152, 20)
        Me.txtServerName.TabIndex = 1
        Me.txtServerName.Text = "(local)"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 20)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = " User Name"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 20)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Server"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.LocalLicenseServer)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtDirectory)
        Me.GroupBox1.Controls.Add(Me.Button5)
        Me.GroupBox1.Controls.Add(Me.btnConnect)
        Me.GroupBox1.Controls.Add(Me.cmbCompanyDB)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtCompanyPWD)
        Me.GroupBox1.Controls.Add(Me.txtSBOUserName)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(8, 176)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(331, 177)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        '
        'LocalLicenseServer
        '
        Me.LocalLicenseServer.Location = New System.Drawing.Point(136, 16)
        Me.LocalLicenseServer.Name = "LocalLicenseServer"
        Me.LocalLicenseServer.Size = New System.Drawing.Size(152, 20)
        Me.LocalLicenseServer.TabIndex = 19
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(16, 16)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(109, 20)
        Me.Label15.TabIndex = 20
        Me.Label15.Text = "SBO License Server"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(20, 119)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(109, 20)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Log Directory"
        '
        'txtDirectory
        '
        Me.txtDirectory.Location = New System.Drawing.Point(132, 119)
        Me.txtDirectory.Name = "txtDirectory"
        Me.txtDirectory.Size = New System.Drawing.Size(152, 20)
        Me.txtDirectory.TabIndex = 17
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(285, 118)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(40, 23)
        Me.Button5.TabIndex = 16
        Me.Button5.Text = "..."
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(93, 148)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(152, 23)
        Me.btnConnect.TabIndex = 10
        Me.btnConnect.Text = "Test Local Connection"
        '
        'cmbCompanyDB
        '
        Me.cmbCompanyDB.Location = New System.Drawing.Point(132, 42)
        Me.cmbCompanyDB.Name = "cmbCompanyDB"
        Me.cmbCompanyDB.Size = New System.Drawing.Size(152, 22)
        Me.cmbCompanyDB.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(109, 20)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Company DB"
        '
        'txtCompanyPWD
        '
        Me.txtCompanyPWD.Location = New System.Drawing.Point(132, 92)
        Me.txtCompanyPWD.Name = "txtCompanyPWD"
        Me.txtCompanyPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCompanyPWD.Size = New System.Drawing.Size(152, 20)
        Me.txtCompanyPWD.TabIndex = 7
        '
        'txtSBOUserName
        '
        Me.txtSBOUserName.Location = New System.Drawing.Point(132, 68)
        Me.txtSBOUserName.Name = "txtSBOUserName"
        Me.txtSBOUserName.Size = New System.Drawing.Size(152, 20)
        Me.txtSBOUserName.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 92)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(109, 20)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Password"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(109, 20)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "SBO User Name"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem1, Me.MenuItem3})
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Service Log File"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.Text = "BP Log File"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(110, 148)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(152, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "Test Remote Connection"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox2.Controls.Add(Me.mainLicenseServer)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.cmbMainSAPCompany)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.MainSAPPWD)
        Me.GroupBox2.Controls.Add(Me.MainSAPUID)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox2.Location = New System.Drawing.Point(358, 176)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox2.Size = New System.Drawing.Size(331, 177)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        '
        'mainLicenseServer
        '
        Me.mainLicenseServer.Location = New System.Drawing.Point(132, 16)
        Me.mainLicenseServer.Name = "mainLicenseServer"
        Me.mainLicenseServer.Size = New System.Drawing.Size(152, 20)
        Me.mainLicenseServer.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(6, 19)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(109, 20)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "SBO License Server"
        '
        'cmbMainSAPCompany
        '
        Me.cmbMainSAPCompany.Location = New System.Drawing.Point(132, 46)
        Me.cmbMainSAPCompany.Name = "cmbMainSAPCompany"
        Me.cmbMainSAPCompany.Size = New System.Drawing.Size(152, 22)
        Me.cmbMainSAPCompany.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 46)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(109, 20)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "Company DB"
        '
        'MainSAPPWD
        '
        Me.MainSAPPWD.Location = New System.Drawing.Point(132, 96)
        Me.MainSAPPWD.Name = "MainSAPPWD"
        Me.MainSAPPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.MainSAPPWD.Size = New System.Drawing.Size(152, 20)
        Me.MainSAPPWD.TabIndex = 7
        '
        'MainSAPUID
        '
        Me.MainSAPUID.Location = New System.Drawing.Point(132, 72)
        Me.MainSAPUID.Name = "MainSAPUID"
        Me.MainSAPUID.Size = New System.Drawing.Size(152, 20)
        Me.MainSAPUID.TabIndex = 6
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(16, 96)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(109, 20)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "Password"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(16, 72)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(109, 20)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "SBO User Name"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Controls.Add(Me.Label17)
        Me.GroupBox3.Controls.Add(Me.Button3)
        Me.GroupBox3.Controls.Add(Me.cmbMainServertype)
        Me.GroupBox3.Controls.Add(Me.MainSQLPWD)
        Me.GroupBox3.Controls.Add(Me.MainSQLUID)
        Me.GroupBox3.Controls.Add(Me.MainSQLServer)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox3.Location = New System.Drawing.Point(362, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox3.Size = New System.Drawing.Size(336, 158)
        Me.GroupBox3.TabIndex = 11
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = " Remote Server Details"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(128, 118)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(163, 23)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Refresh/Update Company List"
        '
        'MainSQLPWD
        '
        Me.MainSQLPWD.Location = New System.Drawing.Point(132, 72)
        Me.MainSQLPWD.Name = "MainSQLPWD"
        Me.MainSQLPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.MainSQLPWD.Size = New System.Drawing.Size(152, 20)
        Me.MainSQLPWD.TabIndex = 3
        '
        'MainSQLUID
        '
        Me.MainSQLUID.Location = New System.Drawing.Point(132, 48)
        Me.MainSQLUID.Name = "MainSQLUID"
        Me.MainSQLUID.Size = New System.Drawing.Size(152, 20)
        Me.MainSQLUID.TabIndex = 2
        '
        'MainSQLServer
        '
        Me.MainSQLServer.Location = New System.Drawing.Point(132, 24)
        Me.MainSQLServer.Name = "MainSQLServer"
        Me.MainSQLServer.Size = New System.Drawing.Size(152, 20)
        Me.MainSQLServer.TabIndex = 1
        Me.MainSQLServer.Text = "(local)"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(16, 72)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 20)
        Me.Label12.TabIndex = 4
        Me.Label12.Text = "Password"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(16, 48)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(100, 20)
        Me.Label13.TabIndex = 5
        Me.Label13.Text = " User Name"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(16, 24)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(100, 20)
        Me.Label14.TabIndex = 6
        Me.Label14.Text = "Server"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 399)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(125, 23)
        Me.Button2.TabIndex = 14
        Me.Button2.Text = "Export Data"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(144, 399)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(125, 23)
        Me.Button4.TabIndex = 15
        Me.Button4.Text = "Close"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 20)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Password"
        '
        'cmbServertype
        '
        Me.cmbServertype.Items.AddRange(New Object() {"2005", "2008"})
        Me.cmbServertype.Location = New System.Drawing.Point(132, 98)
        Me.cmbServertype.Name = "cmbServertype"
        Me.cmbServertype.Size = New System.Drawing.Size(152, 22)
        Me.cmbServertype.TabIndex = 21
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(15, 98)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(100, 20)
        Me.Label16.TabIndex = 22
        Me.Label16.Text = "Server Type"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(15, 95)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(100, 20)
        Me.Label17.TabIndex = 24
        Me.Label17.Text = "Server Type"
        '
        'cmbMainServertype
        '
        Me.cmbMainServertype.Items.AddRange(New Object() {"2005", "2008"})
        Me.cmbMainServertype.Location = New System.Drawing.Point(132, 95)
        Me.cmbMainServertype.Name = "cmbMainServertype"
        Me.cmbMainServertype.Size = New System.Drawing.Size(152, 22)
        Me.cmbMainServertype.TabIndex = 23
        '
        'frmLoginScreen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(757, 481)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Frame1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "frmLoginScreen"
        Me.Text = "Choose Company Database"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region

#Region "Declaration"
    Public oCompany As SAPbobsCOM.Company
    Dim objPC = New clsMain
    Dim sValue As String
    Dim sPath As String
    Private oFileDt As DataTable
    Dim index As Integer
    Dim oHeaderDt As DataTable
    Dim oDetailDt As DataTable

    Dim myConnection As SqlConnection
    Dim oCommand As SqlCommand
    Dim oSqlAdap As SqlDataAdapter
    Dim TransactionCode As String = String.Empty
    '  Private oCompany As SAPbobsCOM.Company
    'Dim oHeaderDt As DataTable
    'Dim oDetailDt As DataTable
    Dim oLineDt_C As DataTable
    Dim sQuery As String

#End Region

#Region "Read School Files"
    Public Function mainFunction1() As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim strPath As String = strExportFilePaty ' System.Configuration.ConfigurationManager.AppSettings("OpenFolder").ToString
            Dim txtFiles = Directory.GetFileSystemEntries(strPath, "*.txt")
            Dim strImpFileName, currentFile As String

            Dim di As New IO.DirectoryInfo(strExportFilePaty)
            Dim aryFi As IO.FileInfo() = di.GetFiles("*.txt")
            Dim fi As IO.FileInfo
            oCompany = objRemoteCompany
            '       connectSAP()
            For Each fi In aryFi
                strImpFileName = fi.FullName
                currentFile = strImpFileName
                traceService("Processing  " & currentFile)
                oFileDt = New DataTable()
                If currentFile.Length > 0 Then
                    Dim intCol As Integer = 0
                    Dim txtRows() As String
                    Dim fields() As String
                    Dim oDr As DataRow
                    txtRows = System.IO.File.ReadAllLines(currentFile)
                    Dim intRow As Integer = 0
                    For Each txtrow As String In txtRows
                        If intRow = 0 Then
                            fields = txtrow.Split(vbTab)
                            For index As Integer = 0 To fields.Length - 1
                                oFileDt.Columns.Add(fields(intCol).ToUpper(), GetType(String)).Caption = fields(intCol).ToUpper()
                                intCol += 1
                            Next
                            Exit For
                        End If
                    Next

                    intRow = 0
                    For Each txtrow As String In txtRows
                        If intRow = 0 Then

                        ElseIf intRow > 0 Then
                            fields = txtrow.Split(vbTab)
                            oDr = oFileDt.NewRow()
                            oDr.ItemArray = fields
                            oFileDt.Rows.Add(oDr)
                        End If
                        intRow = intRow + 1
                    Next
                End If
                traceService("Create Student Informations Data")
                PushToDIAPI(oFileDt)
                MoveToFolder(currentFile)
            Next
            'If 1 = 1 Then 'txtFiles.Count() > 0 Then

            '    For Each currentFile As String In txtFiles
            '        traceService(currentFile)
            '        oFileDt = New DataTable()
            '        If currentFile.Length > 0 Then
            '            Dim intCol As Integer = 0
            '            Dim txtRows() As String
            '            Dim fields() As String
            '            Dim oDr As DataRow
            '            txtRows = System.IO.File.ReadAllLines(currentFile)
            '            Dim intRow As Integer = 0
            '            For Each txtrow As String In txtRows
            '                If intRow = 0 Then
            '                    fields = txtrow.Split(vbTab)
            '                    For index As Integer = 0 To fields.Length - 1
            '                        oFileDt.Columns.Add(fields(intCol).ToUpper(), GetType(String)).Caption = fields(intCol).ToUpper()
            '                        intCol += 1
            '                    Next
            '                    Exit For
            '                End If
            '            Next

            '            intRow = 0
            '            For Each txtrow As String In txtRows
            '                If intRow = 0 Then

            '                ElseIf intRow > 0 Then
            '                    fields = txtrow.Split(vbTab)
            '                    oDr = oFileDt.NewRow()
            '                    oDr.ItemArray = fields
            '                    oFileDt.Rows.Add(oDr)
            '                End If
            '                intRow = intRow + 1
            '            Next
            '        End If
            '        traceService("Pushing Data")
            '        PushToDIAPI(oFileDt)
            '        MoveToFolder(currentFile)
            '    Next

            '     End If
            disconnectSAP()
            Return _retVal
        Catch ex As Exception
            traceService(ex.ToString())
        Finally
            disconnectSAP()
        End Try
    End Function

    Private Sub PushToDIAPI(ByVal oDT As DataTable)
        Try
            Dim oBP As SAPbobsCOM.BusinessPartners
            If oDT.Rows.Count > 0 Then
                For index = 0 To oDT.Rows.Count - 1
                    Dim strCardCode As String = oDT.Rows(index)("Family Id").ToString()
                    Dim strCustomer As String = String.Empty
                    Dim strFamilyContName As String = ""
                    Dim strPassword As String = ""
                    oBP = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    If oBP.GetByKey(strCardCode) Then
                        If oDT.Rows(index)("Paying").ToString = "F" Then
                            oBP.CardName = oDT.Rows(index)("Father").ToString.Replace("""", "")
                            strCustomer = oDT.Rows(index)("Father").ToString.Replace("""", "")
                            strFamilyContName = oDT.Rows(index)("Mother").ToString.Replace("""", "")
                            oBP.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                            oBP.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                            oBP.Cellular = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                            oBP.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                        Else
                            oBP.CardName = oDT.Rows(index)("Mother").ToString.Replace("""", "")
                            strCustomer = oDT.Rows(index)("Mother").ToString.Replace("""", "")
                            strFamilyContName = oDT.Rows(index)("Father").ToString.Replace("""", "")
                            oBP.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                            oBP.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                            oBP.Cellular = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                            oBP.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                        End If

                        'oBP.ContactPerson = oDt.Rows(index)(3)
                        Dim blnAExist As Boolean = False
                        For intAddress As Integer = 0 To oBP.Addresses.Count - 1
                            oBP.Addresses.SetCurrentLine(intAddress)
                            If oBP.Addresses.AddressName2 = strCustomer Then
                                blnAExist = True
                                oBP.Addresses.AddressName = "Home Address"
                                oBP.Addresses.AddressName2 = strCustomer
                                oBP.Addresses.Street = oDT.Rows(index)("Street").ToString.Replace("""", "")
                                oBP.Addresses.ZipCode = oDT.Rows(index)("Zip").ToString.Replace("""", "")
                                oBP.Addresses.City = oDT.Rows(index)("City").ToString.Replace("""", "")
                                'oBP.Addresses.Country = oDt.Rows(index)("State Country")
                                oBP.Addresses.Country = "BH"
                                oBP.Addresses.GlobalLocationNumber = oDT.Rows(index)("Mailing City").ToString.Replace("""", "")
                                oBP.Addresses.BuildingFloorRoom = oDT.Rows(index)("Mailing Street").ToString.Replace("""", "")
                                oBP.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                            End If
                        Next


                        If Not blnAExist Then
                            If oBP.Addresses.Count > 0 Then
                                oBP.Addresses.Add()
                                ' oBP.Addresses.SetCurrentLine(oBP.Addresses.Count - 1)
                                oBP.Addresses.AddressName = "Home Address"
                                oBP.Addresses.AddressName2 = strCustomer.ToString.Replace("""", "")
                                oBP.Addresses.AddressName = strCustomer.ToString.Replace("""", "")
                                oBP.Addresses.Street = oDT.Rows(index)("Street").ToString.Replace("""", "")
                                oBP.Addresses.ZipCode = oDT.Rows(index)("Zip").ToString.Replace("""", "")
                                oBP.Addresses.City = oDT.Rows(index)("City").ToString.Replace("""", "")
                                'oBP.Addresses.Country = oDt.Rows(index)("State Country")
                                oBP.Addresses.Country = "BH"
                                oBP.Addresses.GlobalLocationNumber = oDT.Rows(index)("Mailing City").ToString.Replace("""", "")
                                oBP.Addresses.BuildingFloorRoom = oDT.Rows(index)("Mailing Street").ToString.Replace("""", "")
                                oBP.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                            End If
                        End If

                        oBP.EmailAddress = oDT.Rows(index)("Guardianemail").ToString.Replace("""", "")
                        oBP.Website = oDT.Rows(index)("Special Conditions Custom").ToString.Replace("""", "")
                        oBP.Password = oDT.Rows(index)("Custody Role").ToString.Replace("""", "")
                        oBP.AdditionalID = oDT.Rows(index)("Pick Up").ToString.Replace("""", "")
                        oBP.CompanyRegistrationNumber = oDT.Rows(index)("Pick Up Contact Mobile").ToString.Replace("""", "")
                        Dim oTes As SAPbobsCOM.Recordset
                        oTes = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTes.DoQuery("SELECT T0.[IndCode], T0.[IndName], T0.[IndDesc] FROM OOND T0 where IndName='" & oDT.Rows(index)("Section") & "'")
                        If oTes.RecordCount > 0 Then
                            oBP.Industry = oTes.Fields.Item(0).Value
                        End If
                        'oBP.Industry = oDt.Rows(index)("Section")
                        Dim strEmp As String()
                        Dim strName As String = oDT.Rows(index)("Staff Name").ToString.Replace("""", "")
                        strName = strName.Replace("''", "")
                        strEmp = strName.Split(",")
                        Dim strFirstName, strMiddleName, strlatName, strNameCondition As String
                        If strEmp.Length < 1 Then
                            strNameCondition = ""
                        ElseIf strEmp.Length < 2 Then
                            strFirstName = strEmp(0)
                            strNameCondition = " where firstName='" & strFirstName.Trim() & "'"
                        ElseIf strEmp.Length < 3 Then
                            strFirstName = strEmp(0)
                            strlatName = strEmp(1)
                            strNameCondition = " where firstName='" & strFirstName.Trim() & "' and lastName='" & strlatName.Trim() & "'"
                        ElseIf strEmp.Length < 4 Then
                            strFirstName = strEmp(0)
                            strMiddleName = strEmp(2)
                            strlatName = strEmp(1)
                            strNameCondition = " where firstName='" & strFirstName.Trim() & "' and lastName='" & strlatName.Trim() & "' and MiddleName='" & strMiddleName.Trim() & "'"
                        Else
                            strFirstName = strEmp(0)
                            strMiddleName = strEmp(2)
                            strlatName = strEmp(1)
                            strNameCondition = " where firstName='" & strFirstName.Trim() & "' and lastName='" & strlatName.Trim() & "' and MiddleName='" & strMiddleName.Trim() & "'"
                        End If
                        If strNameCondition <> "" Then
                            oTes.DoQuery("SELECT empID FROM OHEM T0  " & strNameCondition)
                            If oTes.RecordCount > 0 Then
                                oBP.DefaultTechnician = oTes.Fields.Item(0).Value
                            End If
                        End If
                        'oBP.DefaultTechnician = oDt.Rows(index)("Staff Name")

                        Dim blnExist As Boolean = False
                        Dim blnFamilyContact As Boolean = False
                        For intContact As Integer = 0 To oBP.ContactEmployees.Count - 1
                            oBP.ContactEmployees.SetCurrentLine(intContact)
                            If oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number").Trim() Then
                                blnExist = True
                                oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number").ToString.Replace("""", "") 'Person ID
                                oBP.ContactEmployees.Position = oDT.Rows(index)("Home Room").ToString.Replace("""", "")
                                oBP.ContactEmployees.FirstName = oDT.Rows(index)("First Name").ToString.Replace("""", "") 'Contact FirstName
                                oBP.ContactEmployees.MiddleName = oDT.Rows(index)("Middle Name").ToString.Replace("""", "")
                                oBP.ContactEmployees.LastName = oDT.Rows(index)("Last Name").ToString.Replace("""", "")
                                oBP.ContactEmployees.Remarks1 = oDT.Rows(index)("Grade Level")
                                oBP.ContactEmployees.CityOfBirth = oDT.Rows(index)("Ethnicity").ToString.Replace("""", "")
                                strPassword = oDT.Rows(index)("Ssn")
                                'If strPassword.Length > 7 Then
                                '    strPassword = strPassword.Substring(0, 8)
                                'Else
                                '    strPassword = strPassword
                                'End If
                                'oBP.ContactEmployees.Password = strPassword ' oDt.Rows(index)("Ssn")
                                oBP.ContactEmployees.Title = strPassword
                                If oDT.Rows(index)("Gender") = "M" Then
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Male
                                Else
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Female
                                End If

                                oBP.ContactEmployees.DateOfBirth = oDT.Rows(index)("Dob")
                                If oDT.Rows(index)("Paying").ToString = "F" Then
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                                Else
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                                End If
                            End If
                            If oDT.Rows(index)("Paying").ToString = "F" Then
                                If oBP.ContactEmployees.Name = strFamilyContName.Trim Then
                                    blnFamilyContact = True
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                                End If
                            ElseIf oDT.Rows(index)("Paying").ToString = "M" Then
                                If oBP.ContactEmployees.Name = strFamilyContName.Trim Then
                                    blnFamilyContact = True
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                                End If
                                '  MsgBox(oBP.ContactEmployees.Name)
                            End If

                        Next

                        If Not blnExist Then
                            If oBP.ContactEmployees.Count > 0 Then
                                oBP.ContactEmployees.Add()
                                oBP.ContactEmployees.SetCurrentLine(oBP.ContactEmployees.Count - 1)
                                oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number").ToString.Replace("""", "") 'Person ID
                                oBP.ContactEmployees.Position = oDT.Rows(index)("Home Room").ToString.Replace("""", "")
                                oBP.ContactEmployees.FirstName = oDT.Rows(index)("First Name").ToString.Replace("""", "") 'Contact FirstName
                                oBP.ContactEmployees.MiddleName = oDT.Rows(index)("Middle Name").ToString.Replace("""", "")
                                oBP.ContactEmployees.LastName = oDT.Rows(index)("Last Name").ToString.Replace("""", "")
                                oBP.ContactEmployees.Remarks1 = oDT.Rows(index)("Grade Level").ToString.Replace("""", "")

                                oBP.ContactEmployees.CityOfBirth = oDT.Rows(index)("Ethnicity").ToString.Replace("""", "") ' oBP.ContactEmployees.Title = oDT.Rows(index)("Ethnicity")
                                'oBP.ContactEmployees.Password = oDt.Rows(index)("Ssn")
                                strPassword = oDT.Rows(index)("Ssn")
                                'If strPassword.Length > 7 Then
                                '    strPassword = strPassword.Substring(0, 8)
                                'Else
                                '    strPassword = strPassword
                                'End If
                                ' oBP.ContactEmployees.Password = strPassword '
                                oBP.ContactEmployees.Title = strPassword
                                If oDT.Rows(index)("Gender") = "M" Then
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Male
                                Else
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Female
                                End If

                                oBP.ContactEmployees.DateOfBirth = oDT.Rows(index)("Dob")
                                If oDT.Rows(index)("Paying").ToString = "F" Then
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                                Else
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                                End If


                            Else
                                oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number").ToString.Replace("""", "") 'Person ID
                                oBP.ContactEmployees.Position = oDT.Rows(index)("Home Room").ToString.Replace("""", "")
                                oBP.ContactEmployees.FirstName = oDT.Rows(index)("First Name").ToString.Replace("""", "") 'Contact FirstName
                                oBP.ContactEmployees.MiddleName = oDT.Rows(index)("Middle Name").ToString.Replace("""", "")
                                oBP.ContactEmployees.LastName = oDT.Rows(index)("Last Name").ToString.Replace("""", "")
                                oBP.ContactEmployees.Remarks1 = oDT.Rows(index)("Grade Level").ToString.Replace("""", "")
                                oBP.ContactEmployees.Title = oDT.Rows(index)("Ethnicity").ToString.Replace("""", "")
                                'oBP.ContactEmployees.Password = oDt.Rows(index)("Ssn")
                                strPassword = oDT.Rows(index)("Ssn")
                                If strPassword.Length > 7 Then
                                    strPassword = strPassword.Substring(0, 8)
                                Else
                                    strPassword = strPassword
                                End If
                                oBP.ContactEmployees.Password = strPassword '
                                If oDT.Rows(index)("Gender") = "M" Then
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Male
                                Else
                                    oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Female
                                End If

                                oBP.ContactEmployees.DateOfBirth = oDT.Rows(index)("Dob")
                                If oDT.Rows(index)("Paying").ToString = "F" Then
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                                Else
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                                End If
                            End If
                        End If

                        If blnFamilyContact = False Then
                            If oDT.Rows(index)("Paying").ToString = "F" Then
                                If 1 = 1 Then
                                    oBP.ContactEmployees.Add()
                                    oBP.ContactEmployees.SetCurrentLine(oBP.ContactEmployees.Count - 1)
                                    oBP.ContactEmployees.Name = strFamilyContName.ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                                End If
                            ElseIf oDT.Rows(index)("Paying").ToString = "M" Then
                                If 1 = 1 Then
                                    oBP.ContactEmployees.Add()
                                    oBP.ContactEmployees.SetCurrentLine(oBP.ContactEmployees.Count - 1)
                                    oBP.ContactEmployees.Name = strFamilyContName.ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                                End If
                            End If
                        End If

                        If oDT.Rows(index)("Family Rep").ToString() = "1" Then
                            oBP.ContactPerson = oDT.Rows(index)("Student Number")
                        End If

                        Dim intStatus As Integer = oBP.Update()

                        If intStatus <> 0 Then
                            WriteErrorlog(strCardCode & "-" & strCustomer & "-" & " Error While Update : " & oCompany.GetLastErrorDescription(), strErrorLogPath)
                            '   traceService(strCardCode & "-" & strCustomer & "-" & " Error While Update : " & oCompany.GetLastErrorDescription())
                        Else
                            WriteErrorlog(strCardCode & "-" & strCustomer & "-" & "Updated Successfully", strErrorLogPath)
                        End If

                    Else

                        oBP.CardCode = strCardCode

                        If oDT.Rows(index)("Paying").ToString = "F" Then
                            oBP.CardName = oDT.Rows(index)("Father").ToString.Replace("""", "")
                            strCustomer = oDT.Rows(index)("Father").ToString.Replace("""", "")
                            strFamilyContName = oDT.Rows(index)("Mother").ToString.Replace("""", "")
                            oBP.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                            oBP.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                            oBP.Cellular = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                            oBP.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                        Else
                            oBP.CardName = oDT.Rows(index)("Mother").ToString.Replace("""", "")
                            strCustomer = oDT.Rows(index)("Mother").ToString.Replace("""", "")
                            strFamilyContName = oDT.Rows(index)("Father").ToString.Replace("""", "")
                            oBP.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                            oBP.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                            oBP.Cellular = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                            oBP.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                        End If

                        oBP.Addresses.AddressName = "Home Address"
                        oBP.Addresses.AddressName2 = strCustomer
                        ' oBP.Addresses.AddressName = strCustomer.ToString.Replace("""", "")
                        oBP.Addresses.Street = oDT.Rows(index)("Street").ToString.Replace("""", "")
                        oBP.Addresses.ZipCode = oDT.Rows(index)("Zip").ToString.Replace("""", "")
                        oBP.Addresses.City = oDT.Rows(index)("City").ToString.Replace("""", "")
                        'oBP.Addresses.Country = oDt.Rows(index)("State Country")
                        oBP.Addresses.Country = "BH"
                        oBP.Addresses.GlobalLocationNumber = oDT.Rows(index)("Mailing City").ToString.Replace("""", "")
                        oBP.Addresses.BuildingFloorRoom = oDT.Rows(index)("Mailing Street").ToString.Replace("""", "")
                        oBP.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo

                        oBP.EmailAddress = oDT.Rows(index)("Guardianemail")
                        oBP.Website = oDT.Rows(index)("Special Conditions Custom").ToString.Replace("""", "")
                        oBP.Password = oDT.Rows(index)("Custody Role").ToString.Replace("""", "")
                        oBP.AdditionalID = oDT.Rows(index)("Pick Up").ToString.Replace("""", "")
                        oBP.CompanyRegistrationNumber = oDT.Rows(index)("Pick Up Contact Mobile")
                        Dim oTes As SAPbobsCOM.Recordset
                        oTes = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oTes.DoQuery("SELECT T0.[IndCode], T0.[IndName], T0.[IndDesc] FROM OOND T0 where IndName='" & oDT.Rows(index)("Section") & "'")
                        If oTes.RecordCount > 0 Then
                            oBP.Industry = oTes.Fields.Item(0).Value
                        End If
                        'oBP.Industry = oDt.Rows(index)("Section")
                        Dim strEmp As String()
                        Dim strName As String = oDT.Rows(index)("Staff Name")
                        strEmp = strName.Split(",")
                        Dim strFirstName, strMiddleName, strlatName, strNameCondition As String
                        If strEmp.Length < 1 Then
                            strNameCondition = ""
                        ElseIf strEmp.Length < 2 Then
                            strFirstName = strEmp(0)
                            strNameCondition = " where firstName='" & strFirstName.Trim() & "'"
                        ElseIf strEmp.Length < 3 Then
                            strFirstName = strEmp(0)
                            strlatName = strEmp(1)
                            strNameCondition = " where firstName='" & strFirstName.Trim() & "' and lastName='" & strlatName.Trim() & "'"
                        ElseIf strEmp.Length < 4 Then
                            strFirstName = strEmp(0)
                            strMiddleName = strEmp(2)
                            strlatName = strEmp(1)
                            strNameCondition = " where firstName='" & strFirstName.Trim() & "' and lastName='" & strlatName.Trim() & "' and MiddleName='" & strMiddleName.Trim() & "'"
                        Else
                            strFirstName = strEmp(0)
                            strMiddleName = strEmp(2)
                            strlatName = strEmp(1)
                            strNameCondition = " where firstName='" & strFirstName.Trim() & "' and lastName='" & strlatName.Trim() & "' and MiddleName='" & strMiddleName.Trim() & "'"
                        End If
                        If strNameCondition <> "" Then
                            oTes.DoQuery("SELECT empID FROM OHEM T0  " & strNameCondition)
                            If oTes.RecordCount > 0 Then
                                oBP.DefaultTechnician = oTes.Fields.Item(0).Value
                            End If
                        End If
                        oBP.ContactEmployees.Name = oDT.Rows(index)("Student Number").ToString.Replace("""", "") 'Person ID
                        oBP.ContactEmployees.Position = oDT.Rows(index)("Home Room").ToString.Replace("""", "")
                        oBP.ContactEmployees.FirstName = oDT.Rows(index)("First Name").ToString.Replace("""", "") 'Contact FirstName
                        oBP.ContactEmployees.MiddleName = oDT.Rows(index)("Middle Name").ToString.Replace("""", "")
                        oBP.ContactEmployees.LastName = oDT.Rows(index)("Last Name").ToString.Replace("""", "")
                        oBP.ContactEmployees.Remarks1 = oDT.Rows(index)("Grade Level").ToString.Replace("""", "")
                        oBP.ContactEmployees.CityOfBirth = oDT.Rows(index)("Ethnicity").ToString.Replace("""", "")
                        'oBP.ContactEmployees.Password = oDt.Rows(index)("Ssn")
                        strPassword = oDT.Rows(index)("Ssn")
                        If strPassword.Length > 7 Then
                            strPassword = strPassword.Substring(0, 8)
                        Else
                            strPassword = strPassword
                        End If
                        oBP.ContactEmployees.Title = strPassword '
                        If oDT.Rows(index)("Gender") = "M" Then
                            oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Male
                        Else
                            oBP.ContactEmployees.Gender = SAPbobsCOM.BoGenderTypes.gt_Female
                        End If

                        oBP.ContactEmployees.DateOfBirth = oDT.Rows(index)("Dob")
                        If oDT.Rows(index)("Paying").ToString = "F" Then
                            oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                            oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                            oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                            oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                        Else
                            oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                            oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                            oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                            oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                        End If

                        If oDT.Rows(index)("Family Rep").ToString() = "1" Then
                            oBP.ContactPerson = oDT.Rows(index)("Student Number").ToString.Replace("""", "")
                        End If

                        If 1 = 1 Then
                            If oDT.Rows(index)("Paying").ToString = "F" Then
                                If 1 = 1 Then
                                    oBP.ContactEmployees.Add()
                                    oBP.ContactEmployees.SetCurrentLine(1)
                                    oBP.ContactEmployees.Name = strFamilyContName.ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Motherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Mother Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Mother Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Mother Cell Phone 2").ToString.Replace("""", "")
                                End If
                            ElseIf oDT.Rows(index)("Paying").ToString = "M" Then
                                If 1 = 1 Then
                                    oBP.ContactEmployees.Add()
                                    oBP.ContactEmployees.SetCurrentLine(1)
                                    oBP.ContactEmployees.Name = strFamilyContName.ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone1 = oDT.Rows(index)("Fatherdayphone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Phone2 = oDT.Rows(index)("Father Home Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.MobilePhone = oDT.Rows(index)("Father Cell Phone").ToString.Replace("""", "")
                                    oBP.ContactEmployees.Fax = oDT.Rows(index)("Father Cell Phone 2").ToString.Replace("""", "")
                                End If
                            End If
                        End If

                        Dim intStatus As Integer = oBP.Add()
                        If intStatus <> 0 Then
                            '  traceService(strCardCode & "-" & strCustomer & "-" & " Error While Add : " & oCompany.GetLastErrorDescription())
                            WriteErrorlog(strCardCode & "-" & strCustomer & "-" & " Error While Update : " & oCompany.GetLastErrorDescription(), strErrorLogPath)
                            '   traceService(strCardCode & "-" & strCustomer & "-" & " Error While Update : " & oCompany.GetLastErrorDescription())
                        Else
                            WriteErrorlog(strCardCode & "-" & strCustomer & "-" & "Created Successfully", strErrorLogPath)
                        End If

                    End If
                Next
            End If



        Catch ex As Exception
            traceService(ex.ToString)
        End Try


    End Sub

    Private Sub connectSAP1()
        Try
            traceService("Company Connection")
            oCompany = objRemoteCompany ' New SAPbobsCOM.Company()
            'Dim strCompany As String = System.Configuration.ConfigurationManager.AppSettings("CompanyDB").ToString
            'oCompany.CompanyDB = strCompany

            'oCompany.UserName = System.Configuration.ConfigurationManager.AppSettings("UserName").ToString
            'oCompany.Password = System.Configuration.ConfigurationManager.AppSettings("Password").ToString
            'oCompany.Server = System.Configuration.ConfigurationManager.AppSettings("Server").ToString
            'oCompany.DbUserName = System.Configuration.ConfigurationManager.AppSettings("DbUserName").ToString
            'oCompany.DbPassword = System.Configuration.ConfigurationManager.AppSettings("DbPassword").ToString

            'If System.Configuration.ConfigurationManager.AppSettings("DBType").ToString = "2008" Then
            '    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            'ElseIf System.Configuration.ConfigurationManager.AppSettings("DBType").ToString = "2012" Then
            '    'oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            'End If

            'oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            'oCompany.LicenseServer = System.Configuration.ConfigurationManager.AppSettings("LicenseServer").ToString

            'If (0 <> oCompany.Connect()) Then
            '    traceService("Company Connection Failed")
            '    traceService(oCompany.GetLastErrorDescription())
            '    Exit Sub
            'Else
            '    traceService("Connected successfully")
            'End If
        Catch ex As Exception
            ' traceService(ex.ToString())
        End Try
    End Sub

    'Private Sub disconnectSAP()
    '    Try
    '        If oCompany.Connected Then
    '            oCompany.Disconnect()
    '        End If
    '    Catch ex As Exception
    '        traceService(ex.ToString())
    '    End Try
    'End Sub

    Private Sub MoveToFolder(ByVal strExFile As String)
        Try
            Dim OpenFolder As String
            Dim SuceessFolder As String
            OpenFolder = strExportFilePaty ' System.Configuration.ConfigurationManager.AppSettings("OpenFolder").ToString
            SuceessFolder = strSuccessFolder ' System.Configuration.ConfigurationManager.AppSettings("SuccessFolder").ToString
            Dim strFileName As String = Path.GetFileName(strExFile)
            Dim strDesgFile As String
            If System.IO.File.Exists(strExFile) = True Then
                strDesgFile = SuceessFolder + "\" + strFileName
                If File.Exists(strDesgFile) Then
                    File.Delete(strDesgFile)

                End If
                File.Move(strExFile, strDesgFile)
                'System.IO.File.Move(strExFile, SuceessFolder + "\" + strFileName)
                '  File.Copy(strExFile, SuceessFolder + "\" + strFileName, True)
                '  File.Delete(strExFile)
                traceService("File : " & strFileName & " File Moved Successfully")
            End If
        Catch ex As Exception
            traceService(ex.ToString)
        End Try
    End Sub

    Private Sub traceService(ByVal content As String)
        WriteErrorlog(content, strErrorLogPath)
        Exit Sub
        Try


            Dim strFile As String = strErrorLogPath ' "\BAYAN_SERVICE_" + System.DateTime.Now.ToString("yyyyMMdd") + ".txt"
            Dim strPath As String = strErrorLogPath ' System.Windows.Forms.Application.StartupPath.ToString() & strFile

            If Not File.Exists(strPath) Then
                Dim fs As New FileStream(strPath, FileMode.Create, FileAccess.Write)
                Dim sw As New StreamWriter(fs)
                sw.BaseStream.Seek(0, SeekOrigin.[End])
                sw.WriteLine(content)
                sw.Flush()
                sw.Close()
            Else
                Dim fs As New FileStream(strPath, FileMode.Append, FileAccess.Write)
                Dim sw As New StreamWriter(fs)
                sw.BaseStream.Seek(0, SeekOrigin.[End])
                sw.WriteLine(content)
                sw.Flush()
                sw.Close()
            End If
        Catch ex As Exception
            'traceService(ex.ToString)
        End Try
    End Sub

#End Region

#Region "API Calls"

    '****************************************************************************
    'Type              :   Function     
    'Name              :   WritePrivateProfileString
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   Standard API declarations for INI access
    '****************************************************************************

    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Int32

    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Int32, _
    ByVal lpFileName As String) As Int32
#End Region

#Region "Function for Read INI File"

    '****************************************************************************
    'Type              :   Function    
    'Name              :   INIRead
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Read INI File
    '****************************************************************************

    Public Overloads Function INIRead(ByVal INIPath As String, _
        ByVal SectionName As String, ByVal KeyName As String, _
        ByVal DefaultValue As String) As String
        Dim n As Int32
        Dim sData As String
        sData = Space$(1024)
        n = GetPrivateProfileString(SectionName, KeyName, DefaultValue, _
        sData, sData.Length, INIPath)
        If n > 0 Then
            INIRead = sData.Substring(0, n)
        Else
            INIRead = ""
        End If
    End Function
#End Region

#Region "To Read INI File"

    '****************************************************************************
    'Type              :   Function     
    'Name              :   INIRead
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Read INI File
    '****************************************************************************

    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String, ByVal KeyName As String) As String
        Return INIRead(INIPath, SectionName, KeyName, "")
    End Function

    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String) As String
        ' overload 2 returns all keys in a given section of the given file 
        Return INIRead(INIPath, SectionName, Nothing, "")
    End Function

    Public Overloads Function INIRead(ByVal INIPath As String) As String
        ' overload 3 returns all section names given just path 
        Return INIRead(INIPath, Nothing, Nothing, "")
    End Function
#End Region

#Region "To write INI File"

    '****************************************************************************
    'Type              :   Procedure     
    'Name              :   INIWrite
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To write INI File 
    '****************************************************************************

    Public Sub INIWrite(ByVal INIPath As String, ByVal SectionName As String, _
       ByVal KeyName As String, ByVal TheValue As String)
        Call WritePrivateProfileString(SectionName, KeyName, TheValue, INIPath)
    End Sub

    Public Overloads Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String, _
    ByVal KeyName As String) ' delete single line from section 
        Call WritePrivateProfileString(SectionName, KeyName, Nothing, INIPath)
    End Sub

    Public Overloads Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String)
        ' delete section from INI file 
        Call WritePrivateProfileString(SectionName, Nothing, Nothing, INIPath)
    End Sub


#End Region

#Region "To Save Details in ini File"
    '****************************************************************************
    'Type              :   Procedure     
    'Name              :   WriteiniFile
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Save Details in ini File
    '****************************************************************************
    Private Sub WriteiniFile()
        sPath = System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.ini"
        INIWrite(sPath, "SqlServer", "Key1-1", txtServerName.Text) ' build INI file 
        INIWrite(sPath, "SQLUID", "Key1-2", txtDBUserName.Text)
        INIWrite(sPath, "SQLPWD", "Key1-3", txtServerPwd.Text)
        INIWrite(sPath, "ServerType", "Key1-S1", strLocalServertype)
        INIWrite(sPath, "SAPCompany", "Key1-4", strSAPCompany) ' cmbCompanyDB.SelectedText)
        INIWrite(sPath, "SAPUID", "Key1-5", txtSBOUserName.Text)
        INIWrite(sPath, "SAPPWD", "Key1-6", txtCompanyPWD.Text)
        INIWrite(sPath, "LogFile", "Key1-9", txtDirectory.Text)

        INIWrite(sPath, "MainSqlServer", "Key1-10", MainSQLServer.Text) ' build INI file 
        INIWrite(sPath, "MainSQLUID", "Key1-11", MainSQLUID.Text)
        INIWrite(sPath, "MainSQLPWD", "Key1-12", MainSQLPWD.Text)
        INIWrite(sPath, "MainServerType", "Key1-M1", strMainServertype)
        INIWrite(sPath, "MainSAPCompany", "Key1-13", strMainSAPCompany) ' cmbCompanyDB.SelectedText)
        INIWrite(sPath, "MainSAPUID", "Key1-14", MainSAPUID.Text)
        INIWrite(sPath, "MainSAPPWD", "Key1-15", MainSAPPWD.Text)
        INIWrite(sPath, "MainLicense", "Key1-16", mainLicenseServer.Text)
        INIWrite(sPath, "LocalLicense", "Key1-17", LocalLicenseServer.Text)


        INIWrite(sPath, "InterfaceSQLServer", "Key1-27", interfaceServer)
        INIWrite(sPath, "InterfaceSQLUID", "Key1-28", interfaceUID)
        INIWrite(sPath, "InterfaceSQLPWD", "Key1-29", interfacePwd)
        INIWrite(sPath, "InterfaceDB", "Key1-30", interfaceDB)
        INIWrite(sPath, "Vat", "Key1-31", ImpVAT)
        INIWrite(sPath, "NonVat", "Key1-32", ImpNonVAT)

    End Sub

#End Region

#Region "To Read ini File"

    '****************************************************************************
    'Type              :   Procedure     
    'Name              :   ReadiniFile
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Read ini File
    '****************************************************************************

    Private Sub ReadiniFile()
        sPath = System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.ini"
        sValue = INIRead(sPath, "SqlServer", "key1-1", "")
        txtServerName.Text = sValue
        strSQLServer = sValue
        sValue = INIRead(sPath, "SQLUID", "Key1-2", "")
        txtDBUserName.Text = sValue
        strSQLUID = sValue
        sValue = INIRead(sPath, "SQLPWD", "Key1-3", "")
        txtServerPwd.Text = sValue
        strSQLPWD = sValue
        sValue = INIRead(sPath, "ServerType", "Key1-S1", "")
        cmbServertype.Text = sValue
        strLocalServertype = sValue
        sValue = INIRead(sPath, "SAPCompany", "Key1-4", "")
        cmbCompanyDB.SelectedText = sValue
        strSAPCompany = sValue
        sValue = INIRead(sPath, "SAPUID", "Key1-5", "")
        txtSBOUserName.Text = sValue
        strSAPUID = sValue
        sValue = INIRead(sPath, "SAPPWD", "Key1-6", "")
        txtCompanyPWD.Text = sValue
        strSAPPWD = sValue
        strSAPPWD = sValue
        sValue = INIRead(sPath, "ImportDir", "Key1-7", "")
        '  txtDirectory.Text = sValue
        strExportFilePaty = sValue
        'sValue = INIRead(sPath, "ExportFile", "Key1-8", "")
        'strFileStart = sValue
        sValue = INIRead(sPath, "LogFile", "Key1-9", "")
        strErrorLogPath = sValue
        txtDirectory.Text = strErrorLogPath


        sValue = INIRead(sPath, "SussessDir", "key1-8", "")
        strSuccessFolder = sValue

        sValue = INIRead(sPath, "InterfaceSQLServer", "key1-27", "")
        interfaceServer = sValue

        sValue = INIRead(sPath, "InterfaceSQLUID", "key1-28", "")
        interfaceUID = sValue

        sValue = INIRead(sPath, "InterfaceSQLPWD", "key1-29", "")
        interfacePwd = sValue

        sValue = INIRead(sPath, "InterfaceDB", "key1-30", "")
        interfaceDB = sValue

        sValue = INIRead(sPath, "Vat", "key1-31", "")
        ImpVAT = sValue

        sValue = INIRead(sPath, "NonVat", "key1-32", "")
        ImpNonVAT = sValue
        'sValue = INIRead(sPath, "MainSQLUID", "Key1-11", "")
        'MainSQLUID.Text = sValue
        'strMainSQLUID = sValue
        'sValue = INIRead(sPath, "MainSQLPWD", "Key1-12", "")
        'MainSQLPWD.Text = sValue
        'strMainSQLPWD = sValue
        'sValue = INIRead(sPath, "MainServerType", "Key1-M1", "")
        'cmbMainServertype.Text = sValue
        'strMainServertype = sValue

        'sValue = INIRead(sPath, "MainSAPCompany", "Key1-13", "")
        'cmbMainSAPCompany.SelectedText = sValue
        'strMainSAPCompany = sValue
        'sValue = INIRead(sPath, "MainSAPUID", "Key1-14", "")
        'MainSAPUID.Text = sValue
        'strMainSAPUID = sValue
        'sValue = INIRead(sPath, "MainSAPPWD", "Key1-15", "")
        'MainSAPPWD.Text = sValue
        'strMainSAPPWD = sValue
        'strMainSAPPWD = sValue

        'sValue = INIRead(sPath, "MainLicense", "Key1-16", "")
        'mainLicenseServer.Text = sValue
        'strMainLicenseServer = sValue

        sValue = INIRead(sPath, "LocalLicense", "Key1-17", "")
        LocalLicenseServer.Text = sValue
        strLocalLicensserver = sValue


    End Sub

#End Region

#Region "Get DBInformation from XML"

    '****************************************************************************
    'Type              :   Function     
    'Name              :   GetConInfo
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Get DBInformation from XML
    '****************************************************************************

    Public Function GetConInfo() As String
        Dim objXMLDoc As New System.Xml.XmlDocument
        Dim objNode As Xml.XmlNode
        Dim intLoop As Integer
        Dim strConn, strserver, strDb, strUID, strPwd As String
        objXMLDoc.Load(System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.xml")
        With objXMLDoc.GetElementsByTagName("Parameter")
            For intLoop = 0 To .Count - 1
                objNode = .Item(intLoop)
                If objNode.Attributes("Type").Value = "SBOServer" Then
                    strserver = objNode.Attributes("Value").Value
                ElseIf objNode.Attributes("Type").Value = "Database" Then
                    strDb = objNode.Attributes("Value").Value
                ElseIf objNode.Attributes("Type").Value = "DBUser" Then
                    strUID = objNode.Attributes("Value").Value
                ElseIf objNode.Attributes("Type").Value = "DBPass" Then
                    strPwd = objNode.Attributes("Value").Value
                End If
            Next
        End With
        strConn = "Database=" & strDb & ";Server=" & strserver & ";uid=" & strUID & ";Pwd=" & strPwd & ";"
        Return strConn
    End Function


#End Region

#Region "WMS Interface"

    Private Function getConnection() As String
        Dim strConnectionstring As String
        strConnectionstring = "Server=" & interfaceServer & ";Database=" & interfaceDB & ";User ID=" & interfaceUID & ";Password=" & interfacePwd
        Return strConnectionstring
    End Function

    Private Sub mainFunction()
        Try
            connectSAP()

            'importInvoice()
            'importCreditMemo()
            'importIncomingPayment()

            import_Delivery()
            import_Return()
            import_Purchase()
            import_Transfer()
            import_InventoryReceipt()
            import_InventoryIssue()
            import_InventoryCounting()

            disconnectSAP()
        Catch ex As Exception
            'traceService(ex.ToString())
        Finally
            disconnectSAP()
        End Try
    End Sub

    Public Function ExecuteDataSet(ByVal strQuery As String) As DataSet
        Dim _retVal As New DataSet()
        Dim ConnectionString As String = getConnection() ' System.Configuration.ConfigurationManager.AppSettings("ConnectionString")
        myConnection = New SqlConnection(ConnectionString)
        Try
            myConnection.Open()
            oCommand = New SqlCommand()
            If myConnection.State = ConnectionState.Open Then
                oCommand.Connection = myConnection
                oCommand.CommandText = strQuery
                oCommand.CommandType = CommandType.Text
                oSqlAdap = New SqlDataAdapter(oCommand)
                oSqlAdap.Fill(_retVal)
            Else
                Throw New Exception("")
            End If
        Catch ex As Exception
            Throw ex
        Finally
            myConnection.Close()
            oCommand = Nothing
            oSqlAdap = Nothing
        End Try
        Return _retVal
    End Function

    Public Function ValidateInterFaceDBConnection() As Boolean
        Dim _retVal As New DataSet()
        Dim ConnectionString As String = getConnection() ' System.Configuration.ConfigurationManager.AppSettings("ConnectionString")
        myConnection = New SqlConnection(ConnectionString)
        Try
            myConnection.Open()
        Catch ex As Exception
            WriteErrorlog("Connection to Intergration Database is failed : " & ex.Message, strErrorFileName)
            Return False
        Finally
            myConnection.Close()
            oCommand = Nothing
            oSqlAdap = Nothing
        End Try
        Return True
    End Function

    Public Sub ExecuteNonQuery(ByVal str As String)
        Dim ConnectionString As String = getConnection() ' System.Configuration.ConfigurationManager.AppSettings("ConnectionString")
        Dim oCommand As SqlCommand
        myConnection = New SqlConnection(ConnectionString)

        Try
            myConnection.Open()
            oCommand = New SqlCommand(str, myConnection)
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            myConnection.Close()
        End Try
    End Sub

    Private Function getCustomerCode(ByVal aCode As String) As String
        Return aCode
        Dim oBPRec As SAPbobsCOM.Recordset
        oBPRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oBPRec.DoQuery("Select CardCode from OCRD where U_Code='" & aCode & "'")
        If oBPRec.RecordCount > 0 Then
            Return oBPRec.Fields.Item(0).Value
        Else
            Return ""

        End If

    End Function

    'Private Sub importInvoice()
    '    Try
    '        Dim oInvoice As SAPbobsCOM.Documents
    '        Dim ds As DataSet
    '        Dim strQuery As String

    '        Dim VAT As String = ImpVAT ' System.Configuration.ConfigurationManager.AppSettings("VAT")
    '        Dim NONVAT As String = ImpNonVAT  ' System.Configuration.ConfigurationManager.AppSettings("NONVAT")
    '        Dim strLocalCurrncy As String = getLocalCurrency()
    '        strQuery = " Select T0.IssueDate,T0.DeliveryDate, T0.ClientCode, T0.CurrencyCode, T0.TransactionCode,  T0.DiscountAmount, T0.Info2" & _
    '                    " From TransactionHeaders T0 " & _
    '                    " Where T0.TransactionType = 2 And T0.IsProcessed=2 "
    '        ds = ExecuteDataSet(strQuery)
    '        WriteErrorlog("Processing Invoice Creation ", strErrorFileName)
    '        Dim strCardCode As String
    '        If Not IsNothing(ds) Then
    '            oHeaderDt = ds.Tables(0)

    '            If oHeaderDt.Rows.Count > 0 Then
    '                For Each Hrow As DataRow In oHeaderDt.Rows
    '                    oInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
    '                    TransactionCode = Hrow("TransactionCode").ToString().Trim()
    '                    WriteErrorlog("Processing Transaction Code :  " & TransactionCode, strErrorFileName)
    '                    strCardCode = ""
    '                    strCardCode = getCustomerCode(Hrow("ClientCode").ToString().Trim())
    '                    If strCardCode = "" Then
    '                        WriteErrorlog("Customer Code does not exists : " & (Hrow("ClientCode").ToString().Trim()), strErrorFileName)

    '                    Else
    '                        oInvoice.CardCode = strCardCode ' getCustomerCode(Hrow("ClientCode").ToString().Trim()) ' Hrow("ClientCode").ToString().Trim()
    '                        oInvoice.DocDate = Hrow("IssueDate")
    '                        Dim dtDueDate As Date = getDueDate(strCardCode, Hrow("IssueDate"))
    '                        oInvoice.DocDueDate = dtDueDate

    '                        If strLocalCurrncy <> Hrow("CurrencyCode").ToString().Trim() Then
    '                            oInvoice.DocCurrency = Hrow("CurrencyCode").ToString().Trim()
    '                            oInvoice.DocRate = getCurrencyRate(Hrow("CurrencyCode").ToString().Trim(), oInvoice.DocDate)
    '                        End If

    '                        oInvoice.NumAtCard = TransactionCode.ToString

    '                        strQuery = "Select T1.ItemCode, T1.ItemName, T1.BasicUnitQuantity, T1.PriceSmall, T1.DiscountAmount, T1.ItemTaxPercentage, T1.DiscountPercentage" & _
    '                                    " From TransactionDetails T1 " & _
    '                                    " Where T1.TransactionCode= '" & TransactionCode & "'"
    '                        ds = ExecuteDataSet(strQuery)

    '                        If Not IsNothing(ds) Then
    '                            oDetailDt = ds.Tables(0)
    '                            If oDetailDt.Rows.Count > 0 Then
    '                                Dim intRow As Integer = 0
    '                                For Each Drow As DataRow In oDetailDt.Rows
    '                                    If intRow > 0 Then
    '                                        oInvoice.Lines.Add()
    '                                    End If
    '                                    oInvoice.Lines.SetCurrentLine(intRow)
    '                                    oInvoice.Lines.ItemCode = Drow("ItemCode").ToString().Trim()
    '                                    oInvoice.Lines.ItemDescription = Drow("ItemName").ToString().Trim()
    '                                    oInvoice.Lines.Quantity = Drow("BasicUnitQuantity")
    '                                    oInvoice.Lines.UnitPrice = Drow("PriceSmall")
    '                                    oInvoice.Lines.DiscountPercent = Drow("DiscountPercentage")
    '                                    If Drow("ItemTaxPercentage") <> 0 Then
    '                                        oInvoice.Lines.TaxCode = VAT
    '                                    Else
    '                                        oInvoice.Lines.TaxCode = NONVAT
    '                                    End If
    '                                    oInvoice.Lines.WarehouseCode = Hrow("Info2").ToString().Trim()

    '                                    Dim oRecordset As SAPbobsCOM.Recordset
    '                                    oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                                    strQuery = "Select SalUnitMsr From OITM Where ItemCode = '" & Drow("ItemCode") & "'"
    '                                    oRecordset.DoQuery(strQuery)

    '                                    If Not oRecordset.EoF Then
    '                                        If oRecordset.Fields.Item(0).Value.ToString() <> "" Then
    '                                            oInvoice.Lines.MeasureUnit = oRecordset.Fields.Item(0).Value.ToString()
    '                                        End If
    '                                    End If
    '                                    intRow = intRow + 1
    '                                Next
    '                            End If
    '                        End If

    '                        Dim intStatus As Integer = oInvoice.Add()
    '                        If intStatus = 0 Then
    '                            Dim strDocNum As String
    '                            oCompany.GetNewObjectCode(strDocNum)
    '                            Dim oTest As SAPbobsCOM.Recordset
    '                            oTest = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                            oTest.DoQuery("Select * from OINV where DocEntry=" & strDocNum)
    '                            strDocNum = oTest.Fields.Item("DocNum").Value
    '                            strQuery = "Update TransactionHeaders" & _
    '                                       " Set IsProcessed = 3" & _
    '                                       " Where TransactionCode= '" & TransactionCode & "'"
    '                            ExecuteNonQuery(strQuery)
    '                            WriteErrorlog("Invoice : " & strDocNum & " Created for Transaction Code : " & TransactionCode, strErrorFileName)
    '                        Else
    '                            '  MessageBox.Show(oCompany.GetLastErrorDescription())
    '                            WriteErrorlog("Invoice Creation Failed for Transaction Code : " & TransactionCode & " : Error Details : " & oCompany.GetLastErrorDescription(), strErrorFileName)
    '                        End If
    '                    End If
    '                Next
    '            End If
    '        End If
    '        WriteErrorlog("Processing Invoice Creation Competed ", strErrorFileName)

    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Sub

    'Private Sub importCreditMemo()
    '    Try
    '        Dim oCreditMemo As SAPbobsCOM.Documents
    '        Dim ds As DataSet
    '        Dim strQuery As String
    '        Dim strLocalCurrncy As String = getLocalCurrency()

    '        strQuery = "Select T0.IssueDate, T0.ClientCode, T0.CurrencyCode, T0.TransactionCode, T1.ClientName, T0.Info2" & _
    '                      " From TransactionHeaders T0 Join Clients T1" & _
    '                      " On T0.ClientCode= T1.ClientCode" & _
    '                      " Where(T0.TransactionType = 3)"

    '        ds = ExecuteDataSet(strQuery)
    '        WriteErrorlog("Processing Credit Note Creation ", strErrorFileName)
    '        If Not IsNothing(ds) Then
    '            oHeaderDt = ds.Tables(0)
    '            If oHeaderDt.Rows.Count > 0 Then
    '                For Each Hrow As DataRow In oHeaderDt.Rows
    '                    oCreditMemo = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
    '                    TransactionCode = Hrow("TransactionCode").ToString().Trim()
    '                    WriteErrorlog("Processing Transaction Code :  " & TransactionCode, strErrorFileName)
    '                    If getCustomerCode(Hrow("ClientCode").ToString().Trim()) = "" Then
    '                        WriteErrorlog("Customer Code does not exists : " & (Hrow("ClientCode").ToString().Trim()), strErrorFileName)
    '                    Else
    '                        oCreditMemo.CardCode = getCustomerCode(Hrow("ClientCode").ToString().Trim()) ' Hrow("ClientCode").ToString().Trim()
    '                        '   oCreditMemo.CardCode = Hrow("ClientCode").ToString().Trim()
    '                        oCreditMemo.CardName = Hrow("ClientName").ToString().Trim()
    '                        oCreditMemo.DocDate = Hrow("IssueDate")
    '                        If strLocalCurrncy <> Hrow("CurrencyCode").ToString().Trim() Then
    '                            oCreditMemo.DocCurrency = Hrow("CurrencyCode").ToString().Trim()
    '                            oCreditMemo.DocRate = getCurrencyRate(Hrow("CurrencyCode").ToString().Trim(), oCreditMemo.DocDate)
    '                        End If
    '                        oCreditMemo.NumAtCard = TransactionCode.ToString
    '                        strQuery = "Select T1.ItemCode, T1.ItemName, T1.BasicUnitQuantity, T1.PriceSmall" & _
    '                                        " From TransactionDetails T1 " & _
    '                                        " Where T1.TransactionCode= '" & TransactionCode & "'"
    '                        ds = ExecuteDataSet(strQuery)
    '                        If Not IsNothing(ds) Then
    '                            oDetailDt = ds.Tables(0)
    '                            If oDetailDt.Rows.Count > 0 Then
    '                                Dim intRow As Integer = 0
    '                                For Each Drow As DataRow In oDetailDt.Rows
    '                                    If intRow > 0 Then
    '                                        oCreditMemo.Lines.Add()
    '                                    End If
    '                                    oCreditMemo.Lines.SetCurrentLine(intRow)
    '                                    oCreditMemo.Lines.ItemCode = Drow("ItemCode").ToString().Trim()
    '                                    oCreditMemo.Lines.ItemDescription = Drow("ItemName").ToString().Trim()
    '                                    oCreditMemo.Lines.Quantity = Drow("BasicUnitQuantity")
    '                                    oCreditMemo.Lines.UnitPrice = Drow("PriceSmall")
    '                                    oCreditMemo.Lines.WarehouseCode = Hrow("Info2").ToString().Trim()
    '                                    Dim oRecordset As SAPbobsCOM.Recordset
    '                                    oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                                    strQuery = "Select SalUnitMsr From OITM Where ItemCode = '" & Drow("ItemCode") & "'"
    '                                    oRecordset.DoQuery(strQuery)
    '                                    If Not oRecordset.EoF Then
    '                                        If oRecordset.Fields.Item(0).Value.ToString() <> "" Then
    '                                            oCreditMemo.Lines.MeasureUnit = oRecordset.Fields.Item(0).Value.ToString()
    '                                        End If
    '                                    End If
    '                                    intRow = intRow + 1
    '                                Next
    '                            End If
    '                        End If
    '                        Dim intStatus As Integer = oCreditMemo.Add()
    '                        If intStatus = 0 Then
    '                            'MessageBox"Added successfully")
    '                            strQuery = "Update TransactionHeaders" & _
    '                                     " Set IsProcessed = 3" & _
    '                                     " Where TransactionCode= '" & TransactionCode & "'"
    '                            ExecuteNonQuery(strQuery)
    '                            Dim strDocNum As String
    '                            oCompany.GetNewObjectCode(strDocNum)
    '                            Dim oTest As SAPbobsCOM.Recordset
    '                            oTest = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                            oTest.DoQuery("Select * from ORIN where DocEntry=" & strDocNum)
    '                            strDocNum = oTest.Fields.Item("DocNum").Value
    '                            WriteErrorlog("Credit Note : " & strDocNum & " Created for Transaction Code : " & TransactionCode, strErrorFileName)
    '                        Else
    '                            ' MessageBox.Show(oCompany.GetLastErrorDescription())
    '                            WriteErrorlog("Creating Credite Note  for Transaction Code : " & TransactionCode & " Failed. Error : " & oCompany.GetLastErrorDescription, strErrorFileName)
    '                        End If
    '                    End If

    '                Next
    '            End If
    '        End If
    '        WriteErrorlog("Processing Credit Note Creation Completed ", strErrorFileName)

    '    Catch ex As Exception
    '        'traceService(ex.ToString())
    '    End Try
    'End Sub

    'Private Sub importIncomingPayment()
    '    Try
    '        Dim oPayment As SAPbobsCOM.Payments

    '        Dim ds As DataSet
    '        Dim strQuery As String

    '        strQuery = "Select T0.CollectionCode, T0.ClientCode, T0.CollectionDate, T1.CollectionAmount, T1.CollectionDetailType," & _
    '                    " T1.CheckCode, T1.CheckDate, T1.BankCode, T1.BankName,T1.BankName ,T1.CurrencyCode,T1.CurrencyRate " & _
    '                    " From CollectionHeaders T0 Join CollectionDetails T1" & _
    '                    " On T0.CollectionCode = T1.CollectionCode" & _
    '                    " Where (T0.IsProcessed = 2) "
    '        ds = ExecuteDataSet(strQuery)
    '        WriteErrorlog("Processing Incoming Payment Creation ", strErrorFileName)
    '        If Not IsNothing(ds) Then

    '            oHeaderDt = ds.Tables(0)

    '            If oHeaderDt.Rows.Count > 0 Then
    '                For Each Hrow As DataRow In oHeaderDt.Rows
    '                    oPayment = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
    '                    TransactionCode = Hrow("CollectionCode")
    '                    WriteErrorlog("Processing Transaction Code :  " & TransactionCode, strErrorFileName)
    '                    Dim CollectionCode As Integer = Hrow("CollectionCode")
    '                    '  oPayment.CardCode = Hrow("ClientCode").ToString().Trim()
    '                    If getCustomerCode(Hrow("ClientCode").ToString().Trim()) = "" Then
    '                        WriteErrorlog("Customer Code does not exists : " & (Hrow("ClientCode").ToString().Trim()), strErrorFileName)
    '                    Else
    '                        oPayment.CardCode = getCustomerCode(Hrow("ClientCode").ToString.Trim())
    '                        oPayment.DocDate = Hrow("CollectionDate")
    '                        oPayment.DocCurrency = Hrow("CurrencyCode").ToString().Trim()
    '                        oPayment.DocRate = CDbl(Hrow("CurrencyRate").ToString().Trim())
    '                        If Hrow("CollectionDetailType") = 0 Then
    '                            oPayment.CashSum = CDbl(Hrow("CollectionAmount"))
    '                        ElseIf Hrow("CollectionDetailType") = 1 Then
    '                            oPayment.Checks.CheckNumber = Hrow("CheckCode").ToString().Trim()
    '                            oPayment.Checks.DueDate = Hrow("CheckDate")
    '                            oPayment.Checks.CheckSum = CDbl(Hrow("CollectionAmount"))
    '                            oPayment.Checks.BankCode = Hrow("BankCode").ToString().Trim()
    '                        End If
    '                        'Link Invoice for Payment if Record Exist Invoice Record Exist for Collection Code....
    '                        strQuery = " Select InvoiceCode,TotalAmount" & _
    '                                    " From CollectionInvoices " & _
    '                                    " Where CollectionCode = '" & CollectionCode & "'"
    '                        ds = ExecuteDataSet(strQuery)
    '                        If Not IsNothing(ds) Then
    '                            oDetailDt = ds.Tables(0)
    '                            If oDetailDt.Rows.Count > 0 Then
    '                                For Each Drow As DataRow In oDetailDt.Rows
    '                                    Dim oRecordset As SAPbobsCOM.Recordset
    '                                    oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                                    strQuery = "Select DocEntry,(DocTotal - PaidToDate) As 'Balance' From OINV "
    '                                    strQuery += "Where NumAtCard = '" & Drow("InvoiceCode").ToString().Trim() & "'"
    '                                    oRecordset.DoQuery(strQuery)
    '                                    If Not oRecordset.EoF Then
    '                                        oPayment.Invoices.SetCurrentLine(0)
    '                                        If CDbl(oRecordset.Fields.Item("Balance").Value) >= CDbl(Drow("TotalAmount")) Then
    '                                            oPayment.Invoices.DocEntry = CInt(oRecordset.Fields.Item("DocEntry").Value)
    '                                            oPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
    '                                            oPayment.Invoices.SumApplied = CDbl(Drow("TotalAmount"))
    '                                            oPayment.Invoices.Add()
    '                                        Else
    '                                            oPayment.Invoices.DocEntry = CInt(oRecordset.Fields.Item("DocEntry").Value)
    '                                            oPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
    '                                            oPayment.Invoices.SumApplied = CDbl(oRecordset.Fields.Item("Balance").Value)
    '                                            oPayment.Invoices.Add()
    '                                        End If
    '                                    End If
    '                                Next
    '                            End If
    '                        End If

    '                        Dim intStatus As Integer = oPayment.Add()

    '                        If intStatus = 0 Then
    '                            strQuery = " Update CollectionHeaders " & _
    '                                       " Set IsProcessed = 3 " & _
    '                                       " Where CollectionCode = " & CollectionCode
    '                            ExecuteNonQuery(strQuery)
    '                            Dim strDocNum As String
    '                            oCompany.GetNewObjectCode(strDocNum)
    '                            Dim oTest As SAPbobsCOM.Recordset
    '                            oTest = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                            oTest.DoQuery("Select * from ORCT where DocEntry=" & strDocNum)
    '                            strDocNum = oTest.Fields.Item("DocNum").Value
    '                            WriteErrorlog("Incoming Payment : " & strDocNum & " Created for Collection Code : " & TransactionCode, strErrorFileName)
    '                        Else
    '                            ' MessageBox.Show(oCompany.GetLastErrorDescription())
    '                            WriteErrorlog("Creating Incoming Payment for Collection Code : " & TransactionCode & " Failed. Error : " & oCompany.GetLastErrorDescription, strErrorFileName)

    '                        End If
    '                    End If

    '                Next
    '            End If
    '        End If
    '        WriteErrorlog("Processing Incoming Payment Completed ", strErrorFileName)
    '    Catch ex As Exception
    '        'traceService(ex.ToString())
    '    End Try
    'End Sub

    Private Sub import_Delivery()
        Try
            Dim oDelivery As SAPbobsCOM.Documents
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select T0.lvsboh_DocNum,T0.lvsboh_CustomerCode,T0.lvsboh_DocDate,T0.lvsboh_DocDueDate,T0.lvsboh_PostingDate, " & _
                        " T0.lvsboh_DocReference, ISNULL(T0.lvsboh_Remarks,'') As lvsboh_Remarks,  T0.lvsboh_Warehouse " & _
                        " From LVSBODocumentHeader T0 " & _
                        " Where T0.lvsboh_Status = 0  " 
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oDelivery = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                        oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        DocNum = Hrow("lvsboh_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oDelivery.CardCode = Hrow("lvsboh_CustomerCode").ToString().Trim()
                            oDelivery.DocDate = Hrow("lvsboh_DocDate")
                            oDelivery.TaxDate = Hrow("lvsboh_PostingDate")
                            oDelivery.DocDueDate = Hrow("lvsboh_DocDueDate")
                            oDelivery.NumAtCard = DocNum '& "\" & Hrow("lvsboh_DocReference").ToString
                            oDelivery.Comments = Hrow("lvsboh_Remarks")

                            sQuery = " Select T1.lvsbol_ItemCode, T1.lvsbol_Quantity, ISNULL(T1.lvsbol_Price,0) As lvsbol_Price, ISNULL(T1.lvsbol_Discount,0) As lvsbol_Discount, " & _
                                        " T1.lvsbol_BaseType,T1.lvsbol_BaseRef,T1.lvsbol_BaseLine,ISNULL(T1.lvsbol_UOMCode,'') As lvsbol_UOMCode " & _
                                        " From LVSBODocumentLines T1 " & _
                                        " Where T1.lvsbol_DocNum = '" & DocNum & "'" & _
                                        " Order By T1.lvsbol_BaseLine "
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True
                                        If intRow > 0 Then
                                            oDelivery.Lines.Add()
                                        End If
                                        oDelivery.Lines.SetCurrentLine(intRow)

                                        oDelivery.Lines.ItemCode = Drow("lvsbol_ItemCode").ToString().Trim()
                                        oDelivery.Lines.Quantity = Drow("lvsbol_Quantity")
                                        oDelivery.Lines.UnitPrice = Drow("lvsbol_Price")
                                        oDelivery.Lines.DiscountPercent = Drow("lvsbol_Discount")
                                        oDelivery.Lines.WarehouseCode = Hrow("lvsboh_Warehouse").ToString().Trim()
                                        Dim strUOMCode As String = IIf(Drow("lvsbol_UOMCode").ToString() <> "", Drow("lvsbol_UOMCode").ToString(), "")
                                        If strUOMCode.Length > 0 Then
                                            Dim intUOMCode As Integer = getUOMEntry(Drow("lvsbol_UOMCode").ToString())
                                            If intUOMCode > 0 Then
                                                oDelivery.Lines.UoMEntry = intUOMCode
                                            End If
                                        End If
                                        If Not IsDBNull(Drow("lvsbol_BaseRef")) Then
                                            If Drow("lvsbol_BaseRef").ToString() <> "" Then
                                                oDelivery.Lines.BaseType = 17
                                                oDelivery.Lines.BaseEntry = Drow("lvsbol_BaseRef")
                                                oDelivery.Lines.BaseLine = Drow("lvsbol_BaseLine")
                                            End If
                                        End If
                                        intRow = intRow + 1
                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oDelivery.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVSBODocumentHeader" & _
                                               " Set lvsboh_Status = 1 " & _
                                               " Where lvsboh_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)
                                    Dim strDocEntry As String = oCompany.GetNewObjectKey()
                                    Dim strSDocNum As String = String.Empty
                                    sQuery = "Select DocNum From ODLN Where DocEntry = '" & strDocEntry & "'"
                                    oRecordSet.DoQuery(sQuery)
                                    If Not oRecordSet.EoF Then
                                        strSDocNum = oRecordSet.Fields.Item(0).Value
                                    End If
                                    WriteErrorlog("Delivery - Document No : " & DocNum & " imported sucessfully... SAP Document No : " & strSDocNum, sPath)
                                    closeOpenOrders_Lines(DocNum)
                                Else
                                    Dim strMessage As String = oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription()
                                    WriteErrorlog("Delivery - Document No : " & DocNum & " Error :" & strMessage, sPath)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            WriteErrorlog("Delivery - Document No : " & DocNum & " Error :" & ex.Message, sPath)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Public Function closeOpenOrders_Lines(ByVal strDocNum As String) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim ds As DataSet
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oOrder As SAPbobsCOM.Documents
            Dim intStatus As Integer

            sQuery = " Select T1.lvsbol_BaseRef,T1.lvsbol_BaseLine,T1.lvsbol_IsClosing " & _
                                       " From LVSBODocumentHeader T0 " & _
                                       " JOIN LVSBODocumentLines T1 On T0.lvsboh_DocNum = T1.lvsbol_DocNum " & _
                                       " Where T0.lvsboh_DocNum = '" & strDocNum & "' "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oLineDt_C = ds.Tables(0)
                If oLineDt_C.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oLineDt_C.Rows

                        sQuery = "Update LVSBODocumentLines" & _
                                              " Set lvsbol_Status = 1 " & _
                                              " Where lvsbol_DocNum = '" & strDocNum & "'" & _
                                              " And lvsbol_LineNo = '" & Hrow("lvsbol_BaseLine") & "'"
                        ExecuteNonQuery(sQuery)

                        If Not IsDBNull(Hrow("lvsbol_IsClosing")) Then
                            If Hrow("lvsbol_IsClosing").ToString() = "1" Then
                                sQuery = "Select DocEntry,VisOrder From RDR1 Where DocEntry = '" & Hrow("lvsbol_BaseRef") & "' And LineStatus = 'O' " & _
                                    " And LineNum = '" & Hrow("lvsbol_BaseLine") & "'"
                                oRecordSet.DoQuery(sQuery)
                                If Not oRecordSet.EoF Then
                                    oOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                                    If oOrder.GetByKey(CInt(oRecordSet.Fields.Item("DocEntry").Value)) Then
                                        oOrder.Lines.SetCurrentLine(CInt(oRecordSet.Fields.Item("VisOrder").Value))
                                        If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                            oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                            intStatus = oOrder.Update()
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            'Dim strMessage As String = oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription()
            WriteErrorlog("Document No :" & strDocNum & " Error :" & ex.Message, sPath)
        End Try
    End Function

    Private Sub import_Return()
        Try
            Dim oReturn As SAPbobsCOM.Documents
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select T0.lvrboh_DocNum,T0.lvrboh_CustomerCode,T0.lvrboh_DocDate,T0.lvrboh_DocDueDate,T0.lvrboh_PostingDate, " & _
                        " T0.lvrboh_DocReference, ISNULL(T0.lvrboh_Remarks,'') As lvrboh_Remarks,  T0.lvrboh_Warehouse " & _
                        " From LVRBODocumentHeader T0 " & _
                        " Where T0.lvrboh_Status = 0  " 
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oReturn = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns)
                        oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        DocNum = Hrow("lvrboh_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oReturn.CardCode = Hrow("lvrboh_CustomerCode").ToString().Trim()
                            oReturn.DocDate = Hrow("lvrboh_DocDate")
                            oReturn.TaxDate = Hrow("lvrboh_PostingDate")
                            oReturn.DocDueDate = Hrow("lvrboh_DocDueDate")
                            oReturn.NumAtCard = DocNum '& "\" & Hrow("lvrboh_DocReference")
                            oReturn.Comments = Hrow("lvrboh_Remarks")

                            sQuery = " Select T1.lvrbol_ItemCode, T1.lvrbol_Quantity, ISNULL(T1.lvrbol_Price,0) As lvrbol_Price, ISNULL(T1.lvrbol_Discount,0) As lvrbol_Discount, " & _
                                        " T1.lvrbol_BaseType,T1.lvrbol_BaseRef,T1.lvrbol_BaseLine,ISNULL(T1.lvrbol_UOMCode,'') As lvrbol_UOMCode " & _
                                        " From LVRBODocumentLines T1 " & _
                                        " Where T1.lvrbol_DocNum = '" & DocNum & "'" & _
                                        " Order By T1.lvrbol_BaseLine "
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True
                                        If intRow > 0 Then
                                            oReturn.Lines.Add()
                                        End If
                                        oReturn.Lines.SetCurrentLine(intRow)

                                        oReturn.Lines.ItemCode = Drow("lvrbol_ItemCode").ToString().Trim()
                                        oReturn.Lines.Quantity = Drow("lvrbol_Quantity")
                                        oReturn.Lines.UnitPrice = Drow("lvrbol_Price")
                                        oReturn.Lines.DiscountPercent = Drow("lvrbol_Discount")
                                        oReturn.Lines.WarehouseCode = Hrow("lvrboh_Warehouse").ToString().Trim()
                                        Dim strUOMCode As String = IIf(Drow("lvrbol_UOMCode").ToString() <> "", Drow("lvrbol_UOMCode").ToString(), "")
                                        If strUOMCode.Length > 0 Then
                                            Dim intUOMCode As Integer = getUOMEntry(Drow("lvrbol_UOMCode").ToString())
                                            If intUOMCode > 0 Then
                                                oReturn.Lines.UoMEntry = intUOMCode
                                            End If
                                        End If
                                        If Not IsDBNull(Drow("lvrbol_BaseRef")) Then
                                            If Drow("lvrbol_BaseRef").ToString() <> "" Then

                                                Dim strRefCode As String = IIf(Drow("lvrbol_BaseRef").ToString() <> "", Drow("lvrbol_BaseRef").ToString(), "")
                                                If strRefCode.Length > 0 Then
                                                    Dim intBaseRef As Integer = getBaseEntry(Drow("lvrbol_BaseRef").ToString())
                                                    If intBaseRef > 0 Then
                                                        'oReturn.Lines.UoMEntry = intUOMCode
                                                        oReturn.Lines.BaseType = 15
                                                        oReturn.Lines.BaseEntry = intBaseRef 'Drow("lvrbol_BaseRef")
                                                        oReturn.Lines.BaseLine = Drow("lvrbol_BaseLine")
                                                    End If
                                                End If

                                            End If
                                        End If
                                        intRow = intRow + 1
                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oReturn.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVRBODocumentHeader" & _
                                               " Set lvrboh_Status = 1 " & _
                                               " Where lvrboh_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                    Dim strDocEntry As String = oCompany.GetNewObjectKey()
                                    Dim strSDocNum As String = String.Empty
                                    sQuery = "Select DocNum From ORDN Where DocEntry = '" & strDocEntry & "'"
                                    oRecordSet.DoQuery(sQuery)
                                    If Not oRecordSet.EoF Then
                                        strSDocNum = oRecordSet.Fields.Item(0).Value
                                    End If
                                    WriteErrorlog("Sale Return - Document No :" & DocNum & " imported sucessfully...SAP Document No : " & strSDocNum, sPath)

                                    'Close Return Lines Based on Flag...
                                    closeOpenReturn_Lines(DocNum)
                                Else
                                  
                                    Dim strMessage As String = oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription()
                                    WriteErrorlog("Sale Return - Document No :" & DocNum & " Error :" & strMessage, sPath)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            WriteErrorlog("Sale Return - Document No :" & DocNum & " Error :" & ex.Message, sPath)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Public Function closeOpenReturn_Lines(ByVal strDocNum As String) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim ds As DataSet

            sQuery = " Select T1.lvrbol_BaseRef,T1.lvrbol_BaseLine,T1.lvrbol_IsClosing " & _
                                       " From LVRBODocumentHeader T0 " & _
                                       " JOIN LVRBODocumentLines T1 On T0.lvrboh_DocNum = T1.lvrbol_DocNum " & _
                                       " Where T0.lvrboh_DocNum = '" & strDocNum & "' "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oLineDt_C = ds.Tables(0)
                If oLineDt_C.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oLineDt_C.Rows

                        sQuery = "Update LVRBODocumentLines" & _
                                              " Set lvrbol_Status = 1 " & _
                                              " Where lvrbol_DocNum = '" & strDocNum & "'" & _
                                              " And lvrbol_LineNo = '" & Hrow("lvrbol_BaseLine") & "'"

                        ExecuteNonQuery(sQuery)
                    Next
                End If
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            WriteErrorlog("Document No :" & strDocNum & " Error :" & ex.Message, sPath)
        End Try
    End Function

    Private Sub import_Purchase()
        Try
            Dim oGRPO As SAPbobsCOM.Documents
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select T0.lvpboh_DocNum,T0.lvpboh_CustomerCode,T0.lvpboh_DocDate,T0.lvpboh_DocDueDate,T0.lvpboh_PostingDate, " & _
                        " T0.lvpboh_DocReference, ISNULL(T0.lvpboh_Remarks,'') As lvpboh_Remarks,  T0.lvpboh_Warehouse " & _
                        " From LVPBODocumentHeader T0 " & _
                        " Where T0.lvpboh_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oGRPO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        DocNum = Hrow("lvpboh_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oGRPO.CardCode = Hrow("lvpboh_CustomerCode").ToString().Trim()
                            oGRPO.DocDate = Hrow("lvpboh_DocDate")
                            oGRPO.TaxDate = Hrow("lvpboh_PostingDate")
                            oGRPO.DocDueDate = Hrow("lvpboh_DocDueDate")
                            oGRPO.NumAtCard = DocNum '& "\" & Hrow("lvpboh_DocReference")
                            oGRPO.Comments = Hrow("lvpboh_Remarks")

                            sQuery = " Select T1.lvpbol_ItemCode, T1.lvpbol_Quantity, ISNULL(T1.lvpbol_Price,0) As lvpbol_Price, ISNULL(T1.lvpbol_Discount,0) As lvpbol_Discount, " & _
                                        " T1.lvpbol_BaseType,T1.lvpbol_BaseRef,T1.lvpbol_BaseLine,ISNULL(lvpbol_UOMCode,'') As lvpbol_UOMCode " & _
                                        " From LVPBODocumentLines T1 " & _
                                        " Where T1.lvpbol_DocNum = '" & DocNum & "'" & _
                                        " Order By T1.lvpbol_BaseLine "
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True
                                        If intRow > 0 Then
                                            oGRPO.Lines.Add()
                                        End If
                                        oGRPO.Lines.SetCurrentLine(intRow)

                                        oGRPO.Lines.ItemCode = Drow("lvpbol_ItemCode").ToString().Trim()
                                        oGRPO.Lines.Quantity = Drow("lvpbol_Quantity")
                                        oGRPO.Lines.UnitPrice = Drow("lvpbol_Price")
                                        oGRPO.Lines.DiscountPercent = Drow("lvpbol_Discount")
                                        oGRPO.Lines.WarehouseCode = Hrow("lvpboh_Warehouse").ToString().Trim()
                                        Dim strUOMCode As String = IIf(Drow("lvpbol_UOMCode").ToString() <> "", Drow("lvpbol_UOMCode").ToString(), "")
                                        If strUOMCode.Length > 0 Then
                                            Dim intUOMCode As Integer = getUOMEntry(Drow("lvpbol_UOMCode").ToString())
                                            If intUOMCode > 0 Then
                                                oGRPO.Lines.UoMEntry = intUOMCode
                                            End If
                                        End If
                                        If Not IsDBNull(Drow("lvpbol_BaseRef")) Then
                                            If Drow("lvpbol_BaseRef").ToString() <> "" Then
                                                oGRPO.Lines.BaseType = 22
                                                oGRPO.Lines.BaseEntry = Drow("lvpbol_BaseRef")
                                                oGRPO.Lines.BaseLine = Drow("lvpbol_BaseLine")
                                            End If
                                        End If
                                        intRow = intRow + 1
                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oGRPO.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVPBODocumentHeader" & _
                                               " Set lvpboh_Status = 1 " & _
                                               " Where lvpboh_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                    Dim strDocEntry As String = oCompany.GetNewObjectKey()
                                    Dim strSDocNum As String = String.Empty
                                    sQuery = "Select DocNum From OPDN Where DocEntry = '" & strDocEntry & "'"
                                    oRecordSet.DoQuery(sQuery)
                                    If Not oRecordSet.EoF Then
                                        strSDocNum = oRecordSet.Fields.Item(0).Value
                                    End If
                                    WriteErrorlog("GRPO - Document No : " & DocNum & " imported sucessfully...SAP Document No : " & strSDocNum, sPath)

                                    closeOpenPurchaseOrders_Lines(DocNum) ' Close Purchase Lines.
                                Else
                                    Dim strMessage As String = oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription()
                                    WriteErrorlog("GRPO Document No : " & DocNum & " Error :" & strMessage, sPath)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            WriteErrorlog("GRPO Document No : " & DocNum & " Error :" & ex.Message, sPath)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Public Function closeOpenPurchaseOrders_Lines(ByVal strDocNum As String) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim ds As DataSet
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oOrder As SAPbobsCOM.Documents
            Dim intStatus As Integer

            sQuery = " Select T1.lvpbol_BaseRef,T1.lvpbol_BaseLine,T1.lvpbol_IsClosing " & _
                                       " From LVPBODocumentHeader T0 " & _
                                       " JOIN LVPBODocumentLines T1 On T0.lvpboh_DocNum = T1.lvpbol_DocNum " & _
                                       " Where T0.lvpboh_DocNum = '" & strDocNum & "' "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oLineDt_C = ds.Tables(0)
                If oLineDt_C.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oLineDt_C.Rows

                        sQuery = "Update LVPBODocumentLines" & _
                                              " Set lvpbol_Status = 1 " & _
                                              " Where lvpbol_DocNum = '" & strDocNum & "'" & _
                                              " And lvpbol_LineNo = '" & Hrow("lvpbol_BaseLine") & "'"
                        ExecuteNonQuery(sQuery)

                        If Not IsDBNull(Hrow("lvpbol_IsClosing")) Then
                            If Hrow("lvpbol_IsClosing").ToString() = "1" Then
                                sQuery = "Select DocEntry,VisOrder From POR1 Where DocEntry = '" & Hrow("lvpbol_BaseRef") & "' And LineStatus = 'O' " & _
                                    " And LineNum = '" & Hrow("lvpbol_BaseLine") & "'"
                                oRecordSet.DoQuery(sQuery)
                                If Not oRecordSet.EoF Then
                                    oOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                    If oOrder.GetByKey(CInt(oRecordSet.Fields.Item("DocEntry").Value)) Then
                                        oOrder.Lines.SetCurrentLine(CInt(oRecordSet.Fields.Item("VisOrder").Value))
                                        If oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Open Then
                                            oOrder.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                                            intStatus = oOrder.Update()
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            WriteErrorlog("Document No :" & strDocNum & " Error :" & ex.Message, sPath)
            'Throw ex
        End Try
    End Function

    Private Sub import_Transfer()
        Try
            Dim oTransfer As SAPbobsCOM.StockTransfer
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select T0.lvtboh_DocNum,T0.lvtboh_DocDate,T0.lvtboh_DocDueDate,T0.lvtboh_PostingDate, " & _
                        " T0.lvtboh_DocReference, ISNULL(T0.lvtboh_Remarks,'') As lvtboh_Remarks,  T0.lvtboh_FWarehouse,  T0.lvtboh_TWarehouse " & _
                        " From LVTBODocumentHeader T0 " & _
                        " Where T0.lvtboh_Status = 0  " 
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oTransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        DocNum = Hrow("lvtboh_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            ' oTransfer.CardCode = Hrow("lvpboh_CustomerCode").ToString().Trim()
                            oTransfer.DocDate = Hrow("lvtboh_DocDate")
                            oTransfer.TaxDate = Hrow("lvtboh_PostingDate")
                            'oTransfer.DocDueDate = Hrow("lvtboh_DocDueDate")
                            oTransfer.FromWarehouse = Hrow("lvtboh_FWarehouse")
                            'oTransfer.ToWarehouse = Hrow("lvtboh_TWarehouse")
                            'oTransfer.NumAtCard = Hrow("lvpboh_DocReference")
                            oTransfer.Comments = Hrow("lvtboh_Remarks")

                            sQuery = " Select T1.lvtbol_ItemCode, T1.lvtbol_Quantity, ISNULL(T1.lvtbol_Price,0) As lvtbol_Price, ISNULL(T1.lvtbol_Discount,0) As lvtbol_Discount, " & _
                                        " T1.lvtbol_BaseType,T1.lvtbol_BaseRef,T1.lvtbol_BaseLine,ISNULL(T1.lvtbol_UOMCode,'') As lvtbol_UOMCode " & _
                                        " From LVTBODocumentLines T1 " & _
                                        " Where T1.lvtbol_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True

                                        If intRow > 0 Then
                                            oTransfer.Lines.Add()
                                        End If

                                        oTransfer.Lines.SetCurrentLine(intRow)

                                        oTransfer.Lines.ItemCode = Drow("lvtbol_ItemCode").ToString().Trim()
                                        oTransfer.Lines.Quantity = Drow("lvtbol_Quantity")
                                        oTransfer.Lines.UnitPrice = Drow("lvtbol_Price")
                                        oTransfer.Lines.DiscountPercent = Drow("lvtbol_Discount")
                                        oTransfer.Lines.WarehouseCode = Hrow("lvtboh_TWarehouse").ToString().Trim()
                                        Dim strUOMCode As String = IIf(Drow("lvtbol_UOMCode").ToString() <> "", Drow("lvtbol_UOMCode").ToString(), "")
                                        If strUOMCode.Length > 0 Then
                                            Dim intUOMCode As Integer = getUOMEntry(Drow("lvtbol_UOMCode").ToString())
                                            If intUOMCode > 0 Then
                                                oTransfer.Lines.UoMEntry = intUOMCode
                                            End If
                                        End If
                                        If Not IsDBNull(Drow("lvtbol_BaseRef")) Then
                                            If Drow("lvtbol_BaseRef").ToString() <> "" Then
                                                oTransfer.Lines.BaseType = SAPbobsCOM.InvBaseDocTypeEnum.InventoryTransferRequest
                                                oTransfer.Lines.BaseEntry = Drow("lvtbol_BaseRef")
                                                oTransfer.Lines.BaseLine = Drow("lvtbol_BaseLine")
                                            End If
                                        End If

                                        intRow = intRow + 1

                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oTransfer.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVTBODocumentHeader" & _
                                               " Set lvtboh_Status = 1 " & _
                                               " Where lvtboh_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                    Dim strDocEntry As String = oCompany.GetNewObjectKey()
                                    Dim strSDocNum As String = String.Empty
                                    sQuery = "Select DocNum From OWTR Where DocEntry = '" & strDocEntry & "'"
                                    oRecordSet.DoQuery(sQuery)
                                    If Not oRecordSet.EoF Then
                                        strSDocNum = oRecordSet.Fields.Item(0).Value
                                    End If
                                    WriteErrorlog("Transfer - Document No :" & DocNum & " imported sucessfully...SAP Document No : " & strSDocNum, sPath)

                                    closeOpenInventoryTransfer_Lines(DocNum)
                                Else
                                    Dim strMessage As String = oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription()
                                    WriteErrorlog("Transfer Document No :" & DocNum & " Error :" & strMessage, sPath)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            WriteErrorlog("Transfer Document No :" & DocNum & " Error :" & ex.Message, sPath)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Public Function closeOpenInventoryTransfer_Lines(ByVal strDocNum As String) As Boolean
        Try
            Dim _retVal As Boolean = False
            Dim ds As DataSet
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            'Dim oTransferR As SAPbobsCOM.StockTransfer
            'Dim intStatus As Integer

            sQuery = " Select T1.lvtbol_BaseRef,T1.lvtbol_BaseLine,T1.lvtbol_IsClosing " & _
                                       " From LVTBODocumentHeader T0 " & _
                                       " JOIN LVTBODocumentLines T1 On T0.lvtboh_DocNum = T1.lvtbol_DocNum " & _
                                       " Where T0.lvtboh_DocNum = '" & strDocNum & "' "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oLineDt_C = ds.Tables(0)
                If oLineDt_C.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oLineDt_C.Rows

                        sQuery = "Update LVTBODocumentLines" & _
                                              " Set lvtbol_Status = 1 " & _
                                              " Where lvtbol_DocNum = '" & strDocNum & "'" & _
                                              " And lvtbol_LineNo = '" & Hrow("lvtbol_BaseLine") & "'"
                        ExecuteNonQuery(sQuery)

                        'If Not IsDBNull(Hrow("lvtbol_IsClosing")) Then
                        '    If Hrow("lvtbol_IsClosing").ToString() = "1" Then
                        '        sQuery = "Select DocEntry,VisOrder From WTQ1 Where DocEntry = '" & Hrow("lvtbol_BaseRef") & "' And LineStatus = 'O' " & _
                        '            " And LineNum = '" & Hrow("lvtbol_BaseLine") & "'"
                        '        oRecordSet.DoQuery(sQuery)
                        '        If Not oRecordSet.EoF Then
                        '            oTransferR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest)
                        '            If oTransferR.GetByKey(CInt(oRecordSet.Fields.Item("DocEntry").Value)) Then
                        '                oTransferR.Lines.SetCurrentLine(CInt(oRecordSet.Fields.Item("VisOrder").Value))
                        '                oTransferR.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                        '                intStatus = oTransferR.Update()
                        '            End If
                        '        End If
                        '    End If
                        'End If
                    Next
                End If
            End If

            _retVal = True
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub import_InventoryReceipt()
        Try
            Dim oIReceipt As SAPbobsCOM.Documents
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select Distinct T0.lvibo_DocNum,T0.lvibo_DocDate,T0.lvibo_DocDueDate,T0.lvibo_PostingDate " & _
                        " From LVIBODocument T0 " & _
                        " Where T0.lvibo_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oIReceipt = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                        oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        DocNum = Hrow("lvibo_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oIReceipt.DocDate = Hrow("lvibo_DocDate")
                            oIReceipt.TaxDate = Hrow("lvibo_PostingDate")
                            oIReceipt.Comments = DocNum
                            oIReceipt.PaymentGroupCode = -2

                            sQuery = " Select T1.lvibo_ItemCode, T1.lvibo_Quantity, ISNULL(T1.lvibo_Price,0) As lvibo_Price, ISNULL(T1.lvibo_Discount,0) As lvibo_Discount,ISNULL(T1.lvibo_UOMCode,'') As lvibo_UOMCode " & _
                                        " ,T1.lvibo_Warehouse, ISNULL(T1.lvibo_Remarks,'') As lvibo_Remarks From LVIBODocument T1 " & _
                                        " Where T1.lvibo_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True

                                        If intRow > 0 Then
                                            oIReceipt.Lines.Add()
                                        End If

                                        oIReceipt.Lines.SetCurrentLine(intRow)

                                        oIReceipt.Lines.ItemCode = Drow("lvibo_ItemCode").ToString().Trim()
                                        oIReceipt.Lines.Quantity = Drow("lvibo_Quantity")

                                        Dim dblPrice As Double = CDbl(Drow("lvibo_Price"))
                                        If dblPrice = 0 Then
                                            dblPrice = getLastPurchasePrice(Drow("lvibo_ItemCode").ToString().Trim())
                                        End If

                                        oIReceipt.Lines.UnitPrice = dblPrice 'Drow("lvibo_Price")
                                        oIReceipt.Lines.DiscountPercent = Drow("lvibo_Discount")
                                        oIReceipt.Lines.WarehouseCode = Drow("lvibo_Warehouse").ToString().Trim()
                                        oIReceipt.Lines.UseBaseUnits = BoYesNoEnum.tNO

                                        Dim strUOMCode As String = IIf(Drow("lvibo_UOMCode").ToString() <> "", Drow("lvibo_UOMCode").ToString(), "")
                                        If strUOMCode.Length > 0 Then
                                            Dim intUOMCode As Integer = getUOMEntry(Drow("lvibo_UOMCode").ToString())
                                            If intUOMCode > 0 Then
                                                oIReceipt.Lines.UoMEntry = intUOMCode
                                            End If
                                        End If
                                        oIReceipt.Lines.FreeText = Drow("lvibo_Remarks").ToString().Trim()

                                        intRow = intRow + 1

                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oIReceipt.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVIBODocument" & _
                                               " Set lvibo_Status = 1 " & _
                                               " Where lvibo_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                    Dim strDocEntry As String = oCompany.GetNewObjectKey()
                                    Dim strSDocNum As String = String.Empty
                                    sQuery = "Select DocNum From OIGN Where DocEntry = '" & strDocEntry & "'"
                                    oRecordSet.DoQuery(sQuery)
                                    If Not oRecordSet.EoF Then
                                        strSDocNum = oRecordSet.Fields.Item(0).Value
                                    End If

                                    WriteErrorlog("Inventory Goods Receipt - Document No :" & DocNum & " imported sucessfully...SAP Document No : " & strSDocNum, sPath)
                                Else
                                    Dim strMessage As String = oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription()
                                    WriteErrorlog("Inventory Goods Receipt Document No :" & DocNum & " Error :" & strMessage, sPath)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            WriteErrorlog("Inventory Goods Receipt Document No :" & DocNum & " Error :" & ex.Message, sPath)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Private Sub import_InventoryIssue()
        Try
            Dim oIIssue As SAPbobsCOM.Documents
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select Distinct T0.lvobo_DocNum,T0.lvobo_DocDate,T0.lvobo_DocDueDate,T0.lvobo_PostingDate " & _
                        " From LVOBODocument T0 " & _
                        " Where T0.lvobo_Status = 0  "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        oIIssue = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        DocNum = Hrow("lvobo_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oIIssue.DocDate = Hrow("lvobo_DocDate")
                            oIIssue.TaxDate = Hrow("lvobo_PostingDate")
                            oIIssue.Comments = DocNum

                            sQuery = " Select T1.lvobo_ItemCode, T1.lvobo_Quantity, ISNULL(T1.lvobo_Price,0) As lvobo_Price, ISNULL(T1.lvobo_Discount,0) As lvobo_Discount,ISNULL(T1.lvobo_UOMCode,'') As lvobo_UOMCode " & _
                                        " , ISNULL(T1.lvobo_Remarks,'') As lvobo_Remarks,  T1.lvobo_Warehouse From LVOBODocument T1 " & _
                                        " Where T1.lvobo_DocNum = '" & DocNum & "'"
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0
                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True

                                        If intRow > 0 Then
                                            oIIssue.Lines.Add()
                                        End If

                                        oIIssue.Lines.SetCurrentLine(intRow)

                                        oIIssue.Lines.ItemCode = Drow("lvobo_ItemCode").ToString().Trim()
                                        oIIssue.Lines.Quantity = Drow("lvobo_Quantity")
                                        oIIssue.Lines.UnitPrice = Drow("lvobo_Price")
                                        oIIssue.Lines.DiscountPercent = Drow("lvobo_Discount")
                                        oIIssue.Lines.WarehouseCode = Drow("lvobo_Warehouse").ToString().Trim()
                                        oIIssue.Lines.UseBaseUnits = BoYesNoEnum.tNO

                                        Dim strUOMCode As String = IIf(Drow("lvobo_UOMCode").ToString() <> "", Drow("lvobo_UOMCode").ToString(), "")
                                        If strUOMCode.Length > 0 Then
                                            Dim intUOMCode As Integer = getUOMEntry(Drow("lvobo_UOMCode").ToString())
                                            If intUOMCode > 0 Then
                                                oIIssue.Lines.UoMEntry = intUOMCode
                                            End If
                                        End If

                                        oIIssue.Lines.FreeText = Drow("lvobo_Remarks").ToString().Trim()

                                        intRow = intRow + 1

                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim intStatus As Integer = oIIssue.Add()

                                If intStatus = 0 Then
                                    sQuery = "Update LVOBODocument" & _
                                               " Set lvobo_Status = 1 " & _
                                               " Where lvobo_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                    Dim strDocEntry As String = oCompany.GetNewObjectKey()
                                    Dim strSDocNum As String = String.Empty
                                    sQuery = "Select DocNum From OIGE Where DocEntry = '" & strDocEntry & "'"
                                    oRecordSet.DoQuery(sQuery)
                                    If Not oRecordSet.EoF Then
                                        strSDocNum = oRecordSet.Fields.Item(0).Value
                                    End If

                                    WriteErrorlog("Inventory Goods Issue - Document No :" & DocNum & " imported sucessfully...SAP Document No : " & strSDocNum, sPath)
                                Else
                                  
                                    Dim strMessage As String = oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription()
                                    WriteErrorlog("Inventory Goods Issue Document No :" & DocNum & " Error :" & strMessage, sPath)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            WriteErrorlog("Inventory Goods Issue Document No :" & DocNum & " Error :" & ex.Message, sPath)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Private Sub import_InventoryCounting()
        Try
            Dim ds As DataSet
            Dim DocNum As String
            sQuery = " Select Distinct T0.lvcbo_DocNum,T0.lvcbo_CountDate " & _
                        " From LVCBODocument T0 " & _
                        " Where T0.lvcbo_Status = 0 "
            ds = ExecuteDataSet(sQuery)

            If Not IsNothing(ds) Then
                oHeaderDt = ds.Tables(0)

                If oHeaderDt.Rows.Count > 0 Then
                    For Each Hrow As DataRow In oHeaderDt.Rows
                        Dim oCS As SAPbobsCOM.CompanyService = oCompany.GetCompanyService()
                        Dim oICS As SAPbobsCOM.InventoryCountingsService = oCS.GetBusinessService(ServiceTypes.InventoryCountingsService)
                        Dim oIC As SAPbobsCOM.InventoryCounting = oICS.GetDataInterface(InventoryCountingsServiceDataInterfaces.icsInventoryCounting)
                        DocNum = Hrow("lvcbo_DocNum").ToString().Trim()
                        Dim blnRowExist As Boolean = False

                        Try

                            oIC.CountDate = Hrow("lvcbo_CountDate")
                            oIC.SingleCounterType = CounterTypeEnum.ctUser
                            oIC.SingleCounterID = oCompany.UserSignature
                            oIC.Remarks = DocNum

                            sQuery = " Select T1.lvcbo_ItemCode, T1.lvcbo_Quantity,T1.lvcbo_UOMEntry,T1.lvcbo_UOMName " & _
                                    "  ,T1.lvcbo_DocReference, ISNULL(T1.lvcbo_Remarks,'') As lvcbo_Remarks,T1.lvcbo_Warehouse " & _
                                        " From LVCBODocument T1 " & _
                                        " Where T1.lvcbo_DocNum = '" & DocNum & "' "
                            ds = ExecuteDataSet(sQuery)

                            If Not IsNothing(ds) Then
                                oDetailDt = ds.Tables(0)

                                If oDetailDt.Rows.Count > 0 Then
                                    blnRowExist = True
                                    Dim intRow As Integer = 0

                                    Dim oICLS As SAPbobsCOM.InventoryCountingLines = oIC.InventoryCountingLines

                                    For Each Drow As DataRow In oDetailDt.Rows
                                        blnRowExist = True
                                        Dim oICL As SAPbobsCOM.InventoryCountingLine = oICLS.Add

                                        'If intRow > 0 Then
                                        '    oIIssue.Lines.Add()
                                        'End If

                                        ' oIIssue.Lines.SetCurrentLine(intRow)

                                        oICL.ItemCode = Drow("lvcbo_ItemCode").ToString().Trim()
                                        oICL.CountedQuantity = Drow("lvcbo_Quantity")
                                        oICL.Counted = BoYesNoEnum.tYES
                                        oICL.CountedQuantity = Drow("lvcbo_Quantity")
                                        oICL.WarehouseCode = Drow("lvcbo_Warehouse").ToString().Trim()
                                        oICL.UoMCode = Drow("lvcbo_UOMEntry")


                                        'intRow = intRow + 1

                                    Next ' Row
                                End If
                            End If

                            If blnRowExist Then
                                Dim oICP As SAPbobsCOM.InventoryCountingParams = oICS.Add(oIC)

                                If oICP.DocumentEntry > 0 Then

                                    sQuery = "Update LVCBODocument" & _
                                               " Set lvcbo_Status = 1 " & _
                                               " Where lvcbo_DocNum = '" & DocNum & "'"
                                    ExecuteNonQuery(sQuery)

                                    WriteErrorlog("Inventory Counting - Document No :" & DocNum & " imported sucessfully...SAP Document No : " & oICP.DocumentNumber.ToString, sPath)
                                Else
                                    Dim strMessage As String = oCompany.GetLastErrorCode().ToString() & "-" & oCompany.GetLastErrorDescription()
                                    WriteErrorlog("Inventory Counting Document No :" & DocNum & " Error :" & strMessage, sPath)
                                End If
                            Else
                                Throw New Exception("No Row Exists")
                            End If
                        Catch ex As Exception
                            WriteErrorlog("Inventory Counting - Document No :" & DocNum & " Error :" & ex.Message, sPath)
                        End Try ' For Each Header Record
                    Next ' Header
                End If
            End If
        Catch ex As Exception
            traceService(ex.ToString())
        End Try

    End Sub

    Private Sub connectSAP()
        Try
            'traceService("Company Connection")
            oCompany = objRemoteCompany ' New SAPbobsCOM.Company()
            'Dim strCompany As String = System.Configuration.ConfigurationManager.AppSettings("CompanyDB").ToString
            'oCompany.CompanyDB = strCompany

            'oCompany.UserName = System.Configuration.ConfigurationManager.AppSettings("UserName").ToString
            'oCompany.Password = System.Configuration.ConfigurationManager.AppSettings("Password").ToString
            'oCompany.Server = System.Configuration.ConfigurationManager.AppSettings("Server").ToString
            'oCompany.DbUserName = System.Configuration.ConfigurationManager.AppSettings("DbUserName").ToString
            'oCompany.DbPassword = System.Configuration.ConfigurationManager.AppSettings("DbPassword").ToString

            'If System.Configuration.ConfigurationManager.AppSettings("DBType").ToString = "2008" Then
            '    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            'ElseIf System.Configuration.ConfigurationManager.AppSettings("DBType").ToString = "2012" Then
            '    'oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            'End If

            'oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            'oCompany.LicenseServer = System.Configuration.ConfigurationManager.AppSettings("LicenseServer").ToString

            'If (0 <> oCompany.Connect()) Then
            '    ''MessageBoxoCompany.GetLastErrorDescription())
            '    'MsgBox("Failed to connect")
            '    'traceService("Company Connection Failed")
            '    'traceService(oCompany.GetLastErrorDescription())
            '    Exit Sub
            'Else
            '    'traceService("Connected successfully")
            'End If
        Catch ex As Exception
            'traceService(ex.ToString())
        End Try
    End Sub

    Private Sub disconnectSAP()
        Try
            If oCompany.Connected Then
                oCompany.Disconnect()
            End If
        Catch ex As Exception
            'traceService(ex.ToString())
        End Try
    End Sub

    Public Function getDueDate(ByVal cardCode As String, ByVal refDate As Date) As Date
        Dim _retVal As Date
        Try
            Dim vObj As SAPbobsCOM.SBObob
            Dim oRecordSet As SAPbobsCOM.Recordset

            vObj = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = vObj.GetDueDate(cardCode, refDate)

            If Not oRecordSet.EoF Then
                _retVal = CDate(oRecordSet.Fields.Item(0).Value)
            End If

            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getCurrencyRate(ByVal currency As String, ByVal refDate As Date) As Double
        Dim _retVal As Double
        Try
            Dim vObj As SAPbobsCOM.SBObob
            Dim oRecordSet As SAPbobsCOM.Recordset

            vObj = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = vObj.GetCurrencyRate(currency, refDate)

            If Not oRecordSet.EoF Then
                _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
            End If

            Return _retVal
        Catch ex As Exception
            Return 1
            'Throw ex
        End Try
    End Function

    Public Function getLocalCurrency() As String
        Dim _retVal As String = String.Empty
        Try
            Dim vObj As SAPbobsCOM.SBObob
            Dim oRecordSet As SAPbobsCOM.Recordset

            vObj = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = vObj.GetLocalCurrency()
            If Not oRecordSet.EoF Then
                _retVal = (oRecordSet.Fields.Item(0).Value)
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getUOMEntry(ByVal strCode As String) As String
        Dim _retVal As String = String.Empty
        Try
            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset

            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strCode <> "" Then
                oRecordSet.DoQuery("Select UomEntry From OUOM Where UomCode = '" & strCode & "'")
                If Not oRecordSet.EoF Then
                    _retVal = (oRecordSet.Fields.Item(0).Value)
                Else
                    _retVal = -1
                End If
            Else
                _retVal = -1
            End If

            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getBaseEntry(ByVal strRef As String) As String
        Dim _retVal As String = String.Empty
        Try
            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset

            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strRef <> "" Then
                oRecordSet.DoQuery("Select DocEntry From ODLN Where NumAtCard = '" & strRef & "'")
                If Not oRecordSet.EoF Then
                    _retVal = (oRecordSet.Fields.Item(0).Value)
                Else
                    _retVal = -1
                End If
            Else
                _retVal = -1
            End If

            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getLastPurchasePrice(ByVal strItem As String) As Double
        Dim _retVal As Double = 0
        Try
            Dim strQuery As String = String.Empty
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strItem <> "" Then
                oRecordSet.DoQuery("Select ISNULL(LstEvlPric,0) As LastPurPrc  From OITM Where ItemCode = '" & strItem & "'")
                If Not oRecordSet.EoF Then
                    _retVal = CDbl(oRecordSet.Fields.Item(0).Value)
                Else
                    _retVal = 0
                End If
            Else
                _retVal = 0
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "Events"

    Private Sub frmLoginScreen_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

    End Sub

    Private Sub frmLoginScreen_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        System.GC.Collect()
        End
    End Sub
    Private Function CheckDuplicateProcess() As Boolean
        Dim MatchingNames As System.Diagnostics.Process()
        Dim TargetName As String
        TargetName = System.Diagnostics.Process.GetCurrentProcess.ProcessName

        MatchingNames = System.Diagnostics.Process.GetProcessesByName(TargetName)
        If (MatchingNames.Length = 1) Then
            Return True

        Else
            Return False

        End If



    End Function
    '****************************************************************************
    'Type              :   Procedure     
    'Name              :   
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Handle Events
    '****************************************************************************
    Private Sub frmLoginScreen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim strPath, strFilename, strMessage As String
        Try
            If CheckDuplicateProcess() = False Then
                End
            End If
            ReadiniFile()
            If ValidateFilePaths() = False Then
                'End
                Exit Sub
            End If
            strPath = strErrorLogPath ' System.Windows.Forms.Application.StartupPath
            strFilename = Now.ToLongDateString
            strPath = strPath & "\WMS_Log_" & strFilename & ".txt"
            strErrorFileName = strPath
            WriteErrorHeader(strPath, "Start connecting..")
            If ValidateInterFaceDBConnection() = False Then
                End
            End If
            If connectLocalCompany() = False Then ' Or connectMainCompany() = False Then
                strMessage = ("Connection to SAP B1 failed. Check the ConnectionInfo.ini ")
                WriteErrorlog(strMessage, strPath)
                '   MsgBox(strMessage)
                End
            Else
                sPath = strErrorLogPath
                WriteErrorlog("Import Process Started", sPath)
                Me.Hide()
                Me.Hide()
                mainFunction()
                WriteErrorlog("Import Process completed", sPath)
                Me.Hide()
                Me.Hide()
                '  WriteErrorlog("Export Process Completed", sPath)
                'MsgBox("Export process Compleated")
                End
            End If
        Catch ex As Exception
            strMessage = (ex.Message)
            WriteErrorlog(strMessage, strPath)
            'MsgBox(strMessage)
            End
        End Try
    End Sub

#Region "Check the Filepaths"
    Private Function ValidateFilePaths() As Boolean
        Dim strMessage, strpath, strFilename As String
        strpath = strErrorLogPath ' System.Windows.Forms.Application.StartupPath

        If Directory.Exists(strpath) = False Then
            Directory.CreateDirectory(strpath)
            ' strMessage = "Error Log file direcory does not exists. check the connectionInfo.ini"
            ' MsgBox(strMessage)
            ' Return False
        End If


        If Directory.Exists(strpath) = False Then
            Directory.CreateDirectory(strpath)
            'strMessage = ("Export Directory does not exist. Check the connectionInfo.ini")
            'WriteErrorlog(strMessage, strpath)
            'MsgBox(strMessage)
            'Return False
        End If

        If Directory.Exists(strSuccessFolder) = False Then
            Directory.CreateDirectory(strpath)
            'strMessage = ("Export Directory does not exist. Check the connectionInfo.ini")
            'WriteErrorlog(strMessage, strpath)
            'MsgBox(strMessage)
            'Return False
        End If



        strFilename = Now.ToLongDateString



        If Directory.Exists(strExportFilePaty) = False Then
            strMessage = ("Export Directory does not exist. Check the connectionInfo.ini")
            WriteErrorlog(strMessage, strpath)
            MsgBox(strMessage)
            Return False
        End If
        'If strFileStart = "" Then
        '    strMessage = "Export File Name is missing"
        '    MsgBox(strMessage)
        '    Return False
        'End If

        Return True
    End Function
#End Region

    Private Sub btnReferesh_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReferesh.Click
        Try
            If txtServerName.Text = "" Then
                MsgBox("Enter Server")
                Exit Sub
            End If
            If txtDBUserName.Text = "" Then
                MsgBox("Enter UserName")
                Exit Sub
            End If
            If txtServerPwd.Text = "" Then
                MsgBox("Enter Password")
                Exit Sub
            End If
            strLocalServertype = cmbServertype.Text
            oCompany = New SAPbobsCOM.Company
            oCompany.Server = txtServerName.Text
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            If strLocalServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Else
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            End If
            oCompany.DbUserName = txtDBUserName.Text
            oCompany.DbPassword = txtServerPwd.Text
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim lErrCode As Long
            Dim sErrMsg As String
            oRecordSet = Me.oCompany.GetCompanyList
            Me.oCompany.GetLastError(lErrCode, sErrMsg)
            If lErrCode <> 0 Then
                MsgBox(sErrMsg)
                Exit Sub
            Else
                cmbCompanyDB.Items.Clear()
                Do Until oRecordSet.EoF = True
                    cmbCompanyDB.Items.Add(oRecordSet.Fields.Item(0).Value)
                    oRecordSet.MoveNext()
                Loop
                btnConnect.Enabled = True
            End If
        Catch ex As Exception
            MsgBox("Connection to Server =" & txtServerName.Text & ".  Failed")
            Exit Sub
        End Try
    End Sub

#Region "Write into ErrorLog File"
    Private Sub WriteErrorHeader(ByVal apath As String, ByVal strMessage As String)
        Dim aSw As System.IO.StreamWriter
        Dim aMessage As String
        aMessage = Now.ToLocalTime & "---" & strMessage
        If File.Exists(apath) Then
        End If
        aSw = New StreamWriter(apath, True)
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        aPath = strErrorFileName
        Dim aSw As System.IO.StreamWriter
        If File.Exists(aPath) Then
        End If
        aSw = New StreamWriter(aPath, True)
        aMessage = Now.ToLocalTime & "---" & aMessage
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region

#Region "exportSO"
    Private Function GetSalesOrders() As String
        Dim oSalesRec As SAPbobsCOM.Recordset
        oSalesRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strsql As String
        strsql = "SELECT convert(nvarchar(12),dbo.ORDR.DocNum) AS DocumentNoV1, dbo.ORDR.CardCode AS CustomerCode, "
        strsql = strsql & "       convert(nvarchar(10),dbo.ORDR.DocDate,112) AS DocumentDate, '' "
        strsql = strsql & " AS DeliveryDate, coalesce(dbo.ORDR.NumAtCard,'') AS CustomerRef, "
        strsql = strsql & " convert(nvarchar(12),dbo.RDR1.LineNum) as LineNum, dbo.RDR1.ItemCode, "
        strsql = strsql & " dbo.RDR1.Dscription AS ItemDesc, "
        strsql = strsql & " coalesce(dbo.[@MORPHI].Name,'') AS Presentation, convert(nvarchar(12),convert(int,round(dbo.RDR1.Quantity,0))), "
        strsql = strsql & " '' AS PickedQuantity, '' "
        strsql = strsql & " AS ExpDate, '' AS Batch FROM   dbo.ORDR INNER JOIN  dbo.RDR1 ON dbo.ORDR.DocEntry = dbo.RDR1.DocEntry left outer JOIN "
        strsql = strsql & " dbo.[@MORPHI] ON dbo.RDR1.U_TYPE = dbo.[@MORPHI].Code where u_mantisst='N'ORDER BY dbo.DocumentNoV1, dbo.RDR1.LineNum"
        oSalesRec.DoQuery(strsql)
        If oSalesRec.RecordCount > 0 Then
            Return strsql
        Else
            Return ""
        End If
        Return ""
    End Function
#End Region

#Region "Export Incoming Payment Doocuments"
    Private Sub ExportIncomingPaymentDocument()
        Dim oTempRec As SAPbobsCOM.Recordset
        oTempRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec.DoQuery("select DocEntry from ORCT Where isnull(U_Export,'N')='N'")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
            oTempRec.MoveNext()
        Next
    End Sub
#End Region

#Region "Export InComing Payment Details"
    Private Sub ExportIncomingPayment(ByVal aCompany As SAPbobsCOM.Company, ByVal aDocEntry As Integer)
        Dim objMainDoc, ObjRemoteDoc As SAPbobsCOM.Payments
        Dim objremoteRec As SAPbobsCOM.Recordset
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry from ORCT where DocEntry=" & aDocEntry)
        For intRemoteloop As Integer = 0 To objremoteRec.RecordCount - 1
            ObjRemoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            If ObjRemoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                objMainDoc.Address = ObjRemoteDoc.Address
                objMainDoc.ApplyVAT = ObjRemoteDoc.ApplyVAT
                objMainDoc.BankAccount = ObjRemoteDoc.BankAccount
                objMainDoc.BankChargeAmount = ObjRemoteDoc.BankChargeAmount
                objMainDoc.BankCode = ObjRemoteDoc.BankCode
                objMainDoc.BillOfExchangeAgent = ObjRemoteDoc.BillOfExchangeAgent
                objMainDoc.BillOfExchangeAgent = ObjRemoteDoc.BillOfExchangeAmount
                objMainDoc.BillofExchangeStatus = ObjRemoteDoc.BillofExchangeStatus
                objMainDoc.BoeAccount = ObjRemoteDoc.BoeAccount
                objMainDoc.CardCode = ObjRemoteDoc.CardCode
                objMainDoc.CardName = ObjRemoteDoc.CardName
                objMainDoc.CashAccount = ObjRemoteDoc.CashAccount
                objMainDoc.CashSum = ObjRemoteDoc.CashSum
                objMainDoc.CheckAccount = ObjRemoteDoc.CheckAccount

                If ObjRemoteDoc.Checks.CheckSum > 0 Then
                    For intCheck As Integer = 0 To ObjRemoteDoc.Checks.Count - 1
                        If intCheck > 0 Then
                            objMainDoc.Checks.Add()
                            objMainDoc.Checks.SetCurrentLine(intCheck)
                        End If
                        ObjRemoteDoc.Checks.SetCurrentLine(intCheck)
                        objMainDoc.Checks.AccounttNum = ObjRemoteDoc.Checks.AccounttNum
                        objMainDoc.Checks.BankCode = ObjRemoteDoc.Checks.BankCode
                        objMainDoc.Checks.Branch = ObjRemoteDoc.Checks.Branch
                        objMainDoc.Checks.CheckAccount = ObjRemoteDoc.Checks.CheckAccount
                        objMainDoc.Checks.CheckAccount = ObjRemoteDoc.Checks.CheckNumber
                        objMainDoc.Checks.CheckSum = ObjRemoteDoc.Checks.CheckSum
                        objMainDoc.Checks.CountryCode = ObjRemoteDoc.Checks.CountryCode
                        objMainDoc.Checks.Details = ObjRemoteDoc.Checks.Details
                        objMainDoc.Checks.DueDate = ObjRemoteDoc.Checks.DueDate
                        objMainDoc.Checks.ManualCheck = ObjRemoteDoc.Checks.ManualCheck
                        objMainDoc.Checks.Trnsfrable = ObjRemoteDoc.Checks.Trnsfrable
                        '    MsgBox(ObjRemoteDoc.Checks.CheckSum)
                    Next
                End If
                objMainDoc.ContactPersonCode = ObjRemoteDoc.ContactPersonCode
                objMainDoc.CounterReference = ObjRemoteDoc.CounterReference
                If ObjRemoteDoc.CreditCards.CreditSum > 0 Then
                    objMainDoc.CreditCards.AdditionalPaymentSum = ObjRemoteDoc.CreditCards.AdditionalPaymentSum
                    objMainDoc.CreditCards.CardValidUntil = ObjRemoteDoc.CreditCards.CardValidUntil
                    objMainDoc.CreditCards.ConfirmationNum = ObjRemoteDoc.CreditCards.ConfirmationNum
                    objMainDoc.CreditCards.CreditAcct = ObjRemoteDoc.CreditCards.CreditAcct
                    objMainDoc.CreditCards.CreditCard = ObjRemoteDoc.CreditCards.CreditCard
                    objMainDoc.CreditCards.CreditCardNumber = ObjRemoteDoc.CreditCards.CreditCardNumber
                    objMainDoc.CreditCards.CreditSum = ObjRemoteDoc.CreditCards.CreditSum
                    objMainDoc.CreditCards.CreditType = ObjRemoteDoc.CreditCards.CreditType
                    objMainDoc.CreditCards.FirstPaymentDue = ObjRemoteDoc.CreditCards.FirstPaymentDue
                    objMainDoc.CreditCards.FirstPaymentSum = ObjRemoteDoc.CreditCards.FirstPaymentSum
                    objMainDoc.CreditCards.NumOfCreditPayments = ObjRemoteDoc.CreditCards.NumOfCreditPayments
                    objMainDoc.CreditCards.NumOfPayments = ObjRemoteDoc.CreditCards.NumOfPayments
                    objMainDoc.CreditCards.OwnerIdNum = ObjRemoteDoc.CreditCards.OwnerIdNum
                    objMainDoc.CreditCards.OwnerPhone = ObjRemoteDoc.CreditCards.OwnerPhone
                    objMainDoc.CreditCards.SplitPayments = ObjRemoteDoc.CreditCards.SplitPayments
                    objMainDoc.CreditCards.PaymentMethodCode = ObjRemoteDoc.CreditCards.PaymentMethodCode
                End If
                objMainDoc.DeductionPercent = ObjRemoteDoc.DeductionPercent
                objMainDoc.DeductionSum = ObjRemoteDoc.DeductionSum
                objMainDoc.DocCurrency = ObjRemoteDoc.DocCurrency
                objMainDoc.DocRate = ObjRemoteDoc.DocRate
                objMainDoc.DocTypte = ObjRemoteDoc.DocTypte
                objMainDoc.DueDate = ObjRemoteDoc.DueDate
                objMainDoc.DocObjectCode = ObjRemoteDoc.DocObjectCode
                objMainDoc.HandWritten = ObjRemoteDoc.HandWritten
                For intInvoice As Integer = 0 To ObjRemoteDoc.Invoices.Count - 1
                    If ObjRemoteDoc.Invoices.DocEntry > 0 Then
                        If intInvoice > 0 Then
                            objMainDoc.Invoices.Add()
                            objMainDoc.Invoices.SetCurrentLine(intInvoice)
                        End If
                        ObjRemoteDoc.Invoices.SetCurrentLine(intInvoice)
                        objMainDoc.Invoices.AppliedFC = ObjRemoteDoc.Invoices.AppliedFC
                        objMainDoc.Invoices.DiscountPercent = ObjRemoteDoc.Invoices.DiscountPercent
                        objMainDoc.Invoices.DistributionRule = ObjRemoteDoc.Invoices.DistributionRule
                        objMainDoc.Invoices.DocEntry = ObjRemoteDoc.Invoices.DocEntry
                        objMainDoc.Invoices.DocLine = ObjRemoteDoc.Invoices.DocLine
                        objMainDoc.Invoices.InstallmentId = ObjRemoteDoc.Invoices.InstallmentId
                        objMainDoc.Invoices.InvoiceType = ObjRemoteDoc.Invoices.InvoiceType
                        objMainDoc.Invoices.SumApplied = ObjRemoteDoc.Invoices.SumApplied
                    End If
                Next
                objMainDoc.IsPayToBank = ObjRemoteDoc.IsPayToBank
                objMainDoc.JournalRemarks = ObjRemoteDoc.JournalRemarks
                objMainDoc.LocalCurrency = ObjRemoteDoc.LocalCurrency
                objMainDoc.PaymentPriority = ObjRemoteDoc.PaymentPriority
                objMainDoc.PaymentType = ObjRemoteDoc.PaymentType
                objMainDoc.PayToBankAccountNo = ObjRemoteDoc.PayToBankAccountNo
                objMainDoc.PayToBankBranch = ObjRemoteDoc.PayToBankBranch
                objMainDoc.PayToBankCode = ObjRemoteDoc.PayToBankCode
                objMainDoc.PayToBankCountry = ObjRemoteDoc.PayToBankCountry
                objMainDoc.ProjectCode = ObjRemoteDoc.ProjectCode
                objMainDoc.Reference1 = ObjRemoteDoc.Reference1
                objMainDoc.Reference2 = ObjRemoteDoc.Reference2
                objMainDoc.Remarks = ObjRemoteDoc.Remarks
                objMainDoc.Series = ObjRemoteDoc.Series
                objMainDoc.TaxDate = ObjRemoteDoc.TaxDate
                objMainDoc.TaxGroup = ObjRemoteDoc.TaxGroup
                objMainDoc.TransactionCode = ObjRemoteDoc.TransactionCode
                objMainDoc.TransferAccount = ObjRemoteDoc.TransferAccount
                objMainDoc.TransferRealAmount = ObjRemoteDoc.TransferRealAmount
                objMainDoc.TransferReference = ObjRemoteDoc.TransferReference
                objMainDoc.TransferSum = ObjRemoteDoc.TransferSum
                objMainDoc.VatDate = ObjRemoteDoc.VatDate
                objMainDoc.WTAmount = ObjRemoteDoc.WTAmount
                objMainDoc.WtBaseSum = ObjRemoteDoc.WtBaseSum
                objMainDoc.WTCode = ObjRemoteDoc.WTCode
                If ObjRemoteDoc.DocType = SAPbobsCOM.BoRcptTypes.rAccount Then
                    For intAcctLoop As Integer = 0 To ObjRemoteDoc.AccountPayments.Count - 1
                        If ObjRemoteDoc.AccountPayments.SumPaid > 0 Then
                            If intAcctLoop > 0 Then
                                objMainDoc.AccountPayments.Add()
                                objMainDoc.AccountPayments.SetCurrentLine(intAcctLoop)
                            End If
                            ObjRemoteDoc.AccountPayments.SetCurrentLine(intAcctLoop)
                            objMainDoc.AccountPayments.AccountCode = ObjRemoteDoc.AccountPayments.AccountCode
                            objMainDoc.AccountPayments.AccountName = ObjRemoteDoc.AccountPayments.AccountName
                            objMainDoc.AccountPayments.Decription = ObjRemoteDoc.AccountPayments.Decription
                            objMainDoc.AccountPayments.GrossAmount = ObjRemoteDoc.AccountPayments.GrossAmount
                            objMainDoc.AccountPayments.SumPaid = ObjRemoteDoc.AccountPayments.SumPaid
                            objMainDoc.AccountPayments.ProfitCenter = ObjRemoteDoc.AccountPayments.ProfitCenter
                            objMainDoc.AccountPayments.ProjectCode = ObjRemoteDoc.AccountPayments.ProjectCode
                            objMainDoc.AccountPayments.VatAmount = ObjRemoteDoc.AccountPayments.VatAmount
                            objMainDoc.AccountPayments.VatGroup = ObjRemoteDoc.AccountPayments.VatGroup
                        End If
                    Next
                End If
            End If
            If objMainDoc.Add() <> 0 Then
                sPath = strErrorLogPath
                WriteErrorlog("Failed to create Incomint payment  docuemnt :" & objMainCompany.GetLastErrorDescription, sPath)
            Else
                Dim strDocNumber As String
                objMainCompany.GetNewObjectCode(strDocNumber)
                'MsgBox(strDocNumber)
                sPath = strErrorLogPath
                WriteErrorlog("Incoming Payment document  Created Successfully. DocNum : " & strDocNumber, sPath)
                ObjRemoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                ObjRemoteDoc.Update()
            End If
            objremoteRec.MoveNext()
        Next

    End Sub

#End Region



#Region "Impot BP Master"
    Private Sub ImportBPMaster()
        Dim FileName, strItem As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.BusinessPartners
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select CardCode from OCRD ") 'where CardCode='test'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\BP.xml"
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                objMainItem.SaveXML(FileName)
                objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                If objRemoteItem.Add() <> 0 Then
                    WriteErrorlog("Failed to create BP : " & objRemoteCompany.GetLastErrorDescription, sPath)
                Else
                    objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                    objMainItem.Update()
                    WriteErrorlog("New BP Created : " & objRemoteItem.CardCode, sPath)
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub
    Private Sub UpdateBPMaster()
        Dim FileName, strItem As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.BusinessPartners
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select CardCode from OCRD where U_Action='U'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\BP.xml"
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                If objRemoteItem.GetByKey(strItem) Then
                    objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objMainItem.SaveXML(FileName)
                    objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                    If objRemoteItem.Update <> 0 Then
                        WriteErrorlog("Failed to update BP : " & objRemoteItem.CardCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                        WriteErrorlog("BP Updated : " & objRemoteItem.CardCode, sPath)
                    End If
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub
#End Region

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String, ByVal oRecSet As SAPbobsCOM.Recordset) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            oRecSet.DoQuery(strSQL)

            If Convert.ToString(oRecSet.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRecSet.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Export Stock transfer Request Details"
    Private Sub StocktransferRequest()
        Dim FileName, strItem, strSQLquery, strCode As String
        Dim Ecount As Long
        Dim ii, intDocEntry As Long
        strSQLquery = "SELECT T0.[DocEntry], T0.[U_DocDate], T0.[U_WhsCode], T1.[U_ItemCode], T1.[U_ItemName], T1.[U_Qty],T0.U_Status FROM [dbo].[@DABT_STRHEADER]  T0 "
        strSQLquery = strSQLquery & " inner join [dbo].[@DABT_STRLINES]  T1  on  (T0.[DocEntry] = T1.[DocEntry]  )  where t0.U_Status='O' and  isnull(T1.U_ItemCode,'')<>''"
        Dim objMainItem, objRemoteItem As SAPbobsCOM.UserTable
        Dim objMainRecset, objremoteRecSet As SAPbobsCOM.Recordset
        objremoteRecSet = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRecSet.DoQuery(strSQLquery)
        objMainRecset = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intItemLoop As Integer = 0 To objremoteRecSet.RecordCount - 1
            intDocEntry = objremoteRecSet.Fields.Item(0).Value
            objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objMainItem = objMainCompany.UserTables.Item("DABT_StImport")
            strCode = getMaxCode("@DABT_StImport", "Code", objMainRecset)
            objMainItem.Code = strCode
            objMainItem.Name = strCode
            objMainItem.UserFields.Fields.Item("U_TransDate").Value = objremoteRecSet.Fields.Item(1).Value
            objMainItem.UserFields.Fields.Item("U_TransWhs").Value = objremoteRecSet.Fields.Item(2).Value
            objMainItem.UserFields.Fields.Item("U_ItemCode").Value = objremoteRecSet.Fields.Item(3).Value
            objMainItem.UserFields.Fields.Item("U_ItemName").Value = objremoteRecSet.Fields.Item(4).Value
            objMainItem.UserFields.Fields.Item("U_ReqQty").Value = objremoteRecSet.Fields.Item(5).Value
            If objMainItem.Add <> 0 Then
                WriteErrorlog("Failed to import data " & objRemoteCompany.GetLastErrorDescription, sPath)
            Else
                objMainRecset = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objMainRecset.DoQuery("Update [@DABT_STRHEADER] set U_Status='I' where docentry=" & intDocEntry)
            End If
            objremoteRecSet.MoveNext()
        Next
    End Sub
#End Region

#Region "Impot Item Master"
    Private Sub ImportItemMaster()
        Dim FileName, strItem As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.Items
        Dim objMainRecset As SAPbobsCOM.Recordset
        '        Dim sboItem As SAPbobsCOM.Items = oCompany.GetBusinessObject(BoObjectTypes.oItems)
        '        If sboItem.GetByKey(itemcode) = False Then
        '            Exit Sub
        '        End If
        '        oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
        '        sboItem.ItemCode = itemcode + "_d"
        '        sboItem.SaveToFile(String.Format("c:\temp\{0}.xml", itemcode))

        '        Dim sboNewItem As SAPbobsCOM.Items = oCompany.GetBusinessObjectFromXML(String.Format("c:\temp\{0}.xml", itemcode), 0)

        '        If sboNewItem.Add()  0 Then
        '            sbo_application.MessageBox(oCompany.GetLastErrorCode.ToString + Space(1) + oCompany.GetLastErrorDescription)
        '        Else
        '            sbo_application.MessageBox("Item added")
        '        End If
        'Your code should be
        '        sboItem = oSBO.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '        _oSBO.SboCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
        '        sboItem.ItemCode = sItemCode
        '        sboItem.SaveXML(sFileName)
        '        sboNewItem = _oSBO.SboCompany.GetBusinessObjectFromXML(sFileName, 0)
        '        iRet = sboNewItem.Add()
        'If iRet 0 Then
        '            '_oSBO.SboCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        '            _oSBO.SboCompany.GetLastError(iRet, sErrMsg)
        '            _oSBO.SboCompany.MessageBox(sErrMsg & vbCrLf & "Failed to duplicate item record")
        '            Return ""
        '        Else
        '            Return sboNewItem.ItemCode
        '        End If
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select ItemCode from OITM where U_Action='A'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\Item.xml"
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                objMainItem.SaveXML(FileName)
                objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                If objRemoteItem.Add() <> 0 Then
                    WriteErrorlog("Failed to create item : " & objRemoteCompany.GetLastErrorDescription, sPath)
                Else
                    objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                    objMainItem.Update()
                    WriteErrorlog("New Item Created : " & objRemoteItem.ItemCode, sPath)
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub
    Private Sub UpdateItemMaster()
        Dim FileName, strItem As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.Items
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select ItemCode from OITM where U_Action='U'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\Item.xml"
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                If objRemoteItem.GetByKey(strItem) Then
                    objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objMainItem.SaveXML(FileName)
                    objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                    If objRemoteItem.Update <> 0 Then
                        WriteErrorlog("Failed to Update item : " & objRemoteItem.ItemCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                        WriteErrorlog("Update Completed : " & objRemoteItem.ItemCode, sPath)
                    End If
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub
#End Region

#Region "Export Item Details"
    Private Sub exportJournalVouchers(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.JournalVouchers
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        ImportJournalEntries()
    End Sub

    Private Function checkDuplicateJEs(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBranchcode As String, ByVal aTableName As String) As Boolean
        Dim oDuplicateCheck As SAPbobsCOM.Recordset
        Dim strQuery1 As String
        oDuplicateCheck = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strQuery1 = "Select * from " & aTableName & " where U_BaseEntry in (" & aDocEntry & ") and U_BaseNum in (" & aDocNum & ") and U_Branch='" & aBranchcode & "'"
        oDuplicateCheck.DoQuery(strQuery1)
        If oDuplicateCheck.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub ImportJournalEntries()
        Dim objMainJE, objRemoteJE As SAPbobsCOM.JournalEntries
        Dim objremoteRec, otemp As SAPbobsCOM.Recordset
        Dim strAcctCode, strFormatCode As String
        objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select * from OJDT where  isnull(U_JEExport,'N')='N'")
        For intLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objRemoteJE = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            If objRemoteJE.GetByKey(objremoteRec.Fields.Item("TransID").Value) Then
                If 1 = 2 Then 'checkDuplicateJEs(objRemoteJE.JdtNum, objRemoteJE.Number, objMainCompany.CompanyName, "OJDT") = True Then
                    WriteErrorlog("Journal Entry : " & objRemoteJE.Number & " already exported --> DataBase : " & objRemoteCompany.CompanyName, sPath)
                Else
                    objMainJE = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                    objMainJE.DueDate = objRemoteJE.DueDate
                    objMainJE.Memo = objRemoteJE.Memo
                    objMainJE.ProjectCode = objRemoteJE.ProjectCode
                    objMainJE.Reference = objRemoteJE.Reference
                    objMainJE.Reference2 = objRemoteJE.Reference2
                    objMainJE.ReferenceDate = objRemoteJE.ReferenceDate
                    objMainJE.TaxDate = objRemoteJE.TaxDate
                    objMainJE.TransactionCode = objRemoteJE.TransactionCode
                    objMainJE.VatDate = objRemoteJE.VatDate

                    ''  objMainJE.UserFields.Fields.Item("U_BaseEntry").Value = objRemoteJE.JdtNum
                    ' objMainJE.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName
                    ' objMainJE.UserFields.Fields.Item("U_BaseNum").Value = objRemoteJE.Number

                    'UpdateExchangerate_Main("", objRemoteJE.TaxDate, objRemoteCompany)
                    UpdateExchangerate("", objRemoteJE.TaxDate, objRemoteCompany)
                    For intLines As Integer = 0 To objRemoteJE.Lines.Count - 1
                        If intLines > 0 Then
                            objMainJE.Lines.Add()
                        End If
                        objMainJE.Lines.SetCurrentLine(intLines)
                        objRemoteJE.Lines.SetCurrentLine(intLines)

                        strFormatCode = getFormatCode(objRemoteJE.Lines.AccountCode, objMainCompany)
                        strAcctCode = getAccountCode(strFormatCode, objRemoteCompany)
                        objMainJE.Lines.AccountCode = strAcctCode
                        'objMainJE.Lines.AccountCode = objRemoteJE.Lines.AccountCode
                        objMainJE.Lines.AdditionalReference = objRemoteJE.Lines.AdditionalReference
                        objMainJE.Lines.BaseSum = objRemoteJE.Lines.BaseSum
                        objMainJE.Lines.ContraAccount = objRemoteJE.Lines.ContraAccount
                        objMainJE.Lines.CostingCode = objRemoteJE.Lines.CostingCode
                        If objRemoteJE.Lines.FCCurrency <> "" Then
                            objMainJE.Lines.FCCurrency = objRemoteJE.Lines.FCCurrency
                            ' UpdateExchangerate(objRemoteJE.Lines.FCCurrency, objRemoteJE.TaxDate, objRemoteCompany)
                            objMainJE.Lines.FCDebit = objRemoteJE.Lines.FCDebit
                            objMainJE.Lines.FCCredit = objRemoteJE.Lines.FCCredit
                        Else
                            objMainJE.Lines.Credit = objRemoteJE.Lines.Credit
                            objMainJE.Lines.Debit = objRemoteJE.Lines.Debit
                        End If
                        objMainJE.Lines.DueDate = objRemoteJE.Lines.DueDate
                        objMainJE.Lines.LineMemo = objRemoteJE.Lines.LineMemo
                        '  objMainJE.Lines.LocationCode = objRemoteJE.Lines.LocationCode
                        objMainJE.Lines.ProjectCode = objRemoteJE.Lines.ProjectCode
                        objMainJE.Lines.Reference1 = objRemoteJE.Lines.Reference1
                        objMainJE.Lines.Reference2 = objRemoteJE.Lines.Reference2
                        objMainJE.Lines.ReferenceDate1 = objRemoteJE.Lines.ReferenceDate1
                        objMainJE.Lines.ReferenceDate2 = objRemoteJE.Lines.ReferenceDate2
                        objMainJE.Lines.ShortName = objRemoteJE.Lines.ShortName
                        objMainJE.Lines.TaxCode = objRemoteJE.Lines.TaxCode
                        objMainJE.Lines.TaxDate = objRemoteJE.Lines.TaxDate
                        objMainJE.Lines.TaxGroup = objRemoteJE.Lines.TaxGroup
                    Next

                    If objMainJE.Add <> 0 Then
                        WriteErrorlog("Failed to create Journal Entry  " & objRemoteJE.Number & " : Database : " & objMainCompany.CompanyName & " :Error : " & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        Dim strDocNum As String
                        objMainCompany.GetNewObjectCode(strDocNum)
                        WriteErrorlog("Journal entry created successully " & strDocNum & " : Database : " & objRemoteCompany.CompanyName, sPath)
                        objRemoteJE.UserFields.Fields.Item("U_JEExport").Value = "Y"
                        objRemoteJE.Update()

                    End If
                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub
    Private Function getFormatCode(ByVal aCode As String, ByVal aCompany As SAPbobsCOM.Company) As String
        Dim otest As SAPbobsCOM.Recordset
        otest = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select Formatcode from OACT where AcctCode='" & aCode & "'")
        Return otest.Fields.Item(0).Value
    End Function

    Private Function getAccountCode(ByVal aCode As String, ByVal aCompany As SAPbobsCOM.Company) As String
        Dim otest As SAPbobsCOM.Recordset
        otest = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select AcctCode from OACT where FormatCode='" & aCode & "'")
        Return otest.Fields.Item(0).Value
    End Function
    Private Sub UpdateExchangerate(ByVal strCurrency As String, ByVal dtdate As Date, ByVal aCompany As SAPbobsCOM.Company)
        Dim oexchrate, oExchnageMain As SAPbobsCOM.Recordset
        Dim dblRate, dblRate1 As Double
        Dim strSQL As String
        oExchnageMain = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oExchnageMain.DoQuery("Select Currency from ORTT where convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "' and isnull(rate,0)<>0")
        For intLoop As Integer = 0 To oExchnageMain.RecordCount - 1
            strCurrency = oExchnageMain.Fields.Item(0).Value
            If strCurrency <> "" Then
                oexchrate = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = "Select * from ORTT where Currency='" & strCurrency & "' and convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "'"
                oexchrate.DoQuery(strSQL)
                If oexchrate.RecordCount > 0 Then
                    dblRate1 = oexchrate.Fields.Item("Rate").Value
                End If
                oexchrate = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oexchrate.DoQuery("Select * from ORTT where Currency='" & strCurrency & "' and convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "'")
                If oexchrate.RecordCount > 0 Then
                    dblRate = oexchrate.Fields.Item("Rate").Value
                    If dblRate <= 0 Or dblRate <> dblRate1 Then
                        oexchrate = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oexchrate.DoQuery("Update ORTT set Rate=" & dblRate1 & " where Currency='" & strCurrency & "' and  convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "'")
                    End If
                Else
                    oexchrate = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strSQL = "insert into ORTT values ('" & dtdate.ToString("yyyy-MM-dd") & "','" & strCurrency & "'," & dblRate1 & ",'I',1)"
                    oexchrate.DoQuery(strSQL)
                End If
            End If
            oExchnageMain.MoveNext()
        Next
    End Sub


    Private Sub UpdateExchangerate_Main(ByVal strCurrency As String, ByVal dtdate As Date, ByVal aCompany As SAPbobsCOM.Company)
        Dim oexchrate, oexchrate1 As SAPbobsCOM.Recordset
        Dim dblRate, dblRate1 As Double
        Dim strSQL As String
        oexchrate = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "Select * from ORTT where convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "' and Rate<>0"
        oexchrate.DoQuery(strSQL)
        For intRow As Integer = 0 To oexchrate.RecordCount - 1
            strCurrency = oexchrate.Fields.Item("Currency").Value
            dblRate1 = oexchrate.Fields.Item("Rate").Value
            oexchrate1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oexchrate1.DoQuery("Select * from ORTT where Currency='" & strCurrency & "' and convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "'")
            If oexchrate1.RecordCount > 0 Then
                dblRate = oexchrate1.Fields.Item("Rate").Value
                If dblRate <= 0 Then
                    oexchrate = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oexchrate.DoQuery("Update ORTT set Rate=" & dblRate1 & " where Currency='" & strCurrency & "' and rateDate=" & dtdate)
                End If
            Else
                oexchrate1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = "insert into ORTT values ('" & dtdate.ToString("yyyy-MM-dd") & "','" & strCurrency & "'," & dblRate1 & ",'I',1)"
                oexchrate.DoQuery(strSQL)
            End If

            oexchrate.MoveNext()
        Next
    End Sub
    Private Sub ExportSalesOrer(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry from OINV where IsICT='N' and  isnull(U_Export,'N')='N'")
        'Header fields � customer code, customer name, posting date, due date, document date, customer reference number, remarks, Total before discount, discount %, Tax, Total, Applied Amount, Balance Due
        'Line fields � item code, item name, barcode, quantity, unit price, discount %, price after discount, vat code, Gross price, Total (LC)
        For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objremoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                objMainDoc.Address = objremoteDoc.Address
                objMainDoc.Address2 = objremoteDoc.Address2
                objMainDoc.CardCode = objremoteDoc.CardCode
                objMainDoc.CardName = objremoteDoc.CardName
                objMainDoc.CentralBankIndicator = objremoteDoc.CentralBankIndicator
                objMainDoc.ClosingRemarks = objremoteDoc.ClosingRemarks
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                objMainDoc.DeferredTax = objremoteDoc.DeferredTax
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.DocDate = objremoteDoc.DocDate
                objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                objMainDoc.DocRate = objremoteDoc.DocRate
                objMainDoc.DocTotal = objremoteDoc.DocTotal
                objMainDoc.DocType = objremoteDoc.DocType
                objMainDoc.DocumentSubType = objremoteDoc.DocumentSubType
                objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                objMainDoc.TaxDate = objremoteDoc.TaxDate
                objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                'objMainDoc.TransportationCode = objremoteDoc.TransportationCode

                If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                Else
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                    If objremoteDoc.Expenses.LineTotal > 0 Then
                        If IntExp > 0 Then
                            objMainDoc.Expenses.Add()
                            objMainDoc.Expenses.SetCurrentLine(IntExp)
                        End If
                        objremoteDoc.Expenses.SetCurrentLine(IntExp)
                        objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                        objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                        objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                        objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                        objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                        objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                        objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                        objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                        objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                        objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                        objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                    End If
                Next

                For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                    If intLoop > 0 Then
                        objMainDoc.Lines.Add()
                        objMainDoc.Lines.SetCurrentLine(intLoop)
                    End If
                    objremoteDoc.Lines.SetCurrentLine(intLoop)
                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                    objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                Next



                If objMainDoc.Add <> 0 Then
                    sPath = strErrorLogPath
                    WriteErrorlog("Failed to create invoice docuemnt :" & objremoteDoc.DocNum & " Error : " & objMainCompany.GetLastErrorDescription, sPath)
                    '    MsgBox(objMainCompany.GetLastErrorDescription)
                Else
                    Dim strDocNum As String
                    objMainCompany.GetNewObjectCode(strDocNum)
                    sPath = strErrorLogPath
                    WriteErrorlog("Invoice Created Successfully. DocNum : " & strDocNum, sPath)
                    objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                    objremoteDoc.Update()
                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub

    Private Sub ExportInvoicePayment(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim intRemoteDocEntry As Integer
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'objremoteRec.DoQuery("Select DocEntry from OINV where IsICT='Y' and  DocType='I' and  isnull(U_Export,'N')='N'")
        objremoteRec.DoQuery("Select DocEntry from OINV where IsICT='Y' and  isnull(U_Export,'N')='N'")
        For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objremoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                intRemoteDocEntry = objremoteDoc.DocEntry
                objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                objMainDoc.CardCode = objremoteDoc.CardCode
                objMainDoc.DocDate = objremoteDoc.DocDate
                objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                objMainDoc.TaxDate = objremoteDoc.TaxDate
                objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                objMainDoc.Address = objremoteDoc.Address
                objMainDoc.Address2 = objremoteDoc.Address2
                objMainDoc.AgentCode = objremoteDoc.AgentCode
                objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                objMainDoc.DocRate = objremoteDoc.DocRate
                objMainDoc.DocType = objremoteDoc.DocType
                '  objMainDoc.TransportationCode = objremoteDoc.TransportationCode
                If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                Else
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                    If objremoteDoc.Expenses.LineTotal > 0 Then
                        If IntExp > 0 Then
                            objMainDoc.Expenses.Add()
                            objMainDoc.Expenses.SetCurrentLine(IntExp)
                        End If
                        objremoteDoc.Expenses.SetCurrentLine(IntExp)
                        objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                        objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                        objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                        objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                        objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                        objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                        objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                        objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                        objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                        objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                        objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                    End If
                Next

                For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                    If intLoop > 0 Then
                        objMainDoc.Lines.Add()
                        objMainDoc.Lines.SetCurrentLine(intLoop)
                    End If
                    objremoteDoc.Lines.SetCurrentLine(intLoop)
                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                    objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                Next
                If objMainDoc.Add <> 0 Then
                    sPath = strErrorLogPath
                    WriteErrorlog("Failed to create invoice+Payment docuemnt :" & objMainCompany.GetLastErrorDescription, sPath)
                Else
                    Dim strDocNum As String
                    objMainCompany.GetNewObjectCode(strDocNum)
                    sPath = strErrorLogPath
                    WriteErrorlog("Invoice + Payment Created Successfully. DocNum : " & strDocNum, sPath)
                    Dim otempRec As SAPbobsCOM.Recordset
                    otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    otempRec.DoQuery("Update OINV set IsICT='Y' where Docentry=" & strDocNum)
                    intRemoteDocEntry = objremoteDoc.DocEntry
                    ExportIncomingPayment(oCompany, GetIncomingDocument(intRemoteDocEntry))
                    objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                    objremoteDoc.Update()
                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub

    Private Sub ImportCreditNotes(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'objremoteRec.DoQuery("Select DocEntry from ORIN where  DocType='I' and  isnull(U_Export,'N')='N'")
        objremoteRec.DoQuery("Select DocEntry from ORIN where   isnull(U_Export,'N')='N'")

        'Header fields � customer code, customer name, posting date, due date, document date, customer reference number, remarks, Total before discount, discount %, Tax, Total, Applied Amount, Balance Due
        'Line fields � item code, item name, barcode, quantity, unit price, discount %, price after discount, vat code, Gross price, Total (LC)
        For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objremoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                objMainDoc.CardCode = objremoteDoc.CardCode
                objMainDoc.DocDate = objremoteDoc.DocDate
                objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                objMainDoc.TaxDate = objremoteDoc.TaxDate
                objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                objMainDoc.Address = objremoteDoc.Address
                objMainDoc.Address2 = objremoteDoc.Address2
                objMainDoc.AgentCode = objremoteDoc.AgentCode
                objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                objMainDoc.DocRate = objremoteDoc.DocRate
                objMainDoc.DocType = objremoteDoc.DocType
                '  objMainDoc.TransportationCode = objremoteDoc.TransportationCode
                If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                Else
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If

                'Document_LinesClass' 
                objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                    If objremoteDoc.Expenses.LineTotal > 0 Then
                        If IntExp > 0 Then
                            objMainDoc.Expenses.Add()
                            objMainDoc.Expenses.SetCurrentLine(IntExp)
                        End If
                        objremoteDoc.Expenses.SetCurrentLine(IntExp)
                        objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                        objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                        objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                        objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                        objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                        objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                        objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                        objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                        objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                        objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                        objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                    End If
                Next

                For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                    If intLoop > 0 Then
                        objMainDoc.Lines.Add()
                        objMainDoc.Lines.SetCurrentLine(intLoop)
                    End If
                    objremoteDoc.Lines.SetCurrentLine(intLoop)
                    '    objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                    objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                Next

                If objMainDoc.Add <> 0 Then
                    sPath = strErrorLogPath
                    WriteErrorlog("Failed to create Credit Notes docuemnt :" & objMainCompany.GetLastErrorDescription, sPath)
                Else
                    Dim strDocNum As String
                    objMainCompany.GetNewObjectCode(strDocNum)
                    'MsgBox("DOcument created :" & strDocNum)
                    sPath = strErrorLogPath
                    WriteErrorlog("Credit Notes Created Successfully. DocNum : " & strDocNum, sPath)
                    objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                    objremoteDoc.Update()
                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub
#End Region

    Private Sub CopyFilestoCustomers(ByVal aFileName As String, ByVal aLogPath As String)
        Dim otemp As SAPbobsCOM.Recordset
        Dim strFilePath, strDesgfilename, strMessage As String
        otemp = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select * from OCRD where cardtype='C' and U_PharmInt = 'Y'")
        strFilePath = strExportFilePaty
        For intRow As Integer = 0 To otemp.RecordCount - 1
            strFilePath = strFilePath & "\" & otemp.Fields.Item("CardCode").Value
            If Directory.Exists(strFilePath) Then
            Else
                Directory.CreateDirectory(strFilePath)
            End If
            strDesgfilename = strFilePath & "\PROMFLQ.mfp"
            If File.Exists(strDesgfilename) Then
                File.Delete(strDesgfilename)
            End If
            File.Copy(aFileName, strDesgfilename)
            strFilePath = strExportFilePaty
            strMessage = "Exported :  File name : " & strDesgfilename
            WriteErrorlog(strMessage, aLogPath)
            otemp.MoveNext()
        Next


    End Sub

#Region "Connect MainCompany"
    Private Function connectMainCompany() As Boolean
        sPath = strErrorLogPath
        Try
            If cmbMainSAPCompany.Text = "" Then
                MsgBox("Choose Company")
                Return False
            Else
                strMainSAPCompany = cmbMainSAPCompany.Text
            End If
            If MainSAPUID.Text = "" Then
                MsgBox("Enter UserName")
                Return False
            End If
            If MainSAPPWD.Text = "" Then
                MsgBox("Enter Password")
                Return False
            End If
            sPath = strErrorLogPath
            oCompany = New SAPbobsCOM.Company
            oCompany.Server = MainSQLServer.Text
            oCompany.LicenseServer = mainLicenseServer.Text ' "Compaq-PC:30000"
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            ' oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            If strMainServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Else
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            End If
            oCompany.DbUserName = MainSQLUID.Text
            oCompany.DbPassword = MainSQLPWD.Text
            Dim oRecs As SAPbobsCOM.Recordset
            oRecs = oCompany.GetCompanyList()
            WriteErrorlog("Connection to Main SAP B1 started", sPath)
            ' MsgBox(oRecs.RecordCount)
            If oCompany.Connect <> 0 Then
                ' Return False
            Else
                'MsgBox("Connected")
            End If
            If oCompany.Connected Then
                'MsgBox("Company connected")
            End If
            oCompany.CompanyDB = cmbMainSAPCompany.Text
            oCompany.UserName = MainSAPUID.Text
            oCompany.Password = MainSAPPWD.Text
            If oCompany.Connected = True Then
                If (objPC.Initialise(oCompany)) Then
                Else
                    MsgBox("Error in Connection")
                End If
                WriteiniFile()
                objMainCompany = oCompany
                Return True
            Else
                If oCompany.Connect <> 0 Then
                    'MsgBox("Connection to Main SAP B1 failed. Error Description :" & oCompany.GetLastErrorDescription)
                    WriteErrorlog("Connection to Main SAP B1 failed. Error Description :" & oCompany.GetLastErrorDescription, sPath)
                    Return False
                Else
                    WriteiniFile()
                    'MsgBox("Main SAP B1 company Connected successfully")
                    WriteErrorlog("Main SAP B1 Company connected", sPath)
                    objMainCompany = oCompany
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            WriteErrorlog(ex.Message, sPath)
            Return False
        End Try
        Return True
    End Function

    Private Function connectLocalCompany() As Boolean
        sPath = strErrorLogPath
        Try
            If cmbCompanyDB.Text = "" Then
                MsgBox("Choose Company")
                Return False
            Else
                strSAPCompany = cmbCompanyDB.Text
            End If
            If txtServerName.Text = "" Then
                MsgBox("Enter UserName")
                Return False
            End If
            If txtServerPwd.Text = "" Then
                MsgBox("Enter Password")
                Return False
            End If
            oCompany = New SAPbobsCOM.Company
            oCompany.Server = txtServerName.Text
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            If strLocalServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            ElseIf strLocalServertype = "2008" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf strLocalServertype = "2012" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012

            End If

            oCompany.LicenseServer = LocalLicenseServer.Text ' "Compaq-PC:30000"
            oCompany.DbUserName = txtDBUserName.Text
            oCompany.DbPassword = txtServerPwd.Text
            oCompany.CompanyDB = cmbCompanyDB.Text
            oCompany.UserName = txtSBOUserName.Text
            oCompany.Password = txtCompanyPWD.Text



            WriteErrorlog("Connection to local SAP B1 server started", sPath)
            If oCompany.Connected = True Then
                If (objPC.Initialise(oCompany)) Then
                Else
                    'MsgBox("Error in Connection")
                    WriteErrorlog("Error in Connection to Local Server", sPath)
                    Return False
                End If
                WriteiniFile()
                objRemoteCompany = oCompany
                Return True
            Else
                If oCompany.Connect <> 0 Then
                    'MsgBox("Connection to SAP B1 failed. Error Description :" & oCompany.GetLastErrorDescription)
                    WriteErrorlog("Connection to SAP B1 failed. Error Description :" & oCompany.GetLastErrorDescription, sPath)
                    Return False
                Else
                    WriteiniFile()
                    '     MsgBox("Local SAP B1 Company Connected successfully")
                    WriteErrorlog("Local SAP B1 Company Connected successfully", sPath)
                    objRemoteCompany = oCompany
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            WriteErrorlog(ex.Message, sPath)
            Return False
        End Try
        Return True
    End Function
#End Region

    Private Sub btnConnect_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Try
            If connectLocalCompany() = False Then
                MsgBox("Error in connection to local server")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub btnBrowse_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        fldFolderBrowse.ShowDialog()
        txtDirectory.Text = fldFolderBrowse.SelectedPath
    End Sub

#End Region

    Private Sub cmbCompanyDB_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCompanyDB.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If txtServerName.Text = "" Then
                MsgBox("Enter Server")
                Exit Sub
            End If
            If txtDBUserName.Text = "" Then
                MsgBox("Enter UserName")
                Exit Sub
            End If
            If txtServerPwd.Text = "" Then
                MsgBox("Enter Password")
                Exit Sub
            End If
            objMainCompany = New SAPbobsCOM.Company
            objMainCompany.Server = MainSQLServer.Text
            objMainCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            strMainServertype = cmbMainServertype.Text
            If strMainServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Else
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            End If
            objMainCompany.DbUserName = MainSQLUID.Text
            objMainCompany.DbPassword = MainSQLPWD.Text
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim lErrCode As Long
            Dim sErrMsg As String
            oRecordSet = objMainCompany.GetCompanyList
            objMainCompany.GetLastError(lErrCode, sErrMsg)
            If lErrCode <> 0 Then
                MsgBox(sErrMsg)
                Exit Sub
            Else
                cmbCompanyDB.Items.Clear()
                Do Until oRecordSet.EoF = True
                    cmbMainSAPCompany.Items.Add(oRecordSet.Fields.Item(0).Value)
                    oRecordSet.MoveNext()
                Loop
                btnConnect.Enabled = True
            End If
        Catch ex As Exception
            MsgBox("Connection to Server =" & txtServerName.Text & ".  Failed")
            Exit Sub
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If connectMainCompany() = False Then
            MsgBox("Error in connection to main server")
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        End
    End Sub

#Region "Get Payment Document Number"
    Private Function GetIncomingDocument(ByVal aDocEntry As Integer) As Integer
        Dim intPaymentdocentry As Integer
        Dim oRecset As SAPbobsCOM.Recordset
        oRecset = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset.DoQuery("Select DocNum from RCT2 where DocEntry=" & aDocEntry)
        If oRecset.RecordCount > 0 Then
            intPaymentdocentry = oRecset.Fields.Item(0).Value
        Else
            intPaymentdocentry = 0
        End If
        Return intPaymentdocentry
    End Function
#End Region

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox("Export process started")
        If connectLocalCompany() Then
            If connectMainCompany() Then
                sPath = strErrorLogPath
                WriteErrorlog("Export Process Started", sPath)
                Me.Hide()
                Me.Hide()
                WriteErrorlog("Export Invoice Document ", sPath)
                ExportSalesOrer(oCompany)
                WriteErrorlog("Export Invoice+Payment Document ", sPath)
                ExportInvoicePayment(oCompany)
                WriteErrorlog("Export Credit Note Document ", sPath)
                ImportCreditNotes(oCompany)
                WriteErrorlog("Export Incoming Payment Document ", sPath)
                ExportIncomingPaymentDocument()
                WriteErrorlog("Import Master Data Started", sPath)
                ImportItemMaster()
                UpdateItemMaster()
                ImportBPMaster()
                UpdateBPMaster()
                WriteErrorlog("Import Master Data completed", sPath)
                WriteErrorlog("Export Stock Transfer Request started..", sPath)
                StocktransferRequest()
                WriteErrorlog("Export Stock Transfer Request Completed..", sPath)
                Me.Hide()
                Me.Hide()
                End
            End If
        End If
        MsgBox("Export process Compleated")
        End
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        fldFolderBrowse.ShowDialog()
        txtDirectory.Text = fldFolderBrowse.SelectedPath
    End Sub

End Class
