Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

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
    Friend WithEvents txtCode1 As System.Windows.Forms.TextBox
    Friend WithEvents txtCode2 As System.Windows.Forms.TextBox
    Friend WithEvents txtCode3 As System.Windows.Forms.TextBox
    Friend WithEvents txtCode4 As System.Windows.Forms.TextBox
    Friend WithEvents txtCode5 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblVersion As Label
    Friend WithEvents Button1 As Button

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnTest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnTest = New System.Windows.Forms.Button()
        Me.txtCode1 = New System.Windows.Forms.TextBox()
        Me.txtCode2 = New System.Windows.Forms.TextBox()
        Me.txtCode3 = New System.Windows.Forms.TextBox()
        Me.txtCode4 = New System.Windows.Forms.TextBox()
        Me.txtCode5 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblVersion = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnTest
        '
        Me.btnTest.Location = New System.Drawing.Point(204, 220)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(75, 23)
        Me.btnTest.TabIndex = 0
        Me.btnTest.Text = "PRINT"
        '
        'txtCode1
        '
        Me.txtCode1.Location = New System.Drawing.Point(13, 50)
        Me.txtCode1.Name = "txtCode1"
        Me.txtCode1.Size = New System.Drawing.Size(267, 20)
        Me.txtCode1.TabIndex = 1
        Me.txtCode1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCode2
        '
        Me.txtCode2.Location = New System.Drawing.Point(12, 85)
        Me.txtCode2.Name = "txtCode2"
        Me.txtCode2.Size = New System.Drawing.Size(267, 20)
        Me.txtCode2.TabIndex = 2
        Me.txtCode2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCode3
        '
        Me.txtCode3.Location = New System.Drawing.Point(12, 120)
        Me.txtCode3.Name = "txtCode3"
        Me.txtCode3.Size = New System.Drawing.Size(267, 20)
        Me.txtCode3.TabIndex = 3
        Me.txtCode3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCode4
        '
        Me.txtCode4.Location = New System.Drawing.Point(13, 156)
        Me.txtCode4.Name = "txtCode4"
        Me.txtCode4.Size = New System.Drawing.Size(267, 20)
        Me.txtCode4.TabIndex = 4
        Me.txtCode4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtCode5
        '
        Me.txtCode5.Location = New System.Drawing.Point(13, 194)
        Me.txtCode5.Name = "txtCode5"
        Me.txtCode5.Size = New System.Drawing.Size(267, 20)
        Me.txtCode5.TabIndex = 5
        Me.txtCode5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Univers", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(11, 247)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 10)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Version:"
        '
        'lblVersion
        '
        Me.lblVersion.AutoSize = True
        Me.lblVersion.Font = New System.Drawing.Font("Univers", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.Location = New System.Drawing.Point(45, 247)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(29, 10)
        Me.lblVersion.TabIndex = 7
        Me.lblVersion.Text = "Label2"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(136, 247)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lblVersion)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCode5)
        Me.Controls.Add(Me.txtCode4)
        Me.Controls.Add(Me.txtCode3)
        Me.Controls.Add(Me.txtCode2)
        Me.Controls.Add(Me.txtCode1)
        Me.Controls.Add(Me.btnTest)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Unicomp - Overpack Box"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Dim serialnumber(5) As String

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click
        PrintLabel()

    End Sub
    Private Sub PrintLabels()

        Dim LW As Object
        Try
            LW = CreateObject("Lworks3.LabelEngine")

            'Open the label file we want to print

            'LW.FileName = "C:\Unicomp\Templates\OVERPACK.lw3"
            LW.FileName = "c:\unicomp\OVERPACK2.lw3"
            'LW.FileName = "\\uniwfs1\Share\OVERPACK-BOX-2016\template\OVERPACK2.lw3"
            'Set up the label print job.

            LW.Copies = 1
            LW.StartLabel = 1
            LW.TotalLabels = 1
            LW.UpdateSerials = False

            'Run the print job

            'If Not NoPrint Then LW.PrintLabels()

            'Close down LabelWorks
            LW.PrintLabels()
            LW = Nothing

        Catch ex As System.Runtime.InteropServices.COMException
            MessageBox.Show("Error: " & "OVERPACK2.lw3" & " not found" & vbCrLf & "in OVERPACK2.lw3", "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub CreateLabelData()
        Dim sDump As String
        sDump = Nothing

        Dim LB1 As String = txtCode1.Text.ToUpper
        Dim LB2 As String = txtCode2.Text.ToUpper
        Dim LB3 As String = txtCode3.Text.ToUpper
        Dim LB4 As String = txtCode4.Text.ToUpper
        Dim LB5 As String = txtCode5.Text.ToUpper


        Dim sPartNumber As String = Strings.Left(txtCode1.Text, 7)

        Try
            'FileOpen(1, "C:\Unicomp\Templates\OVERPACK.csv", OpenMode.Output)
            FileOpen(1, "c:\unicomp\OVERPACK2.csv", OpenMode.Output)

            PrintLine(1, "BC1,BC2,BC3,BC4,BC5,PART,PART_1,SERIAL_1,PART_2,SERIAL_2,PART_3,SERIAL_3,PART_4,SERIAL_4,PART_5,SERIAL_5")
            'lstOutput.Items.Add("OEMCust,UniPN,OEMPN,SerialStart,SerialEnd,QTY,WorkOrder,PrintReprint,Date")
            sDump = sDump &
            LB1 & "," &
            LB2 & "," &
            LB3 & "," &
            LB4 & "," &
            LB5 & "," &
            sPartNumber


            PrintLine(1, sDump)
            'lstOutput.Items.Add(sDump)
        Catch ex As Exception
            MessageBox.Show("Problem with label date:" & vbCrLf & ex.ToString, "LABEL DATA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

        sDump = Nothing

        FileClose(1)
    End Sub

    Private Sub txtCode1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode1.KeyPress
        Dim s As String = txtCode1.Text
        s.ToUpper()
        txtCode1.Text = s
        If e.KeyChar = Chr(13) Then
            txtCode2.Focus()
        End If
    End Sub

    Private Sub txtCode2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode2.KeyPress
        If e.KeyChar = Chr(13) Then
            txtCode3.Focus()
        End If
    End Sub

    Private Sub txtCode3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode3.KeyPress
        If e.KeyChar = Chr(13) Then
            txtCode4.Focus()
        End If
    End Sub

    Private Sub txtCode4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode4.KeyPress
        If e.KeyChar = Chr(13) Then
            txtCode5.Focus()
        End If
    End Sub

    Private Sub txtCode5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCode5.KeyPress
        If e.KeyChar = Chr(13) Then
            PrintLabel()

        End If
    End Sub

    Private Sub PrintLabel()
        Call CreateLabelData()
        Call PrintLabels()
        Call ClearTexts()
    End Sub

    Private Sub ClearTexts()
        txtCode1.Clear()
        txtCode2.Clear()
        txtCode3.Clear()
        txtCode4.Clear()
        txtCode5.Clear()
        txtCode1.Focus()
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Call ClearTexts()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '  Dim i As New MiscInfo.Info()
        lblVersion.Text = Me.GetType.Assembly.GetName.Version.ToString
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim startInfo As New ProcessStartInfo()
        Dim myprocess As New Process()
        startInfo.FileName = "Notepad"
        startInfo.Verb = "runas"
        ' startInfo.Arguments = "/env /user:" + "Administrator" + " cmd"
        myprocess.StartInfo = startInfo
        myprocess.Start()
    End Sub
End Class
