Public Class Form1
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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnTest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnTest = New System.Windows.Forms.Button
        Me.txtCode1 = New System.Windows.Forms.TextBox
        Me.txtCode2 = New System.Windows.Forms.TextBox
        Me.txtCode3 = New System.Windows.Forms.TextBox
        Me.txtCode4 = New System.Windows.Forms.TextBox
        Me.txtCode5 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnTest
        '
        Me.btnTest.Location = New System.Drawing.Point(205, 231)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(75, 23)
        Me.btnTest.TabIndex = 0
        Me.btnTest.Text = "TEST"
        '
        'txtCode1
        '
        Me.txtCode1.Location = New System.Drawing.Point(13, 50)
        Me.txtCode1.Name = "txtCode1"
        Me.txtCode1.Size = New System.Drawing.Size(267, 20)
        Me.txtCode1.TabIndex = 1
        '
        'txtCode2
        '
        Me.txtCode2.Location = New System.Drawing.Point(12, 85)
        Me.txtCode2.Name = "txtCode2"
        Me.txtCode2.Size = New System.Drawing.Size(267, 20)
        Me.txtCode2.TabIndex = 2
        '
        'txtCode3
        '
        Me.txtCode3.Location = New System.Drawing.Point(12, 120)
        Me.txtCode3.Name = "txtCode3"
        Me.txtCode3.Size = New System.Drawing.Size(267, 20)
        Me.txtCode3.TabIndex = 3
        '
        'txtCode4
        '
        Me.txtCode4.Location = New System.Drawing.Point(13, 156)
        Me.txtCode4.Name = "txtCode4"
        Me.txtCode4.Size = New System.Drawing.Size(267, 20)
        Me.txtCode4.TabIndex = 4
        '
        'txtCode5
        '
        Me.txtCode5.Location = New System.Drawing.Point(13, 194)
        Me.txtCode5.Name = "txtCode5"
        Me.txtCode5.Size = New System.Drawing.Size(267, 20)
        Me.txtCode5.TabIndex = 5
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.txtCode5)
        Me.Controls.Add(Me.txtCode4)
        Me.Controls.Add(Me.txtCode3)
        Me.Controls.Add(Me.txtCode2)
        Me.Controls.Add(Me.txtCode1)
        Me.Controls.Add(Me.btnTest)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click
        Call CreateLabelData()
        Call PrintLabels()
    End Sub

    Private Sub PrintLabels()

        Dim LW As Object
        Try
            LW = CreateObject("Lworks3.LabelEngine")

            'Open the label file we want to print

            LW.FileName = "C:\Unicomp\Templates\OVERPACK.lw3"

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
            MessageBox.Show("Error: " & ".LW3" & " not found" & vbCrLf & "in C:\Unicomp\Templates\", "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub CreateLabelData()
        Dim sDump As String
        sDump = Nothing
        Try
            FileOpen(1, "C:\Unicomp\Templates\OVERPACK.csv", OpenMode.Output)

            PrintLine(1, "BC1,BC2,BC3,BC4,BC5")
            'lstOutput.Items.Add("OEMCust,UniPN,OEMPN,SerialStart,SerialEnd,QTY,WorkOrder,PrintReprint,Date")
            sDump = sDump & _
            txtCode1.Text & "," & _
            txtCode2.Text & "," & _
            txtCode3.Text & "," & _
            txtCode4.Text & "," & _
            txtCode5.Text

            PrintLine(1, sDump)
            'lstOutput.Items.Add(sDump)
        Catch ex As Exception
            MessageBox.Show("Problem with label date:" & vbCrLf & ex.ToString, "LABEL DATA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

        sDump = Nothing

        FileClose(1)
    End Sub
End Class
