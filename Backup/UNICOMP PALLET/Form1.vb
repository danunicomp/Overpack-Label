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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnTest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnTest = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnTest
        '
        Me.btnTest.Location = New System.Drawing.Point(96, 184)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.TabIndex = 0
        Me.btnTest.Text = "TEST"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.btnTest)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click

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

            LW = Nothing
        Catch ex As System.Runtime.InteropServices.COMException
            MessageBox.Show("Error: " & ".LW3" & " not found" & vbCrLf & "in C:\Unicomp\Templates\", "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

End Class
