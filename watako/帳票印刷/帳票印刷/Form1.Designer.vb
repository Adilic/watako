<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.excelInput = New System.Windows.Forms.Button()
        Me.excelInputText = New System.Windows.Forms.TextBox()
        Me.printFile = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'excelInput
        '
        Me.excelInput.Location = New System.Drawing.Point(12, 30)
        Me.excelInput.Name = "excelInput"
        Me.excelInput.Size = New System.Drawing.Size(75, 23)
        Me.excelInput.TabIndex = 0
        Me.excelInput.Text = "インプット"
        Me.excelInput.UseVisualStyleBackColor = True
        '
        'excelInputText
        '
        Me.excelInputText.Location = New System.Drawing.Point(93, 32)
        Me.excelInputText.Name = "excelInputText"
        Me.excelInputText.Size = New System.Drawing.Size(286, 19)
        Me.excelInputText.TabIndex = 1
        '
        'printFile
        '
        Me.printFile.Location = New System.Drawing.Point(281, 119)
        Me.printFile.Name = "printFile"
        Me.printFile.Size = New System.Drawing.Size(98, 36)
        Me.printFile.TabIndex = 4
        Me.printFile.Text = "出力"
        Me.printFile.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 162)
        Me.Controls.Add(Me.printFile)
        Me.Controls.Add(Me.excelInputText)
        Me.Controls.Add(Me.excelInput)
        Me.Name = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents excelInput As System.Windows.Forms.Button
    Friend WithEvents excelInputText As System.Windows.Forms.TextBox
    Friend WithEvents printFile As System.Windows.Forms.Button

End Class
