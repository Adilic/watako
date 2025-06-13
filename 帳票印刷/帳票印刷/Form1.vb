
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices

Public Class Form1

    Private Sub excelInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles excelInput.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel Files|*.xlsx;*.xls"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            excelInputText.Text = openFileDialog.FileName
        End If
    End Sub



    Private Sub printFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles printFile.Click

        Call ExcelOutput(excelInputText.Text)


    End Sub




    Private Sub ExcelOutput(ByVal inputPath As String)

        Dim templatePath As String = "C:\Users\hcs\Desktop\watako_file\template.xlsx"
        Dim outputPath As String = "C:\Users\hcs\Desktop\watako_file\output.xlsx"

        ' excelアプリをインスタンス化
        Dim excelApp As New Application()

        ' Excelの動きを見えるようにする
        excelApp.Visible = True
        Try
            ' excelを開く
            Dim sourceWb As Workbook = excelApp.Workbooks.Open(inputPath)
            Dim sourceSheet As Worksheet = sourceWb.Sheets(1)

            'templateを開く
            Dim templateWb As Workbook = excelApp.Workbooks.Open(templatePath)
            Dim templateSheet As Worksheet = templateWb.Sheets(1)

            ' outpugExcelを新規に作る
            Dim outputWb As Workbook = excelApp.Workbooks.Add()
            'templateからシートをコピーする
            templateSheet.Copy(Before:=outputWb.Sheets(1))
            Dim outputSheet As Worksheet = outputWb.Sheets(1)
            outputSheet.Name = "送り状"

            ' 删除初始空白 Sheet(2,3)
            For i = outputWb.Sheets.Count To 2 Step -1
                outputWb.Sheets(i).Delete()
            Next

            Call CopyAllSlipsFromSourceToOutput(sourceSheet, outputSheet, templateSheet, excelApp)

            ' 设置打印区域
            Dim lastRow As Integer = outputSheet.UsedRange.Rows.Count
            outputSheet.PageSetup.PrintArea = "$A$1:$CD$" & lastRow.ToString()

            'outpugExcelを保存する
            outputWb.SaveAs(outputPath)

            MsgBox("出力完了：" & outputPath)


            outputWb.Close()
            templateWb.Close(False)

            ' Excelを閉じる
            sourceWb.Close(False)
            excelApp.Quit()
        Catch ex As Exception
            MsgBox("エラー：" & ex.Message)
        Finally
            ' オブジェクトをリリース
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
        End Try

    End Sub



    ''' <summary>
    ''' 从源工作表的指定行范围中找到条形码图形，并复制粘贴到目标工作表指定单元格
    ''' （兼容 Excel 2010 Interop 12.0）
    ''' </summary>
    ''' <param name="srcSheet">来源工作表对象</param>
    ''' <param name="startRow">开始行号（如1）</param>
    ''' <param name="endRow">结束行号（如29）</param>
    ''' <param name="destSheet">目标工作表对象</param>
    ''' <param name="destCellAddress">目标单元格（如"C3"）</param>
    Public Sub CopyBarcodeShapeToCell(ByVal srcSheet As Object, _
                                       ByVal startRow As Integer, _
                                       ByVal endRow As Integer, _
                                       ByVal destSheet As Object, _
                                       ByVal destCellAddress As String)


        Dim shp As Object
        For Each shp In srcSheet.Shapes
            ' 条件1：图形名称为 Picture # 或 Picture1 形式
            If shp.Name Like "Picture #" Or shp.Name Like "Picture*" Then
                ' 条件2：图形Top位置在当前送り状行范围内
                Dim row As Integer
                row = shp.TopLeftCell.Row
                If row >= startRow And row <= endRow Then
                    ' 条件3：图形大小过滤（避免空图）
                    If shp.Width > 50 And shp.Height > 10 Then
                        shp.Copy()
                        destSheet.Paste()


                        ' 关键：Interop 12.0 中必须使用 .Item 来获取图形
                        Dim pastedShp As Object
                        pastedShp = destSheet.Shapes.Item(destSheet.Shapes.Count)


                        pastedShp.Top = destSheet.Range(destCellAddress).Top
                        pastedShp.Left = destSheet.Range(destCellAddress).Left

                    End If
                End If
            End If
        Next


    End Sub

    Private Sub CopyOneSlipTextThreeTimes(ByVal sourceSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal srcStartRow As Integer, ByVal destStartRow As Integer)
        Dim oneSlipRows As Integer = 29
        Dim colCount As Integer = 82 ' CD列是第82列


        ' === 从源表读取 29 行 × 82 列的数据 ===
        Dim sourceRange As Range = sourceSheet.Range(sourceSheet.Cells(srcStartRow, 1), sourceSheet.Cells(srcStartRow + oneSlipRows - 1, colCount))
        Dim dataArr As Object(,) = sourceRange.Value2


        ' === 在目标表贴上三次 ===
        Dim i As Integer
        For i = 0 To 2
            Dim targetRow As Integer = destStartRow + i * oneSlipRows
            Dim destRange As Range = outputSheet.Range(outputSheet.Cells(targetRow, 1), outputSheet.Cells(targetRow + oneSlipRows - 1, colCount))
            destRange.Value2 = dataArr
        Next

    End Sub

    Private Function IsSlipBlockEmpty(ByVal sheet As Worksheet, ByVal startRow As Integer, ByVal colCount As Integer, ByVal rowCount As Integer) As Boolean
        For r As Integer = 0 To rowCount - 1
            For c As Integer = 1 To colCount
                Dim cellValue As Object = sheet.Cells(startRow + r, c).Value
                If cellValue IsNot Nothing AndAlso cellValue.ToString().Trim() <> "" Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    Private Sub CopyAllSlipsFromSourceToOutput(ByVal sourceSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal templateSheet As Worksheet, ByVal excelApp As Application)
        Dim oneSlipRows As Integer = 29 '一つの送り状は29行。
        Dim pageRows As Integer = 88
        Dim colCount As Integer = 82


        Dim currentSrcRow As Integer = 1
        Dim currentDestRow As Integer = 1

        Dim rowHeights() As Double = GetTemplateRowHeights(templateSheet, pageRows)
        Do While Not IsSlipBlockEmpty(sourceSheet, currentSrcRow, colCount, oneSlipRows)

            'templateから書式をコピー
            Call CopyTemplateFormatToDestination(templateSheet, outputSheet, currentDestRow, pageRows, colCount, excelApp)
            Call ApplyRowHeights(outputSheet, currentDestRow, rowHeights)

            '各送り状に対して3件コピー
            Call CopyOneSlipTextThreeTimes(sourceSheet, outputSheet, currentSrcRow, currentDestRow)

            'barcodeをコピー
            For i As Integer = 0 To 2
                Dim destRow As Integer = currentDestRow + i * oneSlipRows
                Dim destCellAddress As String = "AE" & destRow.ToString()

                Call CopyBarcodeShapeToCell(
                    sourceSheet,
                    currentSrcRow,
                    currentSrcRow + oneSlipRows - 1,
                    outputSheet,
                    destCellAddress
                )
            Next


            '改ページを挿入
            outputSheet.HPageBreaks.Add(Before:=outputSheet.Rows(currentDestRow + pageRows))

            currentSrcRow += oneSlipRows
            currentDestRow += pageRows
        Loop
    End Sub



    Private Sub CopyTemplateFormatToDestination(ByVal templateSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal destStartRow As Integer, ByVal pageRows As Integer, ByVal colCount As Integer, ByVal excelApp As Application)
        Dim templateRange As Range = templateSheet.Range(templateSheet.Cells(1, 1), templateSheet.Cells(pageRows, colCount))
        Dim destRange As Range = outputSheet.Range(outputSheet.Cells(destStartRow, 1), outputSheet.Cells(destStartRow + pageRows - 1, colCount))
        templateRange.Copy()
        destRange.PasteSpecial(XlPasteType.xlPasteFormats)
        excelApp.CutCopyMode = False

    End Sub
    Private Function GetTemplateRowHeights(ByVal templateSheet As Worksheet, ByVal pageRows As Integer) As Double()
        Dim heights(pageRows - 1) As Double
        For i As Integer = 1 To pageRows
            heights(i - 1) = templateSheet.Rows(i).RowHeight
        Next
        Return heights
    End Function

    Private Sub ApplyRowHeights(ByVal outputSheet As Worksheet, ByVal destStartRow As Integer, ByVal heights() As Double)
        For i As Integer = 0 To heights.Length - 1
            outputSheet.Rows(destStartRow + i).RowHeight = heights(i)
        Next
    End Sub



End Class



