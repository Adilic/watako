
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
        Dim sourceWb As Workbook = Nothing
        Dim templateWb As Workbook = Nothing
        Dim outputWb As Workbook = Nothing

        Try
            ' excelを開く
            sourceWb = excelApp.Workbooks.Open(inputPath)
            Dim sourceSheet As Worksheet = sourceWb.Sheets(1)

            'templateを開く
            templateWb = excelApp.Workbooks.Open(templatePath)
            Dim templateSheet As Worksheet = templateWb.Sheets(1)

            ' outpugExcelを新規に作る
            outputWb = excelApp.Workbooks.Add()
            'templateからシートをコピーする
            templateSheet.Copy(Before:=outputWb.Sheets(1))
            Dim outputSheet As Worksheet = outputWb.Sheets(1)
            outputSheet.Name = "送り状"

            ' 删除初始空白 Sheet(2,3)
            For i = outputWb.Sheets.Count To 2 Step -1
                outputWb.Sheets(i).Delete()
            Next

            ' 设置打印区域
            Dim lastRow As Integer = CopyAllSlipsFromSourceToOutput(sourceSheet, outputSheet, templateSheet, excelApp)
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
            System.Runtime.InteropServices.Marshal.ReleaseComObject(outputWb)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(templateWb)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWb)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
        End Try

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

    Private Function CopyAllSlipsFromSourceToOutput(ByVal sourceSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal templateSheet As Worksheet, ByVal excelApp As Application) As Integer
        Dim oneSlipRows As Integer = 29 '一つの送り状は29行。
        Dim pageRows As Integer = 88
        Dim colCount As Integer = 82


        Dim currentSrcRow As Integer = 1
        Dim currentDestRow As Integer = 1

        Dim barcodeMap As Dictionary(Of Integer, Shape) = GetBarcodeShapeMap(sourceSheet)

        Do While Not IsSlipBlockEmpty(sourceSheet, currentSrcRow, colCount, oneSlipRows)

            'templateから書式をコピー
            Call CopyTemplateFormatToDestination(templateSheet, outputSheet, currentDestRow, pageRows, colCount, excelApp)

            '各送り状に対して3件コピー
            Call CopyOneSlipTextThreeTimes(sourceSheet, outputSheet, currentSrcRow, currentDestRow)

            'barcodeをコピー
            For i As Integer = 0 To 2
                Dim destRow As Integer = currentDestRow + i * oneSlipRows
                Dim destCellAddress As String = "AE" & destRow.ToString()

                Call CopyBarcodeShapeToCell(barcodeMap, currentSrcRow, outputSheet, destCellAddress, excelApp)

            Next


            '改ページを挿入
            outputSheet.HPageBreaks.Add(Before:=outputSheet.Rows(currentDestRow + pageRows))

            currentSrcRow += oneSlipRows
            currentDestRow += pageRows
        Loop
        Return currentDestRow - 1
    End Function


    Private Sub CopyTemplateFormatToDestination(ByVal templateSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal destStartRow As Integer, ByVal pageRows As Integer, ByVal colCount As Integer, ByVal excelApp As Application)
        ' 直接复制模板的整行，包括行高和所有格式
        Dim fromRow As Integer = 1
        Dim toRow As Integer = destStartRow

        ' 拼接行字符串，如 "1:88" → "101:188"
        Dim fromRowStr As String = fromRow & ":" & (fromRow + pageRows - 1)
        Dim toRowStr As String = toRow & ":" & (toRow + pageRows - 1)

        templateSheet.Rows(fromRowStr).Copy()
        outputSheet.Rows(toRowStr).PasteSpecial(XlPasteType.xlPasteAll)

        ' 清除剪贴板虚线框
        excelApp.CutCopyMode = False
    End Sub

    Private Function GetBarcodeShapeMap(ByVal sheet As Worksheet) As Dictionary(Of Integer, Shape)
        Dim shapeMap As New Dictionary(Of Integer, Shape)

        For Each shp As Shape In sheet.Shapes
            ' 筛选出可能是条形码的图片
            If shp.Name Like "Picture*" AndAlso shp.Width > 50 AndAlso shp.Height > 10 Then
                Dim row As Integer = shp.TopLeftCell.Row

                ' 一个行号只保存一张图，避免重复
                If Not shapeMap.ContainsKey(row) Then
                    shapeMap(row) = shp
                End If
            End If
        Next

        Return shapeMap
    End Function

    Private Sub CopyBarcodeShapeToCell(ByVal shapeMap As Dictionary(Of Integer, Shape), ByVal sourceRow As Integer, ByVal destSheet As Worksheet, ByVal destCellAddress As String, ByVal excelApp As Application)

        ' 如果该送状没有对应图形，跳过
        If Not shapeMap.ContainsKey(sourceRow) Then Exit Sub

        ' 取出图形
        Dim srcShape As Shape = shapeMap(sourceRow)

        srcShape.Copy()
        destSheet.Paste()
        Dim pastedShape As Shape = destSheet.Shapes.Item(destSheet.Shapes.Count)


        pastedShape.Top = destSheet.Range(destCellAddress).Top
        pastedShape.Left = destSheet.Range(destCellAddress).Left

        excelApp.CutCopyMode = False
    End Sub



End Class









