
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

        Dim templatePath As String = "C:\Users\HCS\Desktop\watako_file\template.xlsx"
        Dim outputPath As String = "C:\Users\HCS\Desktop\watako_file\output.xlsx"

        'excelアプリをインスタンス化
        Dim excelApp As New Application()

        'Excelの動きを見えるようにする
        excelApp.Visible = True
        Dim sourceWb As Workbook = Nothing
        Dim templateWb As Workbook = Nothing
        Dim outputWb As Workbook = Nothing

        Try
            'excelを開く
            sourceWb = excelApp.Workbooks.Open(inputPath)
            Dim sourceSheet As Worksheet = sourceWb.Sheets(1)

            'templateを開く
            templateWb = excelApp.Workbooks.Open(templatePath)
            Dim templateSheet As Worksheet = templateWb.Sheets(1)

            'outpugExcelを新規に作る
            outputWb = excelApp.Workbooks.Add()

            'templateからシートをコピーする
            templateSheet.Copy(Before:=outputWb.Sheets(1))
            Dim outputSheet As Worksheet = outputWb.Sheets(1)
            outputSheet.Name = "送り状"

            'Sheet2,3を削除する
            For i = outputWb.Sheets.Count To 2 Step -1
                outputWb.Sheets(i).Delete()
            Next

            '送り状をコピーペーして、プリントエリアを設定する
            Dim lastRow As Integer = CopyAllSlipsFromSourceToOutput(sourceSheet, outputSheet, templateWb, excelApp)
            outputSheet.PageSetup.PrintArea = "$A$1:$AZ$" & lastRow.ToString()

            'outpugExcelを保存する
            outputWb.SaveAs(outputPath)

            MsgBox("出力完了：" & outputPath)


            outputWb.Close()
            templateWb.Close(False)

            'Excelを閉じる
            sourceWb.Close(False)
            excelApp.Quit()
        Catch ex As Exception
            MsgBox("エラー：" & ex.Message)
        Finally
            'オブジェクトをリリース
            System.Runtime.InteropServices.Marshal.ReleaseComObject(outputWb)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(templateWb)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWb)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
        End Try

    End Sub




    ''' <summary>
    ''' (主调度函数) 遍历源文件中的每一个送り状，并为每一个都生成一个三联打印页。
    ''' </summary>
    Private Function CopyAllSlipsFromSourceToOutput(ByVal sourceSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal templateWb As Workbook, ByVal excelApp As Application) As Integer


        Dim pageRows As Integer = 87    ' 整个页面（包含三个送り状和间隙）的总行数

        Dim sourceSlipRows As Integer = 29
        ' --- 2. 在这里定义条形码的定位信息 ---
        Dim sourceSlipCols As Integer = 52
        Dim barcodeAnchorAddress As String = "V1"

        ' --- 3. 初始化循环变量 ---
        Dim currentSrcRow As Integer = 1
        Dim currentDestRow As Integer = 1
        Dim barcodeMap As Dictionary(Of Integer, Shape) = GetBarcodeShapeMap(sourceSheet)
        Dim templateSheet As Worksheet = templateWb.Sheets(1)


        ' --- 4. 按“源送り状”为单位进行大循环 ---
        Do While Not IsSlipBlockEmpty_CountA(sourceSheet, currentSrcRow, sourceSlipCols, sourceSlipRows, excelApp)

            ' 5. 为新的一页铺设背景格式
            Call CopyTemplateFormatToDestination(templateSheet, outputSheet, currentDestRow, pageRows, sourceSlipCols, excelApp)


            ' --- 6. 在一页内循环三次，每次都从源文件精确填充 ---
            For i As Integer = 0 To 2
                ' --- 开始处理第 i 个送り状副本 ---

                ' 7. 偏移値
                Dim destRowForThisSlip As Integer = currentDestRow + (i * 30) ' 27行内容 + 3行间隔 = 30行步进


                ' --- 任务一：从源文件填充所有文字 ---
                '    注意：srcBlockStartRow 参数始终是 currentSrcRow，确保数据源唯一
                Call FillAllTextForOneSlip(sourceSheet, outputSheet, templateWb, currentSrcRow, destRowForThisSlip)


            Next

            ' 8. 更新光标，准备处理下一个源送り状和下一页
            outputSheet.HPageBreaks.Add(Before:=outputSheet.Rows(currentDestRow + pageRows))
            currentSrcRow += sourceSlipRows
            currentDestRow += pageRows
        Loop

        Return currentDestRow - 1
    End Function


    Private Sub CopyTemplateFormatToDestination(ByVal templateSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal destStartRow As Integer, ByVal pageRows As Integer, ByVal colCount As Integer, ByVal excelApp As Application)

        Dim fromRow As Integer = 1
        Dim toRow As Integer = destStartRow

        '例えば： 行"1:87"の内容 →行 "101:187"二ピーペーする
        Dim fromRowStr As String = fromRow & ":" & (fromRow + pageRows - 1)
        Dim toRowStr As String = toRow & ":" & (toRow + pageRows - 1)
        'templateの行をコピーする
        templateSheet.Rows(fromRowStr).Copy()
        outputSheet.Rows(toRowStr).PasteSpecial(XlPasteType.xlPasteAll)

        'クリップボードを消去する
        excelApp.CutCopyMode = False
    End Sub


    ''' <summary>
    ''' (核心工作函数) 根据映射表，为单个送り状填充所有“文字”信息。
    ''' </summary>
    Private Sub FillAllTextForOneSlip(ByVal sourceSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal templateWb As Workbook, ByVal srcBlockStartRow As Integer, ByVal destBlockStartRow As Integer)
        Dim mapSheet As Worksheet = templateWb.Sheets("マップ") ' 确保您的map表名是 "マップ"
        Dim lastMapRow As Integer = mapSheet.Cells(mapSheet.Rows.Count, "C").End(XlDirection.xlUp).Row
        If lastMapRow < 2 Then Exit Sub
        Dim mappingArray As Object(,) = mapSheet.Range("A2:D" & lastMapRow).Value2

        For i As Integer = 1 To UBound(mappingArray, 1)
            Dim itemType As String = If(mappingArray(i, 1) IsNot Nothing, mappingArray(i, 1).ToString(), "").Trim()
            Dim itemDesc As String = If(mappingArray(i, 2) IsNot Nothing, mappingArray(i, 2).ToString(), "")
            Dim sourceStartAddr As String = mappingArray(i, 3).ToString()
            Dim destStartAddr As String = mappingArray(i, 4).ToString()


            Select Case itemType
                Case "Fixed", ""
                    Dim sourceCell As Range = sourceSheet.Range(sourceStartAddr).Offset(srcBlockStartRow - 1, 0)
                    Dim destCell As Range = outputSheet.Range(destStartAddr).Offset(destBlockStartRow - 1, 0)
                    destCell.NumberFormatLocal = "@"
                    destCell.Value = sourceCell.Text
                Case "Detail"
                    If itemDesc = "伝票番号" Or itemDesc = "先方No" Then
                        Call CopyDetailList_Simple(sourceSheet, outputSheet, srcBlockStartRow, destBlockStartRow, sourceStartAddr, destStartAddr)
                    ElseIf itemDesc = "納品明細" Then
                        Call CopyDetailList_Complex(sourceSheet, outputSheet, srcBlockStartRow, destBlockStartRow, sourceStartAddr, destStartAddr)
                    End If
            End Select
        Next
    End Sub

    ''' <summary>
    ''' (辅助函数) 复制简单的单列列表。
    ''' </summary>
    Private Sub CopyDetailList_Simple(ByVal sourceSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal srcBlockStartRow As Integer, ByVal destBlockStartRow As Integer, ByVal sourceStartAddr As String, ByVal destStartAddr As String)
        Dim srcCell As Range = sourceSheet.Range(sourceStartAddr).Offset(srcBlockStartRow - 1, 0)
        Dim destCell As Range = outputSheet.Range(destStartAddr).Offset(destBlockStartRow - 1, 0)
        Dim rowOffset As Integer = 0
        Do While Not String.IsNullOrEmpty(srcCell.Offset(rowOffset, 0).Text)
            Dim currentDestCell As Range = destCell.Offset(rowOffset, 0)
            ' 【终极格式修正】
            currentDestCell.NumberFormatLocal = "@"
            currentDestCell.Value = srcCell.Offset(rowOffset, 0).Text

            rowOffset += 1
        Loop
    End Sub


    ''' <summary>
    ''' (辅助函数) 复制复杂的、隔行不同列的列表。
    ''' </summary>
    Private Sub CopyDetailList_Complex(ByVal sourceSheet As Worksheet, ByVal outputSheet As Worksheet, ByVal srcBlockStartRow As Integer, ByVal destBlockStartRow As Integer, ByVal sourceStartAddr As String, ByVal destStartAddr As String)
        Dim srcStartCell As Range = sourceSheet.Range(sourceStartAddr).Offset(srcBlockStartRow - 1, 0)
        Dim destStartCell As Range = outputSheet.Range(destStartAddr).Offset(destBlockStartRow - 1, 0)
        Dim itemIndex As Integer = 0
        Do While Not String.IsNullOrEmpty(srcStartCell.Offset(itemIndex * 2, 0).Text)
            ' --- TODO: 这是您需要根据源文件和目标模板定义的“复制模式” ---
            Dim srcNameCell = srcStartCell.Offset(itemIndex * 2, 0)
            Dim destNameCell = destStartCell.Offset(itemIndex * 2, 0)
            destNameCell.NumberFormatLocal = "@"
            destNameCell.Value = srcNameCell.Text

            Dim srcQtyCell = srcNameCell.Offset(1, 3) ' 源：品名下一行，右一列
            Dim destQtyCell = destNameCell.Offset(1, 3)
            destQtyCell.NumberFormatLocal = "@"
            destQtyCell.Value = srcQtyCell.Text

            itemIndex += 1
        Loop
    End Sub

    Private Function GetBarcodeShapeMap(ByVal sheet As Worksheet) As Dictionary(Of Integer, Shape)
        'バーコードを格納するディクショナリを宣言
        Dim shapeMap As New Dictionary(Of Integer, Shape)
        'インプットにあるすべてのshqesに対しえバーコードを探す
        For Each shp As Shape In sheet.Shapes
            'バーコードを選別する
            If shp.Name Like "Picture*" AndAlso shp.Width > 50 AndAlso shp.Height > 10 Then
                'バーコード所在の行番号を記録
                Dim row As Integer = shp.TopLeftCell.Row
                '一行番号一バーコード、重複不可
                If Not shapeMap.ContainsKey(row) Then
                    'ディクショナリに行番号とその行にあるバーコードを格納する
                    shapeMap(row) = shp
                End If
            End If
        Next

        Return shapeMap
    End Function

    Private Sub CopyBarcodeShapeToCell(ByVal shapeMap As Dictionary(Of Integer, Shape), ByVal sourceRow As Integer, ByVal destSheet As Worksheet, ByVal destCellAddress As String, ByVal excelApp As Application)

        '送り状にバーコードがないの場合、処理を飛ばす
        If Not shapeMap.ContainsKey(sourceRow) Then Exit Sub

        '該当行番号のバーコードを取る
        Dim srcShape As Shape = shapeMap(sourceRow)
        'コピーぺー
        srcShape.Copy()
        destSheet.Paste()
        'コピーペーされたバーコードを選択して
        Dim pastedShape As Shape = destSheet.Shapes.Item(destSheet.Shapes.Count)

        '貼付先に移動する
        pastedShape.Top = destSheet.Range(destCellAddress).Top
        pastedShape.Left = destSheet.Range(destCellAddress).Left
        'クリップボードを消去する
        excelApp.CutCopyMode = False
    End Sub

    Private Function IsSlipBlockEmpty_CountA(ByVal sheet As Worksheet, ByVal startRow As Integer, ByVal colCount As Integer, ByVal rowCount As Integer, ByVal excelApp As Application) As Boolean

        '次に処理すべきのrangeを定義する
        Dim targetRange As Range = sheet.Range(sheet.Cells(startRow, 1), sheet.Cells(startRow + rowCount - 1, colCount))

        ' CountAで次に処理すべきのrangeにデータがあるかを判断する
        Dim nonEmtpyCellCount As Double = excelApp.WorksheetFunction.CountA(targetRange)
        'データがないの場合、falseを返す
        Return nonEmtpyCellCount = 0

    End Function


End Class
