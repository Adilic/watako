
Imports Microsoft.Office.Interop.Excel

Public Class Form1

    Private Sub excelInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles excelInput.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel Files|*.xlsx;*.xls"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            excelInputText.Text = openFileDialog.FileName
        End If
    End Sub



    Private Sub printFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles printFile.Click

        Call ControlB4PaperPrintWithPreview(excelInputText.Text)


    End Sub





    Private Sub ControlB4PaperPrintWithPreview(ByVal filePath As String)
        ' excelアプリをインスタンス化
        Dim excelApp As New Application()

        ' Excelの動きを見えるようにする
        excelApp.Visible = True

        ' excelを開く
        Dim workbook As Workbook = excelApp.Workbooks.Open(filePath)
        Dim worksheet As Worksheet = workbook.Sheets(1)

        ' コピ元の行数列数を取得
        Dim rowCount As Integer = worksheet.UsedRange.Rows.Count
        Dim columnCount As Integer = worksheet.UsedRange.Columns.Count

        ' 印刷用の臨時シートを作成
        Dim tempSheet As Worksheet = workbook.Sheets.Add(After:=workbook.Sheets(workbook.Sheets.Count))
        tempSheet.Name = "TempPrintSheet"

        ' 列の幅
        For i As Integer = 1 To columnCount
            tempSheet.Columns(i).ColumnWidth = worksheet.Columns(i).ColumnWidth
        Next

        ' 行の高
        For j As Integer = 1 To rowCount
            tempSheet.Rows(j).RowHeight = worksheet.Rows(j).RowHeight
        Next

        '横の改ページを削除
        Do While tempSheet.HPageBreaks.Count > 0
            tempSheet.HPageBreaks(1).Delete()
        Loop


        ' 縦の改ページを削除
        Do While tempSheet.VPageBreaks.Count > 0
            tempSheet.VPageBreaks(1).Delete()
        Loop


        '自動的に改ページ付与しないように
        tempSheet.DisplayPageBreaks = False


        Dim currentRow As Integer = 1 ' カレント行
        Dim labelsCount As Integer = 0 '  "お届予定日"の数をカウント
        Dim tempSheetRow As Integer = 1 ' 開始行

        'すべての行をループして、「お届予定日」を探す
        Do While currentRow <= worksheet.UsedRange.Rows.Count
            ' F列が「お届予定日」を含んでいるかを確認する
            If InStr(worksheet.Cells(currentRow, 6).Value, "お届予定日") > 0 Then
                ' 送り状の終了行を見つけて、数量を記録する
                labelsCount += 1

                ' 現在の送り状のデータを一時的なシートにコピーする
                Dim startRow As Integer = currentRow - 28 ' 各送り状が29行であると仮定
                worksheet.Range("A" & startRow & ":CD" & currentRow).Copy(Destination:=tempSheet.Range("A" & tempSheetRow))

                ' 一時シートの行数を更新する
                tempSheetRow += 29 ' 次回は一時シートの次の部分にコピーする

                ' 3枚の送り状がいっぱいになったら、改ページを挿入する
                If labelsCount = 4 Then
                    tempSheet.HPageBreaks.Add(Before:=tempSheet.Rows(tempSheetRow))
                    labelsCount = 0 ' カウントをリセットし、次のページを準備
                End If
            End If

            ' 次の行を処理する
            currentRow += 1
        Loop



        ' 用紙をB4にセットし、印刷範囲を設定します
        With tempSheet.PageSetup
            .PaperSize = XlPaperSize.xlPaperB4
            .Orientation = XlPageOrientation.xlPortrait
            .Zoom = 65                  '65％縮小
            .FitToPagesWide = 1         ' 各ページの幅は B4 用紙 1 枚に収まります
            .FitToPagesTall = False

            .CenterHorizontally = True  ' 水平方向に中央揃え
            .CenterVertically = True    ' 垂直方向に中央揃え
        End With

        ' Excelの動きを見えるようにする
        excelApp.Visible = True


        tempSheet.PrintPreview()

        ' クリップボードをクリア
        excelApp.CutCopyMode = False


        MsgBox("印刷が完了しました")

        ' Excelを閉じる
        workbook.Close(False)
        excelApp.Quit()

        ' オブジェクトをリリース
        System.Runtime.InteropServices.Marshal.ReleaseComObject(tempSheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)

    End Sub






End Class



