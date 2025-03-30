Sub ListFilesInFolder()
  Dim folderPath As String
  Dim fileName As String
  Dim ws As Worksheet
  Dim filenamesRange As Range
  Dim rowNum As Integer

  ' 「設定」シート
  Set ws = ThisWorkbook.Sheets("設定")
  'ファイルの内容をクリア
  Set filenamesRange = ws.Range("B2:B1000")
  filenamesRange.ClearContents

  ' Excelファイルのあるフォルダを基準に相対パスを指定
  folderPath = "./比較対象"

    ' ファイル取得開始
    fileName = Dir(folderPath & "*.*") ' すべてのファイルを対象

    rowNum = 2 ' データ開始行

    ' ループでファイルを取得
    Do While fileName <> ""
      ws.Cells(rowNum, 1).Value = fileName
      rowNum = rowNum + 1
      fileName = Dir ' 次のファイルを取得
    Loop

    MsgBox "フォルダ内のファイルをリストアップしました。", vbInformation
End Sub
