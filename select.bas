'フォルダにあるCSVファイル名を扱うクラス
Option Explicit

' メンバ変数
Private ws As Worksheet
Private dev_ws As Worksheet
'比較対象のファイルを選択するセル
Private m_compare1_range As range
Private m_compare2_range As range
'フォルダ内のCSVファイル名を格納するセル範囲
Private m_file_names_range As range
'フォルダへのパスを格納するセル
Private m_folder_path As String

' コンストラクタ
Public Sub Class_Initialize()
  '比較シート
  Set ws = ThisWorkbook.Sheets("比較")
  Set m_compare1_range = ws.Range("G1")
  Set m_compare2_range = ws.Range("K1")

  '開発シート
  Set dev_ws = ThisWorkbook.Sheets("開発用")
  'フォルダ内のCSVファイル名を格納するセル範囲
  Set m_file_names_range = dev_ws.Range("D2:D1000")
  'CSV格納フォルダのパス
  m_folder_path = dev_ws.Range("E2").value
End Sub

'm_folder_pathのファイルを取得して書き込む
Sub files_to_range()
  Dim fileName As String
  Dim rowNum As Integer

  m_file_names_range.ClearContents ' セル範囲をクリア

  ' フォルダパスが「\」で終わっていない場合は追加
  If
    Right(m_folder_path, 1) <> "\" Then m_folder_path = m_folder_path & "\"
  End If

  fileName = Dir(m_folder_path & "*.*") ' すべてのファイルを対象
  '存在しない場合は終了
  If fileName = "" Then
    MsgBox "指定されたフォルダにファイルが存在しません。", vbInformation
   Exit Sub
  End If
  rowNum = 1 ' データ開始行

  ' ループでファイルを取得
  Do While fileName <> ""
    Debug.Print (fileName)
    m_file_names_range.Cells(rowNum, 1).value = fileName
    rowNum = rowNum + 1
    fileName = Dir ' 次のファイルを取得
  Loop
End Sub

