let
    ソース = Csv.Document(File.Contents("C:\Users\user\OneDrive\プロジェクト\Excel\検査\data\1867_全検査結果一覧 - コピー.csv"),[Delimiter=",", Columns=12, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    変更された型 = Table.TransformColumnTypes(ソース,{{"Column1", type text}, {"Column2", type text}, {"Column3", type text}, {"Column4", type text}, {"Column5", type text}, {"Column6", type text}, {"Column7", type text}, {"Column8", type text}, {"Column9", type text}, {"Column10", type text}, {"Column11", type text}, {"Column12", type text}}),
    昇格されたヘッダー数 = Table.PromoteHeaders(変更された型, [PromoteAllScalars=true]),
    変更された型1 = Table.TransformColumnTypes(昇格されたヘッダー数,{{"検査カテゴリ", type text}}),
    追加されたインデックス = Table.AddIndexColumn(変更された型1, "インデックス", 1, 1, Int64.Type),
    結合列を追加 = Table.DuplicateColumn(
    追加されたインデックス,
    "行番号",              // 複製元の列名
    "結合"        // 新しく作成する列名
    ),
    結合列の置換 = 
    Table.ReplaceValue(結合列を追加,"〃〃","結合",Replacer.ReplaceValue,{"結合"}),
    // すべての列で「〃〃」を null に置き換え
    null置換 = Table.ReplaceValue(結合列の置換, "〃〃", null, Replacer.ReplaceValue, Table.ColumnNames(結合列の置換)),
    // null を上の値で埋める
    埋められた値 = Table.FillDown(null置換, Table.ColumnNames(null置換)),
    キャリッジリターン改行コードの削除 = Table.ReplaceValue(
    埋められた値,                    // 変換対象のテーブル (前のステップ名)
    Character.FromNumber(13),  // 置換対象文字 (CHAR(13) = \r)
    "",                        // 置換後の文字列（ここでは空文字に置換）
    Replacer.ReplaceText,      // テキスト置換モード
    {"コメント","対象ソースコード","修正ソースコード"}             // 置換を適用する列名
),
    並べ替えられた列 = Table.ReorderColumns(キャリッジリターン改行コードの削除,{"管理番号", "URL", "検査カテゴリ", "検査項目", "達成基準/達成方法", "行番号", "検査結果", "コメント", "対象ソースコード", "修正ソースコード", "登録者", "更新者", "インデックス"})
in
    並べ替えられた列