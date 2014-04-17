# このスクリプトと同じフォルダにある Markdown ファイル (*.md) を MS-Word 文書に差し込みます。

# カレントフォルダを、このスクリプトファイルがあるフォルダに移動する
pushd (split-path $PSCommandPath)

# Markdown > HTML 変換に、.NET用のライブラリ "MarkdownDeep" を読み込み、インスタンス生成。
Add-Type -Path '.\MarkdownDeep.NET.1.5\lib\.NetFramework 3.5\MarkdownDeep.dll'
$mdprocc = New-Object MarkdownDeep.Markdown
$mdprocc.ExtraMode = $true # 表のマークアップに対応するために Markdown Extra モードを有効化

# MS-Word を COM 経由でインスタンス化、テンプレートのWord文書ファイルを読み込み
# (このテンプレートに、Markdownから変換した HTML を差し込みしていく)
$msword = New-Object -ComObject "Word.Application"
$msword.Visible = $true
$doc = $msword.Documents.Open(((ls ".\template.docx").FullName))
$doc.Activate()

# HTML を差し込む位置にカレットを移動
# ポイント: いったん文末まで移動するが、その1文字(1行)手前にカレットを戻すのがポイント。
# 文末にカレットを置いて、そこにHTMLを差し込むと、フッターが、差し込んだHTMLにあるものとされて、消されてしまう。
# (このテンプレートでは、フッターにページ番号フィールドを置いてあり、その折角のフッターを消されると嫌なので)
$msword.Selection.EndKey(6) > $null
$msword.Selection.Move(1,-1) > $null

# カレントフォルダにある、拡張子=.md のファイルを列挙し、MarkdownDeep を使って HTML に変換しつつ、Word文書に差し込み。
ls .\*.md | 
sort Name |
% {
    # .md を文字列として読み込み ～ HTML文字列に変換 ～ 一時ファイルへ書き出し
    $mdcontent = cat $_ -Encoding UTF8 -Raw
    $html = ("<html><body>" + $mdprocc.Transform($mdcontent) + "</body></html>")
    Set-Content -Path .\~temp.html -Encoding UTF8 $html

    # Word文書へ、変換した(一時ファイルの) HTML を差し込み
    $msword.Selection.InsertFile(((ls ".\~temp.html").FullName), "", $false, $false, $false)
}

# 一時ファイルを清掃
del .\~temp.html

# 最後に、Word文書全体を選択してフィールド更新することで、目次を完成させる。
$msword.Selection.WholeStory()
$msword.Selection.Fields.Update() > $null

# このスクリプトではここまで。
# お好みで、別名で保存したり、MS-Word標準の機能で PDF へエクスポートしたりできます。
