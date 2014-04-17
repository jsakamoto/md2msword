# md2msword (Markdown to MS-Word)

## これは何?

Markdown 書式で書かれた、拡張子が .md のファイルを、MS-Word 文書に差し込みする、PowerShell スクリプトです。

## 動作

PowerShell コンソールで

    > .\md2msword.ps1

などと実行すると、

1. このスクリプトファイルがあるフォルダにある、template.docx を MS-Word を起動して開き、
2. 次に、このスクリプトファイルがあるフォルダにある、拡張子が .md のファイルを列挙し、
3. それら .md ファイルを Markdown 書式のテキストファイルであるものとしてファイル名順に読み込み、 HTML に変換しつつ、
4. 先に開いた Word文書(template.docx) の本文に差し込んでいきます。
5. すべての .md ファイルを処理し終わったら、Word文書内のフィールドを更新します。目次フィールドなどがあれば、目次が完成します。

## 動作確認済環境

- Windows8.1(x64)
- PowerShell ver.3.0
- Microsoft Word 2013
