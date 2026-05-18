# LaTeX Equation Editor for PowerPoint

ChatGptによる開発プロジェクトの一つです．
PowerPoint の Office.js アドインのプロトタイプです。
LaTeX 記法の数式を MathJax で SVG に変換し、スライドに挿入します。
挿入した図形には `Shape.tags` と alt text に元の LaTeX を保存するため、あとから数式をクリックしてタスクペインで再編集できる仕組みになっています。

## 機能

- LaTeX → MathJax SVG 変換
- PowerPoint スライドへの SVG 挿入
- 選択した数式から LaTeX ソースを読み込み
- 選択した数式の更新
- 色、display style、基準高さの指定
- PPTX 保存後も LaTeX ソースを保持

## 導入方法〜数式挿入まで

### macの場合

1. ターミナルで以下を実行します
```bash
mkdir -p "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

cp manifest.xml "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/manifest.xml"
```
2. PowerPoint を開きます。
3. 「ホーム」タブの「アドイン」から `LaTeX数式` ボタンを選択します。
4. LaTeX を入力して `新しく挿入` を押します。

### windowsの場合

現在開発中

## 編集の流れ

1. スライド上の挿入済み数式をクリックします。
2. アドインが開いていれば、 LaTeX が自動読み込みされます。
3. 自動で読み込まれない場合は `選択数式を読み込み` を押します。
4. LaTeX を修正し、`選択数式を更新` を押します。

## 注意点

- PowerPoint のネイティブ数式オブジェクトではなく、SVG 画像として挿入します。
- そのため、PowerPoint 標準の数式エディタで直接編集する方式ではありません。
- 編集可能性は、このアドインが Shape tags / alt text に保存した LaTeX を読み戻すことで実現しています。
- MathJax と Office.js は CDN から読み込んでいます。オフライン運用したい場合は、ライブラリをローカル配信に変更してください。


## 免責事項
当アドインの情報・内容の利用により生じたトラブルや損害について、運営者は一切の責任を負いません。
