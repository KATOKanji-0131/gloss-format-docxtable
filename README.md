# gloss-format-docxtable

テキストを整形して、.docxの表の形でインターリニアーグロスにします。

# Features
* プレーンテキスト形式で入力された例文から.docxの表によるインターリニアーグロスを生成します。
* 幅がいっぱいになったら自動で折り返します。
* グロス行の略号は自動でスモールキャピタルになります。

# Requirement
* Python 3
* pandas
* python-docx

# Installation
```bash
pip3 install pandas
pip3 install python-docx
```

# Usage
### 例文テキストの準備

```example.txt
\gla 私 は 先生 だ
\glb I TOP teacher COP
\glft I am a teacher.

\gla 明日 は 晴れ だ
\glb tomorrow TOP sunny COP
\glft It will be sunny tomorrow.
```
`example.txt`に例文を書き込んでください。

* 形態素の行を`\gla `から、グロスの行を`\glb `から、訳の行を`\glft `から開始してください。
* 形態素とグロスは、半角スペースまたはタブで区切ってください。

### 行指定子の変更
どの行を形態素、グロス、訳として読み込むかの指定は`\gla ` `\gla ` `\gla `で行っています。これを例えばそれぞれ`\morpheme`、`\gloss`、`\translation`へ変更したい場合、`variables.py`の
```variables.py
morph_spcf = r"\\gla" # 形態素行の指定子
gl_spcf = r"\\glb" # グロス行の指定子
trsl_spcf = r"\\glft" # 訳行の指定子
```
という部分を
```variables.py
morph_spcf = r"\\morpheme" # 形態素行の指定子
gl_spcf = r"\\gloss" # グロス行の指定子
trsl_spcf = r"\\translation" # 訳行の指定子
```
としてください。
# Note

* 同一の列にあるセルが全て同じ幅で出力されてしまいます。今後のアップデートで、セルの幅がもっと縮まるように修正します。
* 訳が一つのセルに押し込められてしまっています。セルを結合して横に長く表示したいのですが、結合すると表示に影響が出てしまうのでこのままにしています。今後のアップデートで、訳を横に長く表示できるように修正します。
* セルの余白が広すぎますが、今後のアップデートでもっと詰めて表示できるように修正します。
* 現在は例文番号がアラビア数字の連番だけですが、数字の後ろにアルファベットをつけて4a, 4b...のように出力できるようにアップデートします。

# Author

* 加藤幹治 KATO, Kanji
* 東京外国語大学大学院/日本学術振興会特別研究員 TUFS/JSPS
* jiateng.ganzhi[at]gmail.com

# License
"gloss-format-docxtable" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).