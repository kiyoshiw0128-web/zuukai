# zukai-slides

図解・グラフ・フロー・2x2 などのビジネスの図を使い、**タイトルと絵を見るだけで言いたいことが推測できる** 日本語の PowerPoint を作る、Claude Code（Agent Skills）用のスキルです。
全スライドのノートに、話し手がそのまま読める **ナレーション（話し言葉の台本）** を必ず入れます。

- 1枚1メッセージ。見出しは言い切りの1文
- 割合はマス、比較は棒、手順はフロー、判断ははい／いいえの図。部品をそろえてあります
- 研究や調査の結果には「根拠」、自分たちの考えには「提案」の札を付けて分けます
- ナレーションのないスライド、18pt 未満の文字、はみ出しは、保存と検査で止まります

## 入れ方

```bash
git clone https://github.com/kiyoshiw0128-web/zukai-slides ~/.claude/skills/zukai-slides
pip install python-pptx
```

Claude Code に「この資料を図解のパワポにして。ナレーションも付けて」のように頼むと使われます。

## 中身

| ファイル | 役目 |
|---|---|
| `SKILL.md` | 手順（聞き手を決める → 見出しだけの筋書き → 図を選ぶ → ナレーション → 作る → 検査） |
| `references/storyline.md` | 筋書きの型（結論先出し、問題→原因→対策、空・雨・傘、SCQA） |
| `references/visual-patterns.md` | 伝えたいこと → 図 の選び方 |
| `references/narration.md` | ナレーションの型と、耳で聞いて分かる文の書き方 |
| `references/design-rules.md` | 色・文字・レイアウトの決まり |
| `scripts/zukai_core.py` | デッキ、基本の図形、アイコン（図形だけで描く） |
| `scripts/zukai_parts.py` | 図解の部品（マス、棒、フロー、判断図、前後比較、2x2、時系列、役割カード、やること、表） |
| `scripts/check_deck.py` | 検査（ナレーション、文字の大きさ、はみ出し、文字量）と読み上げ台本の書き出し |
| `scripts/preview_html.py` | PowerPoint がなくても配置を確かめる簡易プレビュー |
| `examples/` | 見本「会議を30分で終える」（数字は架空）と、その読み上げ台本 |

## スクリプトだけで使う

```python
import sys; sys.path.insert(0, "zukai-slides/scripts")
from zukai_core import Deck, CONTENT
from zukai_parts import bars

d = Deck()
sl = d.slide("議題を先に送ると、会議は短くなる",
             "左の棒が、議題を送らなかった会議で平均52分。右が送った会議で31分です。…")
bars(sl, sl.panel(CONTENT, "会議の平均時間（分）", tag="根拠"),
     [("議題なし", 52, "gray"), ("議題を送る", 31, "accent")], unit="分")
sl.band("だから、議題は前日までに送る")
d.save("会議.pptx")
```

```bash
python scripts/check_deck.py 会議.pptx --script 会議_台本.md
python -m pytest tests
```

フォントは [Noto Sans JP](https://fonts.google.com/noto/specimen/Noto+Sans+JP) を前提にしています。ない環境では `Deck(font="BIZ UDPゴシック")` などに変えてください。

## ライセンス

MIT
