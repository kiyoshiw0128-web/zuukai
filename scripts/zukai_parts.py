"""zukai-slides の図解部品。どれも (Slide, Box, ...) を受け取り、Box の中に収めて描く。

どの部品を使うかは references/visual-patterns.md の「伝えたいこと → 図」の表で選ぶ。
"""
from __future__ import annotations

from zukai_core import ICONS, MIN_PT, Box, Slide

STEP_COLORS = ("accent", "brown", "plum", "navy", "gray")


def flow_steps(sl: Slide, b, steps: list[tuple[str, str]], focus: int | None = None) -> None:
    """左から右への手順。steps は (短い名前, 1行の説明)。focus 番目だけ色を変える。"""
    b = Box.of(b)
    y = b.y + max(0.0, (b.h - 2.4) / 2)
    cols = b.split_cols(len(steps), gap=0.5)
    for i, (c, (name, sub)) in enumerate(zip(cols, steps)):
        fill = "accent" if i == focus else "navy"
        sl.label(Box(c.x, y, c.w, 1.3), name, fill, "white", 22)
        sl.text(Box(c.x, y + 1.45, c.w, 0.95), sub, 20, "ink", align="center")
        if i < len(steps) - 1:
            sl.arrow(c.x + c.w + 0.06, y + 0.65, c.x + c.w + 0.44, y + 0.65)


def decision_flow(sl: Slide, b, questions: list[str], outcomes: list[str]) -> list[Box]:
    """はい／いいえの問いで行き先が決まる図。outcomes は len(questions)+1 個。

    問いは上の段に左から並び、「はい」で真下の行き先へ、「いいえ」で右の問いへ進む。
    最後の問いの「いいえ」は一番右の行き先へ下りる。戻り値は行き先の箱（中に絵を足せる）。
    """
    b = Box.of(b)
    n = len(outcomes)
    if n != len(questions) + 1:
        raise ValueError("outcomes は questions より1つ多くしてください")
    cols = b.split_cols(n, gap=0.9)
    qh = 1.05
    out_y = b.y + qh + 0.95
    boxes = []
    for i, c in enumerate(cols):
        ob = Box(c.x, out_y, c.w, b.y + b.h - out_y)
        sl.rect(ob, "paper", rounded=True)
        sl.label(Box(c.cx - 0.3, out_y - 0.3, 0.6, 0.6), str(i + 1),
                 STEP_COLORS[i % len(STEP_COLORS)], "white", 20)
        sl.text(Box(ob.x + 0.1, ob.y + 0.4, ob.w - 0.2, ob.h - 0.5), outcomes[i], 20,
                "ink", True, "center")
        boxes.append(ob)
        if i < len(questions):
            sl.rect(Box(c.x, b.y, c.w, qh), "white", rounded=True, line="navy")
            sl.text(Box(c.x, b.y, c.w, qh), questions[i], 20, "navy", True, "center", "middle")
            sl.arrow(c.cx, b.y + qh, c.cx, out_y - 0.32, "navy")
            sl.text(Box(c.cx + 0.08, b.y + qh + 0.05, 0.8, 0.4), "はい", 18, "navy", True)
            nx = cols[i + 1].x if i + 1 < len(questions) else cols[-1].cx
            if i + 1 < len(questions):
                sl.arrow(c.x + c.w, b.y + qh / 2, nx, b.y + qh / 2, "gray")
            else:
                sl.arrow(c.x + c.w, b.y + qh / 2, nx, b.y + qh / 2, "gray", head=False)
                sl.arrow(nx, b.y + qh / 2, nx, out_y - 0.32, "gray")
            sl.text(Box(c.x + c.w, b.y + qh / 2 - 0.45, 0.9, 0.4), "いいえ", 18, "gray", True, "center")
    return boxes


def ratio_grid(sl: Slide, b, ratio: float, headline: str, note: str = "",
               cols: int = 5, rows: int = 2) -> None:
    """1マス＝1割のように、割合をマスの数で見せる（数字を読まなくても多い少ないが分かる）。"""
    b = Box.of(b)
    total = cols * rows
    filled = round(ratio * total)
    cell = min((b.w - 0.2 * (cols - 1)) / cols, 0.8)
    gx = b.cx - (cell * cols + 0.2 * (cols - 1)) / 2
    for k in range(total):
        r, c = divmod(k, cols)
        sl.rect(Box(gx + c * (cell + 0.22), b.y + r * (cell + 0.22), cell, cell),
                "accent" if k < filled else "white", rounded=True,
                line=None if k < filled else "line")
    ty = b.y + rows * (cell + 0.22) + 0.1
    sl.text(Box(b.x, ty, b.w, 0.6), headline, 24, "ink", True, "center")
    if note:
        sl.text(Box(b.x, ty + 0.65, b.w, 0.45), note, 18, "gray", align="center")


def bars(sl: Slide, b, items: list[tuple[str, float, str]], unit: str = "%",
         vmax: float | None = None) -> None:
    """縦の棒で比べる。items は (ラベル, 値, 色名)。値は棒の上に大きく書く。"""
    b = Box.of(b)
    vmax = vmax or max(v for _, v, _ in items) * 1.15
    cols = b.split_cols(len(items), gap=0.6)
    base = b.y + b.h - 0.55
    for c, (name, v, color) in zip(cols, items):
        h = (base - b.y - 0.65) * v / vmax
        bw = min(c.w, 1.6)
        sl.rect(Box(c.cx - bw / 2, base - h, bw, h), color)
        sl.text(Box(c.x, base - h - 0.6, c.w, 0.55), f"{v:g}{unit}", 28, color, True, "center")
        sl.text(Box(c.x, base + 0.08, c.w, 0.45), name, 20, "ink", True, "center")
    sl.rect(Box(b.x, base, b.w, 0.03), "line")


def before_after(sl: Slide, b, before: tuple[str, list[str]], after: tuple[str, list[str]],
                 arrow_label: str = "") -> None:
    """変える前と後を左右に並べ、真ん中の矢印で「何をしたか」を示す。"""
    b = Box.of(b)
    left, mid, right = Box(b.x, b.y, b.w * 0.42, b.h), Box(b.x + b.w * 0.42, b.y, b.w * 0.16, b.h), \
        Box(b.x + b.w * 0.58, b.y, b.w * 0.42, b.h)
    for box, (head, lines), fill, color in ((left, before, "white", "gray"),
                                            (right, after, "white", "navy")):
        sl.rect(box, fill, rounded=True, line=color)
        sl.label(Box(box.x + 0.2, box.y + 0.2, box.w - 0.4, 0.6), head, color, "white", 22)
        sl.text(Box(box.x + 0.3, box.y + 1.0, box.w - 0.6, box.h - 1.1),
                "\n".join(f"・{t}" for t in lines), 20, "ink")
    sl.arrow(mid.x + 0.15, mid.cy, mid.x + mid.w - 0.15, mid.cy, "accent", 6)
    if arrow_label:
        sl.text(Box(mid.x, mid.cy - 1.0, mid.w, 0.8), arrow_label, 18, "accent", True, "center",
                "bottom")


def matrix_2x2(sl: Slide, b, x_axis: tuple[str, str], y_axis: tuple[str, str],
               cells: list[str], focus: int | None = None) -> None:
    """2つの軸で4つに分ける。cells は 左上・右上・左下・右下 の順。軸は (低い側, 高い側)。"""
    b = Box.of(b)
    g = Box(b.x + 1.7, b.y, b.w - 1.7, b.h - 0.5)
    quads = [Box(g.x + (i % 2) * g.w / 2 + 0.05, g.y + (i // 2) * g.h / 2 + 0.05,
                 g.w / 2 - 0.1, g.h / 2 - 0.1) for i in range(4)]
    for i, (q, t) in enumerate(zip(quads, cells)):
        on = i == focus
        sl.rect(q, "accent" if on else "white", rounded=True, line=None if on else "line")
        sl.text(q, t, 22, "white" if on else "ink", True, "center", "middle")
    sl.text(Box(b.x, g.y, 1.6, 0.5), y_axis[1], 18, "gray", True, "right")
    sl.text(Box(b.x, g.y + g.h - 0.5, 1.6, 0.5), y_axis[0], 18, "gray", True, "right")
    sl.text(Box(g.x, g.y + g.h + 0.05, 2.5, 0.45), x_axis[0], 18, "gray", True)
    sl.text(Box(g.x + g.w - 2.5, g.y + g.h + 0.05, 2.5, 0.45), x_axis[1], 18, "gray", True, "right")


def timeline(sl: Slide, b, events: list[tuple[str, str]], focus: int | None = None) -> None:
    """横一本の線に、いつ・何をを並べる。"""
    b = Box.of(b)
    top = b.y + max(0.0, (b.h - 2.6) / 2)
    y = top + 1.0
    sl.rect(Box(b.x, y - 0.04, b.w, 0.08), "line")
    cols = b.split_cols(len(events), gap=0.2)
    for i, (c, (when, what)) in enumerate(zip(cols, events)):
        color = "accent" if i == focus else "navy"
        sl.oval(c.cx, y, 0.5, color)
        sl.text(Box(c.x, top, c.w, 0.55), when, 24, color, True, "center")
        sl.text(Box(c.x, y + 0.45, c.w, 1.1), what, 22, "ink", True, "center")


def role_cards(sl: Slide, b, cards: list[dict]) -> None:
    """誰が何をするかを並べたカード。cards は {"title", "icons": ["ai","person"],
    "rows": [("AI","下書きする"), ("人","直す")]} の形。"""
    b = Box.of(b)
    who_color = {"AI": "accent", "人": "navy"}
    for i, (c, card) in enumerate(zip(b.split_cols(len(cards), gap=0.3), cards)):
        sl.rect(c, "paper", rounded=True)
        sl.label(Box(c.cx - 0.3, c.y - 0.3, 0.6, 0.6), str(i + 1),
                 STEP_COLORS[i % len(STEP_COLORS)], "white", 20)
        icons = card.get("icons", [])
        isz = 0.75 if c.h >= 3.8 else 0.45
        for k, name in enumerate(icons):
            ix = c.cx + (k - (len(icons) - 1) / 2) * (isz * 1.9)
            ICONS[name](sl, ix, c.y + 0.4 + isz * 1.1, isz)
        y = c.y + 0.6 + isz * 2.2
        if card.get("title"):
            sl.text(Box(c.x + 0.1, y, c.w - 0.2, 0.5), card["title"], 22, "ink", True, "center")
            y += 0.65
        for who, what in card.get("rows", []):
            sl.label(Box(c.x + 0.2, y, 0.62, 0.48), who, who_color.get(who, "gray"), "white", 18)
            sl.text(Box(c.x + 0.92, y, c.w - 1.0, 0.48), what, 20,
                    who_color.get(who, "gray"), True, anchor="middle")
            y += 0.55


def action_list(sl: Slide, b, items: list[tuple[str, str, str]]) -> None:
    """番号つきの「やること」。items は (やること, ひとことの具体, 根拠の場所)。"""
    b = Box.of(b)
    for i, (r, (head, sub, ref)) in enumerate(zip(b.split_rows(len(items), gap=0.15), items)):
        sl.rect(r, "paper", rounded=True)
        sl.label(Box(r.x + 0.15, r.cy - 0.3, 0.6, 0.6), str(i + 1), "navy", "white", 20)
        ref_w = 2.6 if ref else 0
        sl.text(Box(r.x + 0.95, r.y, r.w - 1.1 - ref_w, r.h * 0.55), head, 22, "navy", True,
                anchor="bottom")
        sl.text(Box(r.x + 0.95, r.y + r.h * 0.55, r.w - 1.1 - ref_w, r.h * 0.45), sub, 18, "gray")
        if ref:
            sl.text(Box(r.x + r.w - ref_w, r.y, ref_w - 0.2, r.h), ref, 18, "gray",
                    align="right", anchor="middle")


def big_number(sl: Slide, b, value: str, label: str, color: str = "accent") -> None:
    """ひとつの数字だけを大きく見せる。"""
    b = Box.of(b)
    sl.text(Box(b.x, b.y, b.w, b.h * 0.65), value, 72, color, True, "center", "bottom")
    sl.text(Box(b.x, b.y + b.h * 0.67, b.w, b.h * 0.33), label, 22, "ink", True, "center")


def table(sl: Slide, b, header: list[str], rows: list[list[str]],
          col_ratio: list[float] | None = None) -> None:
    """記入用などの表。header は navy、行は白と薄い色の交互。"""
    b = Box.of(b)
    n = len(header)
    ratio = col_ratio or [1] * n
    total = sum(ratio)
    xs, x = [], b.x
    for r in ratio:
        xs.append((x, b.w * r / total))
        x += b.w * r / total
    rh = min(0.7, b.h / (len(rows) + 1))
    for j, (cx, cw) in enumerate(xs):
        sl.label(Box(cx + 0.02, b.y, cw - 0.04, rh), header[j], "navy", "white", 18)
    for i, row in enumerate(rows, 1):
        fill = "white" if i % 2 else "mist"
        for j, (cx, cw) in enumerate(xs):
            cell = Box(cx + 0.02, b.y + i * rh, cw - 0.04, rh - 0.04)
            sl.rect(cell, fill, line="line")
            sl.text(cell, row[j] if j < len(row) else "", max(MIN_PT, 18), "ink",
                    align="center", anchor="middle")
