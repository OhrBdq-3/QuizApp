# -*- coding: utf-8 -*-
"""设计令牌：间距、圆角、字号。

颜色不放在这里——颜色要跟随浅色/深色主题和主题色预设，由 settings.py 的
THEME_PALETTES 提供（见 app._palette()）。这里只放与主题无关的"尺子"，
让界面各处的留白、圆角、字号有统一刻度，避免每次随手填数字。
"""

# ---------- 间距（4 的倍数，越小越紧） ----------
SP_1 = 4
SP_2 = 8
SP_3 = 12
SP_4 = 16
SP_5 = 20
SP_6 = 24
SP_8 = 32
SP_10 = 40

# ---------- 圆角 ----------
# 整体走「暖纸质 + 大圆角」：卡片/面板明显更圆，按钮、徽标跟着放大一档。
R_SM = 8          # 小控件、输入框外壳、徽标
R_MD = 13         # 选项行、次级面板
R_LG = 18         # 卡片、大面板
R_PILL = 999      # 胶囊

# ---------- 字号（逻辑值，会再乘 FONT_SCALE） ----------
FS_MICRO = 9      # 角标、图例
FS_META = 10      # 次要说明
FS_BODY = 11      # 正文
FS_STRONG = 12    # 卡片标题
FS_TITLE = 14     # 区块标题
FS_Q = 17         # 题干
FS_H1 = 24        # 页面大标题


def _parts(hex_color):
    h = str(hex_color).lstrip("#")
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    return [int(h[i:i + 2], 16) for i in (0, 2, 4)]


def shade(hex_color, factor):
    """factor>1 变暗，factor<1 变亮（和 settings._shade 同一套语义）。

    放在这里是因为 Canvas 自绘的组件（卡片投影）要按父容器底色现场派生颜色，
    而 widgets.py 不该反向依赖 settings.py。
    """
    r, g, b = _parts(hex_color)
    if factor >= 1:
        r, g, b = (int(c * (2 - factor)) for c in (r, g, b))
    else:
        r, g, b = (int(c + (255 - c) * (1 - factor)) for c in (r, g, b))
    r, g, b = (max(0, min(255, c)) for c in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"


def luminance(hex_color):
    """亮度 0=黑 1=白，用来判断底色深浅（投影色要不要往下压）。"""
    r, g, b = _parts(hex_color)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


def shadow_of(hex_color):
    """由容器底色派生投影色：暖底压深一档，深色底直接吃黑。"""
    if luminance(hex_color) < 0.5:
        return "#000000"
    return shade(hex_color, 1.09)


def rounded_rect(canvas, x1, y1, x2, y2, r=8, **kw):
    """在 Canvas 上画圆角矩形，返回 item id。

    Tk 的 Canvas 没有圆角矩形图元，这里用带 smooth 的多边形近似：
    四个角各留 r 的余量，splinesteps 控制平滑度。轮廓（outline）
    会沿着平滑后的路径画，所以填充和描边都是圆角的。
    """
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    if r <= 0.5:
        return canvas.create_rectangle(x1, y1, x2, y2, **kw)
    if r * 2 >= min(x2 - x1, y2 - y1) - 0.5:
        # 半径顶到边了：直接画椭圆，比平滑多边形更规整
        return canvas.create_oval(x1, y1, x2, y2, **kw)
    pts = [
        x1 + r, y1, x2 - r, y1,          # 上边
        x2, y1, x2, y1 + r,              # 右上角
        x2, y2 - r,                      # 右边
        x2, y2, x2 - r, y2,              # 右下角
        x1 + r, y2,                      # 下边
        x1, y2, x1, y2 - r,              # 左下角
        x1, y1 + r,                      # 左边
        x1, y1,                          # 左上角，回到起点
    ]
    kw.setdefault("smooth", True)
    kw.setdefault("splinesteps", 16)
    return canvas.create_polygon(pts, **kw)


def rounded_rect_poly(canvas, x1, y1, x2, y2, r=8, **kw):
    """和 rounded_rect 一样画圆角矩形，但**始终**是多边形。

    尺寸要随布局变化、之后还要用 rounded_rect_item() 改坐标的图元必须用这个：
    rounded_rect 在半径顶到边时会退化成 create_oval，而椭圆 item 只吃 4 个
    坐标点，再想把它改成 12 点的圆角矩形是改不动的（图形会停在初始尺寸上）。
    """
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    if r <= 0.5:
        return canvas.create_polygon(x1, y1, x2, y1, x2, y2, x1, y2, **kw)
    pts = [
        x1 + r, y1, x2 - r, y1,
        x2, y1, x2, y1 + r,
        x2, y2 - r,
        x2, y2, x2 - r, y2,
        x1 + r, y2,
        x1, y2, x1, y2 - r,
        x1, y1 + r,
        x1, y1,
    ]
    kw.setdefault("smooth", True)
    kw.setdefault("splinesteps", 16)
    return canvas.create_polygon(pts, **kw)


def rounded_rect_item(canvas, item, x1, y1, x2, y2, r=8):
    """把已有的多边形 item 重设为新的圆角矩形（切题时复用 item，省去增删）。"""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    if r <= 0.5:
        # 仍是多边形：点数必须一致，不能只给 4 个点（那是矩形的坐标）
        canvas.coords(item, x1, y1, x2, y1, x2, y2, x1, y2)
        return
    pts = [
        x1 + r, y1, x2 - r, y1,
        x2, y1, x2, y1 + r,
        x2, y2 - r,
        x2, y2, x2 - r, y2,
        x1 + r, y2,
        x1, y2, x1, y2 - r,
        x1, y1 + r,
        x1, y1,
    ]
    canvas.coords(item, *pts)
