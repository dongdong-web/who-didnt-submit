"""生成程序图标：蓝底白勾。改配色改这里，然后重新打包。"""

from PIL import Image, ImageDraw

SIZE = 256
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# 圆角蓝底
d.rounded_rectangle([6, 6, SIZE - 6, SIZE - 6], radius=52, fill=(47, 111, 237, 255))

# 白色对勾
d.line(
    [(66, 134), (112, 180), (192, 84)],
    fill=(255, 255, 255, 255),
    width=24,
    joint="curve",
)

img.save(
    "icon.ico",
    sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
)
print("已生成 icon.ico")
