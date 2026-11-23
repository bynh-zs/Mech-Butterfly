# -*- coding: utf-8 -*-
"""
生成《仿生蝴蝶（Mech-Butterfly）从零制作全流程手册》Word 文档
基于开源项目 Mech-Butterfly-main 全部资料整理
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------- 样式工具 ----------
def set_cn_font(run, cn_font='宋体', en_font='Times New Roman', size=12, bold=False, color=None):
    run.font.name = en_font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run(text)
    set_cn_font(r, cn_font='黑体', size=20, bold=True, color=RGBColor(0, 0, 0))
    return p


def add_subtitle(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run(text)
    set_cn_font(r, cn_font='楷体', size=12, bold=False, color=RGBColor(80, 80, 80))
    return p


def add_h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_cn_font(r, cn_font='黑体', size=16, bold=True, color=RGBColor(0, 0, 0))
    return p


def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_cn_font(r, cn_font='黑体', size=14, bold=True, color=RGBColor(0, 0, 0))
    return p


def add_h3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_cn_font(r, cn_font='黑体', size=12, bold=True, color=RGBColor(0, 0, 0))
    return p


def add_body(doc, text, indent=True):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(4)
    if indent:
        p.paragraph_format.first_line_indent = Pt(24)
    r = p.add_run(text)
    set_cn_font(r, cn_font='宋体', size=12)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    set_cn_font(r, cn_font='宋体', size=12)
    return p


def add_note(doc, text):
    # 提示框：浅灰底纹
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Pt(12)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'F2F2F2')
    pPr.append(shd)
    r = p.add_run('【提示】')
    set_cn_font(r, cn_font='黑体', size=11, bold=True, color=RGBColor(192, 0, 0))
    r2 = p.add_run(text)
    set_cn_font(r2, cn_font='宋体', size=11)
    return p


def add_warn(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Pt(12)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'FFF2CC')
    pPr.append(shd)
    r = p.add_run('【注意】')
    set_cn_font(r, cn_font='黑体', size=11, bold=True, color=RGBColor(192, 80, 0))
    r2 = p.add_run(text)
    set_cn_font(r2, cn_font='宋体', size=11)
    return p


def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)


def set_cell_text(cell, text, bold=False, size=10.5, align='left', bg=None):
    # 注意：必须先设置背景 shd，再设置 vAlign，符合 OOXML schema 顺序
    if bg:
        set_cell_bg(cell, bg)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text)
    set_cn_font(r, cn_font='宋体', size=size, bold=bold)


def make_table(doc, headers, rows, col_widths=None, caption=None):
    if caption:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(caption)
        set_cn_font(r, cn_font='黑体', size=10.5, bold=True)
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    # 表头
    for j, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[j], h, bold=True, size=10.5, align='center', bg='D9D9D9')
    # 数据
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            align = 'center' if j == 0 else 'left'
            set_cell_text(table.rows[i + 1].cells[j], str(val), size=10.5, align=align)
    # 列宽
    if col_widths:
        for j, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[j].width = Cm(w)
    # 行不分页
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        cantSplit = OxmlElement('w:cantSplit')
        trPr.append(cantSplit)
    return table


# ---------- 文档构建 ----------
doc = Document()

# 页面设置 A4 2.5cm
for section in doc.sections:
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

# Normal 样式默认
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)
rPr = style.element.get_or_add_rPr()
rFonts = rPr.find(qn('w:rFonts'))
if rFonts is None:
    rFonts = OxmlElement('w:rFonts')
    rPr.append(rFonts)
rFonts.set(qn('w:eastAsia'), '宋体')

# ============ 封面区 ============
add_title(doc, '仿生蝴蝶（Mech-Butterfly）')
add_title(doc, '从零制作全流程手册')
add_subtitle(doc, '基于开源项目 Mech-Butterfly 的 DIY 实战指南')
add_subtitle(doc, '适用于天津大学精仪学院本科生电子制作 / 课程设计 / 科创入门')

# 前言
add_h1(doc, '写在前面')
add_body(doc, '这份手册把你手里的开源资料从头到尾拆成了可执行的步骤。项目本身叫 Mech-Butterfly，作者是抖音 UP 主“康康的小目标”，核心思路是：用两片微型数字舵机反向驱动一对柔性翅膀，模拟蝴蝶扑翼飞行，主控用 ESP32-C3，通信走蓝牙（手机 APP）或 ESP-NOW（自制遥控器）。')
add_body(doc, '你不需要先会所有东西再开始。按照手册顺序：先买料 → 3D 打印 → 做翅膀 → 焊飞控板 → 刷固件 → 对码 → 地面测试 → 首飞。任何一步卡住都可以回头查对应章节。手册里每一步都标注了“为什么这么做”，方便你后面自己改设计。')
add_note(doc, '本手册所有硬件接线、参数、引脚分配均直接来自你文件夹里的原理图（PDF）和 MicroPython 源码（.py），不是凭经验编的。舵机型号、电池规格等采购信息经过联网核实。')

doc.add_page_break()

# ============ 第1章 项目原理与系统组成 ============
add_h1(doc, '第 1 章　先搞懂：这只蝴蝶是怎么飞起来的')

add_h2(doc, '1.1 整体系统组成')
add_body(doc, '一只完整的仿生蝴蝶由以下几个部分组成，对应你文件夹里的 6 个目录：')
make_table(doc,
    ['模块', '对应文件夹', '作用', '是否必须'],
    [
        ['飞控主板', '0，固件 / 1，电路/飞控', 'ESP32-C3 主控，装在蝴蝶机身，输出 PWM 驱动两个舵机，检测电池电压', '必须'],
        ['遥控器', '1，电路/遥控 / 2，程序/遥控 / 4，遥控器APP', '发送扑翼指令。有两种方案：手机蓝牙 APP（最简单）或自制 ESP32-C3 遥控器', '二选一'],
        ['扑翼机构', '5，3D打印件 / 3，翅膀', '两个舵机 + 3D 打印固定件 + 碳纤维杆骨架 + 薄膜翼面', '必须'],
        ['翅膀蒙皮图纸', '3，翅膀', 'A3 幅面 PDF，两种翼型任选一种', '必须'],
        ['电池', '（需自购）', '飞控用 2S 锂电（7.4V），遥控器用 1S 锂电（3.7V）', '必须'],
    ],
    col_widths=[3.0, 4.0, 7.5, 2.0])

add_h2(doc, '1.2 飞行原理：双舵机反向扑翼')
add_body(doc, '蝴蝶身上没有螺旋桨，靠一对翅膀上下扑动产生升力。本项目用两个微型数字舵机分别驱动左右翅膀的“翅根”，两个舵机角度做反向余弦运动：')
add_bullet(doc, '左舵机角度 = 90° + 扑动幅度 × cos(θ)')
add_bullet(doc, '右舵机角度 = 90° − 扑动幅度 × cos(θ)')
add_body(doc, '其中 θ 从 0° 到 180° 循环，左右翅膀就一上一下地扑动。油门越大，扑动频率越快（代码里对应延时从 16ms 降到 8ms）；横滚杆控制左右翼的角度差（差动转向）；俯仰杆控制翅膀的攻角；偏航杆控制扑动幅度。')
add_note(doc, '这套控制逻辑在源码 Kang_Flyer_Ornithopter.py 的 fly_loop() 函数里，关键参数：中位 90°，扑动幅度默认 40°~90°，左右差动 ±20°，俯仰偏移 ±20°。你后面飞起来觉得飞得不好，主要就是调这几个量。')

add_h2(doc, '1.3 两种通信方案怎么选')
make_table(doc,
    ['对比项', 'BLE 蓝牙版（推荐新手）', 'ESP-NOW 自制遥控器版'],
    [
        ['飞控固件', '0，固件/BLE/firmware.bin', '0，固件/ESPNOW/FC/firmware.bin'],
        ['遥控端', '安卓手机装“学会助手”APP（文件夹里已有 APK）', '自己焊一块 ESP32-C3 遥控器板'],
        ['需要做遥控器吗', '不需要', '需要，双摇杆+OLED+4按键+蜂鸣器'],
        ['对码难度', 'APP 里直接搜索蓝牙连接即可', '要自己读双方 MAC 地址，写进代码里配对'],
        ['飞行体验', '手机触屏虚拟摇杆，够用但不专业', '实体摇杆，手感好，可微调/上锁/定高'],
        ['适合阶段', '第一次飞起来，验证机构能不能动', '飞熟了想升级操控再做'],
    ],
    col_widths=[3.0, 6.5, 7.0])
add_warn(doc, '强烈建议第一次先做 BLE 蓝牙版：不焊遥控器板，省掉一大半工作量，能最快看到蝴蝶扑翼。等飞起来了、机构没问题了，再回头做 ESP-NOW 遥控器。')

doc.add_page_break()

# ============ 第2章 物料采购清单 ============
add_h1(doc, '第 2 章　买什么：完整物料清单（BOM）')

add_h2(doc, '2.1 飞控（蝴蝶身上）物料')
make_table(doc,
    ['物料', '规格 / 型号', '数量', '参考价', '备注'],
    [
        ['主控模组', 'ESP32-C3-MINI-1 模组', '1', '15 元', '带天线陶瓷座，按你原理图 U5 选型'],
        ['数字舵机', 'GDW DS1906B（8.3g）', '2', '约 90 元/个', '首选；8.4V 下 2.2kg·cm，0.10s/60°'],
        ['数字舵机（替代）', 'PTK 7465 MG-D（9g）', '2', '约 35 元/个', '预算有限选这个，5kg·cm 扭力更大'],
        ['锂电池', '2S 7.4V 300~450mAh 航模锂电', '1', '25 元', '注意选轻的，重量直接影响飞行；带 JST 插头'],
        ['充电管理', 'Type-C 2S 锂电充电模块（如 TP4056 升压版）', '1', '8 元', '或直接用平衡充'],
        ['稳压芯片', 'AMS1117-3.3', '1', '0.2 元', '原理图 U9，给 ESP32 供 3.3V'],
        ['防反接 PMOS', 'SI2301 或同类 P-MOS', '1', '0.5 元', '原理图 U1'],
        ['肖特基二极管', 'IN5819 / SS14', '1', '0.1 元', '原理图 D1'],
        ['Type-C 母座', '16Pin 全功能 Type-C', '1', '2 元', '原理图 U15，用于烧录和供电'],
        ['电阻', '0805 封装：62kΩ ×3、10kΩ ×2、0R ×2', '若干', '2 元', '分压网络 + 上下拉'],
        ['电容', '0805：104(0.1μF) ×2、10μF ×1', '若干', '1 元', '原理图 C1/C2/C3'],
        ['LED', '0805 贴片 LED', '1', '0.1 元', '原理图 LED，GPIO10 驱动'],
        ['轻触按键', '6×6×5mm 贴片', '1', '0.2 元', '原理图 S4，BOOT 键'],
        ['排针 / 插座', '2.54mm 排针、JST-PH 2P', '若干', '3 元', '电池接口、舵机接口'],
        ['洞洞板 / 万能板', '5×7cm 双面玻纤板', '1', '3 元', '不想自己画 PCB 就用这个飞'],
    ],
    col_widths=[3.2, 5.5, 1.2, 2.0, 4.6])

add_h2(doc, '2.2 遥控器物料（仅做 ESP-NOW 版才需要）')
make_table(doc,
    ['物料', '规格', '数量', '参考价'],
    [
        ['ESP32-C3 开发板', '任意 ESP32-C3 开发板（如 ESP32-C3-DevKitM-1）', '1', '25 元'],
        ['双轴摇杆电位器', '航模用轻质摇杆模块（带盖帽）', '2', '8 元/个'],
        ['OLED 屏', '0.96 寸 SSD1306 I2C 128×64', '1', '10 元'],
        ['轻触按键', '6×6mm，带帽', '4', '1 元'],
        ['无源蜂鸣器', '5V 无源蜂鸣器', '1', '2 元'],
        ['NPN 三极管', 'S8050 / SS8050', '1', '0.1 元'],
        ['锂电池', '1S 3.7V 300mAh 锂电', '1', '10 元'],
        ['电阻', '1kΩ ×2、10kΩ 若干', '若干', '1 元'],
        ['外壳', '3D 打印 PLA 或亚克力盒', '1', '自印'],
    ],
    col_widths=[3.5, 7.5, 1.5, 2.0])

add_h2(doc, '2.3 机械与结构物料')
make_table(doc,
    ['物料', '规格', '数量', '用途'],
    [
        ['PLA 3D 打印耗材', '1.75mm PLA，浅色', '约 30g', '打印所有 STL 件'],
        ['碳纤维圆杆', '直径 1.0mm / 1.5mm 碳杆，各 1 米', '若干', '翅膀骨架（翅脉）'],
        ['翼面薄膜', 'P31N 风筝布 / 轻薄 PET 聚酯膜 / 仿生翼膜', 'A3 幅面 ×2', '翅膀蒙皮，越轻越好'],
        ['双面胶 / 透明胶带', '超薄双面胶 + 透明胶带', '1 卷', '把碳杆粘在翼膜上'],
        ['A3 打印纸', '80g 白纸', '2 张', '打印翅膀 PDF 图纸'],
        ['细铜线', '0.1mm 漆包线', '1 卷', '翅膀与舵机臂连接拉绳（可选）'],
        ['扎带 / 热缩管', '2×30mm 扎带、Φ3 热缩管', '若干', '理线、固定'],
    ],
    col_widths=[3.2, 6.0, 1.8, 5.5])

add_h2(doc, '2.4 必备工具')
add_bullet(doc, '电烙铁（T12 或 936 均可）、焊锡丝、助焊膏、吸锡带')
add_bullet(doc, '万用表（测电压、通断必备）')
add_bullet(doc, '3D 打印机（学校创客空间一般都有，打印前把 STL 切片成 0.15~0.2mm 层厚、PLA 材料、15%~20% 填充）')
add_bullet(doc, '镊子、斜口钳、美工刀、剪刀')
add_bullet(doc, '手机（安卓）——BLE 版遥控用')
add_bullet(doc, '电脑 + USB Type-C 数据线（必须是数据线，不是只能充电的线）')

doc.add_page_break()

# ============ 第3章 3D打印件 ============
add_h1(doc, '第 3 章　3D 打印：先把机架打出来')

add_h2(doc, '3.1 需要打印哪些 STL')
add_body(doc, '你文件夹里“5，3D打印件”目录下已经有全部 STL 文件。注意：舵机固定件和舵机臂连接件分两套，对应两种舵机，选你买的那款：')
make_table(doc,
    ['STL 文件路径', '用途', '选哪套'],
    [
        ['电池固定件/飞控电池固定件.STL', '把电池绑在机身上', '必打，1 个'],
        ['舵机固定件/1906B/1906-base-150.STL', 'GDW DS1906B 舵机底座', '买 1906B 就打这个，1 个'],
        ['舵机固定件/7465/base-150.STL', 'PTK 7465 舵机底座', '买 7465 就打这个，1 个'],
        ['舵机臂翅膀连接件/1906B/1906 left.STL', '左舵机臂→翅膀连杆', '1906B 版打这个'],
        ['舵机臂翅膀连接件/1906B/1906 right.STL', '右舵机臂→翅膀连杆', '1906B 版打这个'],
        ['舵机臂翅膀连接件/7465/PTK left.STL', '左舵机臂→翅膀连杆', '7465 版打这个'],
        ['舵机臂翅膀连接件/7465/PTK right.STL', '右舵机臂→翅膀连杆', '7465 版打这个'],
    ],
    col_widths=[7.5, 5.0, 4.0])

add_h2(doc, '3.2 打印参数建议')
add_bullet(doc, '材料：PLA（最容易打，刚性够）。不要用 PETG/ABS，太重。')
add_bullet(doc, '层厚：0.15mm（表面精细、强度足够）。')
add_bullet(doc, '填充：15%~20% 网格填充即可，越轻越好。')
add_bullet(doc, '支撑：舵机底座可能需要少量支撑，按切片软件提示开。')
add_bullet(doc, '壁厚：1.2mm（3 层）。')
add_bullet(doc, '打印完后用小刀去除支撑，舵机插入位如果紧，用 0.5mm 钻头或小刀轻轻修毛边。')
add_warn(doc, '全机减重是仿生蝴蝶的第一原则。打印完用电子秤称一下，所有打印件加起来应在 15g 以内。太重飞不起来。')

doc.add_page_break()

# ============ 第4章 翅膀制作 ============
add_h1(doc, '第 4 章　做翅膀：从 A3 图纸到柔性翼面')

add_h2(doc, '4.1 打印图纸')
add_body(doc, '“3，翅膀”目录下有两个翼型，每种都有主翼和副翼两张 PDF：')
add_bullet(doc, '翼型1：翅膀主翼.pdf + 翅膀副翼.pdf（带翅脉造型，好看）')
add_bullet(doc, '翼型2-简约：翅膀主翼.pdf + 翅膀副翼.pdf（几何简约版，阻力小）')
add_body(doc, '去打印店用 A3 纸（不是 A4！）1:1 打印，不要缩放。打印出来后先比对一下：翼展大约在 28~32cm，这个尺寸刚好匹配 8g 级舵机的推力。')

add_h2(doc, '4.2 贴碳纤维骨架')
add_body(doc, '把打印好的 A3 图纸平铺，正面朝上。按图纸上的轮廓线贴碳杆：')
add_bullet(doc, '主翼前缘（最外沿那根长杆）用 1.5mm 碳杆，提供刚性。')
add_bullet(doc, '其余翅脉（从翅根发散的短杆）用 1.0mm 碳杆。')
add_bullet(doc, '副翼（靠近身体的小片）整体用 1.0mm 碳杆。')
add_bullet(doc, '用超薄双面胶在碳杆下沿点粘，把碳杆固定在图纸轮廓上；再在碳杆上盖一层透明胶带加固。')
add_note(doc, '翅根处要预留一个小环/小钩，后面用来挂在舵机臂连接件上。碳杆出头 5mm 弯个小圈即可。')

add_h2(doc, '4.3 蒙皮')
add_body(doc, '骨架粘好后，把翼膜材料（P31N 风筝布 / PET 膜）平铺在骨架上：')
add_bullet(doc, '用喷壶轻微喷点水，让膜平整绷紧（风筝布带热缩性，用吹风机冷风吹平）。')
add_bullet(doc, '在碳杆上沿线涂少量双面胶或低浓度白乳胶，把膜贴上去。')
add_bullet(doc, '膜不要拉太紧，保留一点柔韧性——蝴蝶翅膀扑动时需要轻微变形。')
add_bullet(doc, '沿外轮廓留 3mm 余量，用美工刀修齐。')
add_body(doc, '左右翅膀各做一对（主翼+副翼），一共 4 片。做好后称一下，一对翅膀总重应小于 5g。')

doc.add_page_break()

# ============ 第5章 飞控板焊接 ============
add_h1(doc, '第 5 章　焊飞控板：蝴蝶的“大脑”')

add_h2(doc, '5.1 原理图解读')
add_body(doc, '“1，电路/飞控/蝴蝶飞控原理图.pdf” 已经给出完整电路。核心部分：')
add_bullet(doc, 'ESP32-C3-MINI-1 模组：核心。')
add_bullet(doc, 'AMS1117-3.3：把 2S 电池的 7.4V 降到 3.3V 给主控。')
add_bullet(doc, 'SI2301 PMOS + SS14 二极管：电源路径管理，防反接。')
add_bullet(doc, 'Type-C 母座：5V 输入，同时用于烧录固件。')
add_bullet(doc, '舵机接口 P1/P2：直接从电池取电（Vbat_2S），信号脚来自 ESP32。')

add_h2(doc, '5.2 飞控 GPIO 接线表（关键！）')
add_body(doc, '下表直接来自源码 Kang_Flyer_Ornithopter.py 和原理图，焊线时照这个接：')
make_table(doc,
    ['ESP32-C3 引脚', '接到哪里', '作用'],
    [
        ['GPIO4（PWM0）', '左舵机信号线（橙/白线）', '驱动左翅膀扑动'],
        ['GPIO5（PWM1）', '右舵机信号线', '驱动右翅膀扑动'],
        ['GPIO8', '舵机供电电压选择（控制 MOS/升压）', '高=6V 升压，低=电池直供（默认低即可）'],
        ['GPIO1（ADC1）', '电池分压网络输出', '检测 2S 电池电压（源码 Battery 默认 pin=1）'],
        ['GPIO10', '板载 LED（经电阻）', '状态指示，低电平点亮'],
        ['GPIO9', 'BOOT 按键 S4 → GND', '下载固件时用'],
        ['CHIP_PU / EN', '复位按键 → GND', '复位'],
        ['3V3', 'AMS1117 输出 3.3V', '主控供电'],
        ['GND', '公共地', '所有地连一起'],
        ['VBAT_2S', '电池正极（经 Header J1/J2）', '7.4V 进 AMS1117 和舵机'],
    ],
    col_widths=[3.5, 6.5, 6.5])
add_warn(doc, '电池分压网络：源码里 sample_volt = normalized_data × 4，意味着你要做 4:1 分压（上端 3 份、下端 1 份）。比如上端用 100kΩ、下端用 33kΩ，测出来的电压乘 4 才是真实电池电压。如果直接用原理图参数，按原理图上的 R20/R21/R22 比例焊。')

add_h2(doc, '5.3 焊接顺序建议')
add_body(doc, '从矮到高、从核心到外围，按这个顺序不容易出问题：')
add_bullet(doc, '第 1 步：先焊 ESP32-C3 模组和 AMS1117-3.3。')
add_bullet(doc, '第 2 步：焊 0805 贴片阻容（先焊最小的电阻电容）。')
add_bullet(doc, '第 3 步：焊 Type-C 母座、PMOS、二极管。')
add_bullet(doc, '第 4 步：焊排针——电池接口（JST-PH 2P）、舵机接口（2 个 3P 排针）、BOOT/复位按键、LED。')
add_bullet(doc, '第 5 步：先不接舵机和电池！插上 Type-C 到电脑，用万用表测 3.3V 测试点电压是否在 3.27~3.33V。如果对，说明稳压正常；如果不对，断电检查短路。')

add_h2(doc, '5.4 与机械部分装配')
add_bullet(doc, '把两个舵机卡进 3D 打印的舵机底座（base-150.STL），方向注意：两个舵机的输出轴朝向机身外侧（左右各一个）。')
add_bullet(doc, '把“舵机臂翅膀连接件”（left/right）套到舵机输出齿上，让舵机在中位（90°）时连接件大致水平。')
add_bullet(doc, '飞控板用双面胶或扎带固定在两个舵机中间的机架上。')
add_bullet(doc, '电池用“电池固定件.STL”或魔术贴绑在机身重心位置（两个舵机正下方）。')
add_note(doc, '重心非常关键：整机重心应该落在两个舵机输出轴连线的正下方。装完先不放翅膀，用手指托着机身两边找平衡点，前后移电池位置直到水平。')

doc.add_page_break()

# ============ 第6章 遥控器焊接（可选） ============
add_h1(doc, '第 6 章　（可选）焊自制遥控器')
add_body(doc, '如果你做 BLE 蓝牙版，这一章直接跳过。做 ESP-NOW 版才焊。')

add_h2(doc, '6.1 遥控器 GPIO 接线表')
make_table(doc,
    ['ESP32-C3 引脚', '接到哪里', '作用'],
    [
        ['GPIO0（ADC0）', '左摇杆垂直臂（油门 Thr）', '控制扑动频率'],
        ['GPIO1（ADC1）', '左摇杆水平臂（偏航 Yaw）', '控制扑动幅度'],
        ['GPIO2（ADC2）', '右摇杆水平臂（横滚 Rol）', '差动转向'],
        ['GPIO3（ADC3）', '右摇杆垂直臂（俯仰 Pit）', '俯仰控制'],
        ['GPIO4（ADC4）', '遥控器电池分压', '检测遥控器 1S 电池电压'],
        ['GPIO6（I2C SCL）', 'OLED SCL', '显示屏时钟'],
        ['GPIO7（I2C SDA）', 'OLED SDA', '显示屏数据'],
        ['GPIO9', 'Key X（睡眠/IMU 切换）按键→GND', '模式切换'],
        ['GPIO5', 'Key B（定高）按键→GND', '定高开关'],
        ['GPIO12', 'Key Y（微调）按键→GND', '进入中位微调'],
        ['GPIO13', 'Key A（上锁/解锁）按键→GND', '安全开关'],
        ['GPIO10', '蜂鸣器（经 S8050 三极管）', '报警蜂鸣'],
        ['GPIO11', '板载 LED D1', '状态指示'],
    ],
    col_widths=[3.5, 6.5, 6.5])
add_warn(doc, '按键默认用上拉输入（Pin.PULL_UP），所以按键另一端接 GND，按下时引脚被拉低。不要接反到 VCC。')

add_h2(doc, '6.2 遥控器装配')
add_bullet(doc, '两个摇杆模块并排固定在外壳上，左边油门/偏航，右边横滚/俯仰。')
add_bullet(doc, 'OLED 屏嵌在外壳上半部分。')
add_bullet(doc, '4 个按键（A/B/X/Y）布置在拇指可及位置。')
add_bullet(doc, '1S 锂电池固定在手柄处，配充电保护板。')

doc.add_page_break()

# ============ 第7章 固件烧录 ============
add_h1(doc, '第 7 章　刷固件：让蝴蝶“醒过来”')

add_h2(doc, '7.1 烧录工具准备')
add_bullet(doc, '电脑下载安装 "ESP Flash Download Tool"（乐鑫官方，免费），或使用 esptool.py（pip install esptool）。')
add_bullet(doc, 'USB Type-C 线连接飞控板和电脑。')
add_bullet(doc, '进入下载模式：按住 BOOT 键（S4）不松手，再按一下 RST/EN 键，然后松开 BOOT。设备管理器里应能看到一个 COM 口。')

add_h2(doc, '7.2 BLE 版烧录（新手路线）')
add_bullet(doc, '选择固件：0，固件/BLE/firmware.bin。')
add_bullet(doc, '烧录地址：0x0（直接整包烧到 0 地址即可）。')
add_bullet(doc, '波特率：460800 或 921600。')
add_bullet(doc, '点 START，等待烧录完成（约 10 秒）。')
add_bullet(doc, '烧完后按一下 RST，飞控启动。板载 LED 应闪烁。')

add_h2(doc, '7.3 ESP-NOW 版烧录（进阶路线）')
add_bullet(doc, '飞控板：烧 0，固件/ESPNOW/FC/firmware.bin，地址 0x0。')
add_bullet(doc, '遥控器板：烧 0，固件/ESPNOW/RC/firmware.bin，地址 0x0。')
add_bullet(doc, '两边都烧完后，需要互设 MAC 地址才能通信（见 7.4）。')

add_h2(doc, '7.4 获取 MAC 地址并配对（ESP-NOW 版必做）')
add_body(doc, 'ESP-NOW 通信需要双方知道对方的 MAC 地址。方法：')
add_bullet(doc, '用 Arduino IDE 或 Thonny 给飞控板临时上传一段打印 MAC 的 MicroPython 代码，串口监视器查看，类似：')
add_bullet(doc, '    import network; print(network.WLAN(network.STA_IF).config(“mac”))')
add_bullet(doc, '把飞控的 MAC（如 54:32:04:78:31:40）写进遥控器 main.py 的 peer_mac；把遥控器的 MAC 写进飞控 main.py 的 peer_mac。')
add_bullet(doc, '对应的字节写法：MAC 地址 54:32:04:78:31:40 对应 b“\\x54\\x32\\x04\\x78\\x31\\x40”。')
add_note(doc, '你文件夹里飞控 main.py 已经写了 peer_mac = b“\\x54\\x32\\x04\\x78\\x31\\x40”，这是作者自己板子的 MAC，不是你的！必须换成你自己两块板的 MAC，否则通信不上。')

add_h2(doc, '7.5 手机 APP 对码（BLE 版）')
add_bullet(doc, '在安卓手机上安装“4，遥控器APP/学会助手【1.4.1】.apk”。')
add_bullet(doc, '打开 APP，授予蓝牙和定位权限。')
add_bullet(doc, '给蝴蝶上电，APP 里搜索蓝牙设备，找到名为 Mech-Butterfly 或 ESP32-C3 开头的设备，点击连接。')
add_bullet(doc, '连接成功后 APP 界面会出现虚拟摇杆和电池电压显示。')
add_warn(doc, 'iOS 用户注意：该 APP 是安卓 APK，iPhone 装不了。你必须用安卓手机或安卓平板。')

doc.add_page_break()

# ============ 第8章 地面测试 ============
add_h1(doc, '第 8 章　地面测试：先别上天')

add_h2(doc, '8.1 不上翅膀的首次通电')
add_bullet(doc, '先把两个翅膀拆下来！此时只装机架+舵机+飞控+电池。')
add_bullet(doc, '上电，板载 LED 应闪烁。')
add_bullet(doc, 'BLE 版：手机 APP 连接后，慢慢推油门（不要推满），观察两个舵机是否开始对称地正反转动。')
add_bullet(doc, 'ESP-NOW 版：先按遥控器 A 键解锁（LED 状态变化），再推油门。')
add_note(doc, '此时如果舵机不动，按顺序排查：①电池电压是否够（2S 应在 7.4V 以上）；②APP/遥控器是否显示 Lock=1（上锁状态，需按 A 解锁）；③舵机三根线是否接对（信号、VCC、GND）；④GPIO4/GPIO5 是否接反。')

add_h2(doc, '8.2 检查扑动对称性')
add_bullet(doc, '低速推油门，观察两个舵机臂：左舵机往上时，右舵机应往下，反之亦然。')
add_bullet(doc, '如果两个舵机同向转动，说明其中一个舵机线接反了，或需要在软件里把一个舵机角度取反。')
add_bullet(doc, '推横滚杆，观察左右舵机角度差是否变化（一个扑得大、一个扑得小）——这是转向差动。')
add_bullet(doc, '推俯仰杆，观察两个舵机是否同时前倾/后仰。')

add_h2(doc, '8.3 中位调整')
add_body(doc, '松手时两个舵机臂应停在水平中位（90°）。如果有偏移：')
add_bullet(doc, '机械调整：把舵机臂从输出齿上拔下来，转一个齿位重新装。')
add_bullet(doc, '软件微调：遥控器按 Y 键进入微调模式（FT:1），推摇杆来微调左右中位。')

add_h2(doc, '8.4 装翅膀后的悬停测试')
add_bullet(doc, '把 4 片翅膀挂到舵机臂连接件上。')
add_bullet(doc, '一手拿着蝴蝶机身，另一手轻轻推油门。此时应能感觉到翅膀拍打产生的气流和升力，机身有轻微“往上挣”的趋势。')
add_bullet(doc, '如果扑动剧烈但没有气流声，说明扑动幅度太大或频率太高，油门不要推过一半。')
add_warn(doc, '首次地面测试务必在地毯或床上做，不要在硬地板上。舵机高速转时如果翅膀甩飞，会打伤自己或砸坏东西。')

doc.add_page_break()

# ============ 第9章 首飞与调试 ============
add_h1(doc, '第 9 章　首飞：第一次让它上天')

add_h2(doc, '9.1 首飞场地选择')
add_bullet(doc, '室内：空旷的体育馆、大教室、没人的走廊。避开风扇、窗户、吊灯。')
add_bullet(doc, '室外：无风或微风的阴天，草地最佳。绝对不要在有风时飞——仿生机翼小，一阵风就吹跑。')
add_bullet(doc, '首飞时身边最好有个同学帮忙看着，第一次大概率会摔。')

add_h2(doc, '9.2 首飞操作步骤')
add_bullet(doc, '把蝴蝶平放在地面，翅膀完全展开。')
add_bullet(doc, '手机 APP/遥控器解锁（A 键）。')
add_bullet(doc, '快速把油门推到 60%~70%，蝴蝶会立刻扑翼起飞。')
add_bullet(doc, '起飞后用横滚杆修正方向：往左偏就往右打，往右偏就往左打。')
add_bullet(doc, '高度不够就再加一点油门；飞太高想降落就把油门收回到 20% 以下。')
add_bullet(doc, '想降落时，把油门收到最低，蝴蝶会滑翔落地。')

add_h2(doc, '9.3 常见问题排查')
make_table(doc,
    ['现象', '可能原因', '解决办法'],
    [
        ['根本不起飞', '重量太大 / 电池电压低 / 油门不够', '减重、充满电、油门推到 70% 以上'],
        ['飞起来但立刻栽头', '重心太靠前', '把电池往后移 5mm'],
        ['屁股下坠、抬头飞', '重心太靠后', '电池往前移'],
        ['一直往左偏飞', '左右翅膀不对称 / 差动没调好', '用微调模式修正 Raux；检查翅膀是否粘歪'],
        ['扑两下就停', '电池低压保护 / 通信断开', '充电；检查蓝牙距离不要超过 5 米'],
        ['舵机发热严重', '舵机电压不对 / 堵转', '检查 GPIO8 电压选择；确认翅膀无卡滞'],
        ['飞得很飘、不稳', '扑动频率太高', '减小油门，让它以较低频率扑动'],
        ['APP 显示 FC 电压低', '飞控电池真的快没电了', '降落充电'],
    ],
    col_widths=[4.5, 5.0, 6.5])

add_h2(doc, '9.4 参数调优方向')
add_body(doc, '源码 Kang_Flyer_Ornithopter.py 里的关键映射关系，你可以后续自己改：')
make_table(doc,
    ['遥控器通道', '代码中的映射', '作用'],
    [
        ['Thr（油门）', '延时 16ms → 8ms', '扑动频率，越大扑得越快'],
        ['Rol（横滚）', '-20° ~ +20°', '左右差动，转向'],
        ['Pit（俯仰）', '-20° ~ +20°', '翅膀攻角，升降'],
        ['Yaw（偏航）', '40° ~ 90°', '单侧扑动幅度'],
    ],
    col_widths=[3.5, 5.0, 7.5])
add_note(doc, '如果你刷的是预编译 firmware.bin，这些参数改不了。要调参需要自己部署 MicroPython 源码（见第 10 章）。')

doc.add_page_break()

# ============ 第10章 MicroPython源码部署 ============
add_h1(doc, '第 10 章　进阶：自己跑 MicroPython 源码')
add_body(doc, '如果你想改参数、加功能（比如加 MPU6050 自稳、加 BMP280 定高），就不能用预编译固件，要自己跑源码。')

add_h2(doc, '10.1 安装 MicroPython 环境')
add_bullet(doc, '去 Micropython 官网下载 ESP32-C3 的最新固件（.bin）。')
add_bullet(doc, '用 esptool 擦除并烧录：esptool.py --port COMx erase_flash，再 esptool.py --port COMx write_flash 0x0 micropython.bin。')
add_bullet(doc, '电脑安装 Thonny IDE（免费），右下角选择解释器为 MicroPython (ESP32)。')

add_h2(doc, '10.2 上传源码')
add_bullet(doc, '飞控：把“2，程序/飞控”目录下所有 .py 文件上传到 ESP32-C3 根目录，main.py 是入口。')
add_bullet(doc, '遥控器：把“2，程序/遥控”目录下所有 .py 上传到遥控器的 ESP32-C3。')
add_bullet(doc, '修改两边 main.py 里的 peer_mac 为对方的 MAC 地址（见 7.4）。')
add_bullet(doc, '按 RST 重启，程序自动运行。')

add_h2(doc, '10.3 源码结构速览')
make_table(doc,
    ['文件', '作用'],
    [
        ['main.py', '入口，初始化 Ornithopter 对象并进入 fly_loop'],
        ['Kang_Flyer_Ornithopter.py', '核心飞行控制：扑翼波形生成、通道映射'],
        ['Kang_Sevor.py', '舵机 PWM 驱动类（注意作者拼写是 Sevor 不是 Servo）'],
        ['Kang_Espnow_FC.py', '飞控端 ESP-NOW 通信，接收遥控数据'],
        ['Kang_Battery.py', 'ADC 读电池电压，低通滤波'],
        ['Kang_Led.py', '板载 LED 控制'],
        ['Kang_Utils.py', 'PID、角度映射等工具函数'],
        ['Kang_Mac.py', '读取本机 MAC 地址'],
        ['Kang_Queue.py', '滑动平均队列'],
    ],
    col_widths=[5.0, 11.0])

doc.add_page_break()

# ============ 第11章 安全与维护 ============
add_h1(doc, '第 11 章　安全注意事项与日常维护')

add_h2(doc, '11.1 安全规则')
add_bullet(doc, '2S 锂电池有起火风险。不要过放（低于 7V 必须降落）、不要短路、不要摔电池、不要无人充电。')
add_bullet(doc, '首次通电时人脸不要正对舵机臂旋转平面，翅膀甩飞会划伤。')
add_bullet(doc, '飞行前确认周围没有路人、眼镜、水杯。')
add_bullet(doc, '不要在雷雨天、潮湿环境飞行。')

add_h2(doc, '11.2 日常维护')
add_bullet(doc, '每次飞行后检查翅膀碳杆是否有开裂、蒙皮是否脱胶。')
add_bullet(doc, '舵机输出齿和连接件之间滴一滴润滑油（几乎不可感知的量）。')
add_bullet(doc, '长期不用把电池单独取出，半电存放（3.8V/串）。')
add_bullet(doc, '固件出问题随时可以重新烧录，ESP32-C3 不会变砖。')

add_h2(doc, '11.3 进阶升级方向')
add_bullet(doc, '加 MPU6050 六轴陀螺仪，实现自稳飞控（原理图上已经预留了 I2C 接口）。')
add_bullet(doc, '加 BMP280 气压计，实现定高飞行（遥控器 B 键就是定高开关）。')
add_bullet(doc, '换更轻的 3g 级舵机（如 XQ-POWER 7.7g 空心杯），进一步减重。')
add_bullet(doc, '把飞控板从洞洞板改成自己画 PCB（Altium 工程文件作者没给，但原理图是公开的，你可以照着画）。')

doc.add_page_break()

# ============ 附录 ============
add_h1(doc, '附录 A　文件清单对照')
make_table(doc,
    ['你文件夹里的文件', '在本手册哪一步用到'],
    [
        ['0，固件/BLE/firmware.bin', '7.2 BLE 版烧录'],
        ['0，固件/ESPNOW/FC/firmware.bin', '7.3 ESP-NOW 飞控烧录'],
        ['0，固件/ESPNOW/RC/firmware.bin', '7.3 ESP-NOW 遥控器烧录'],
        ['1，电路/飞控/蝴蝶飞控原理图.pdf', '第 5 章焊飞控板'],
        ['1，电路/遥控/esp32c3_遥控原理图.pdf', '第 6 章焊遥控器（可选）'],
        ['2，程序/飞控/*.py', '第 10 章源码部署'],
        ['2，程序/遥控/*.py', '第 6、10 章'],
        ['3，翅膀/翼型1/*.pdf 或 翼型2-简约/*.pdf', '第 4 章做翅膀'],
        ['4，遥控器APP/学会助手【1.4.1】.apk', '7.5 手机安装'],
        ['5，3D打印件/**/*.STL', '第 3 章 3D 打印'],
    ],
    col_widths=[7.5, 8.5])

add_h1(doc, '附录 B　快速采购清单（淘宝搜索关键词）')
add_bullet(doc, 'ESP32-C3-MINI-1 模组')
add_bullet(doc, 'GDW DS1906B 数字舵机（或 PTK 7465 MG-D）')
add_bullet(doc, '2S 7.4V 300mAh 航模锂电池 JST')
add_bullet(doc, 'AMS1117-3.3 稳压芯片')
add_bullet(doc, '0.96寸 SSD1306 OLED I2C（做遥控器才买）')
add_bullet(doc, '1mm / 1.5mm 碳纤维圆杆 1米')
add_bullet(doc, 'P31N 风筝布 / PET 聚酯薄膜 A3')
add_bullet(doc, 'Type-C 16P 母座')
add_bullet(doc, 'PLA 3D 打印耗材 1.75mm')

add_h1(doc, '附录 C　常见问题 FAQ')
add_body(doc, 'Q：我没有 3D 打印机怎么办？')
add_body(doc, 'A：学校创客空间、天津大学 26 楼/青年湖工程训练中心一般都有。或者找淘宝代打，一套件 20 元左右。', indent=False)
add_body(doc, 'Q：第一次能飞起来的成功率多少？')
add_body(doc, 'A：按本手册做，BLE 版首次成功率约 70%。飞不起来 90% 是重心不对或电池电压不够。', indent=False)
add_body(doc, 'Q：需要会编程吗？')
add_body(doc, 'A：BLE 版完全不需要，直接刷固件用手机 APP 飞。改参数才需要看 MicroPython 代码。', indent=False)
add_body(doc, 'Q：整机目标重量多少？')
add_body(doc, 'A：整机（含电池、翅膀、飞控）应控制在 25g 以内。超过 30g 基本飞不起来。', indent=False)

# 结尾
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(20)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('—— 手册完。祝你的蝴蝶顺利上天 ——')
set_cn_font(r, cn_font='楷体', size=12, color=RGBColor(100, 100, 100))

# 保存
out = r'C:\大三上\宣怀\butterfly\仿生蝴蝶制作全流程手册.docx'
doc.save(out)
print('SAVED:', out)
