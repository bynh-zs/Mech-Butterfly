# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

doc = Document()

# 页面设置 A4 2.5cm
for sec in doc.sections:
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

# 默认字体
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def set_run_font(run, size=12, bold=False, east='宋体'):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), east)

def add_title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run_font(r, 18, True, '黑体')
    p.paragraph_format.space_after = Pt(18)

def add_h1(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, 16, True, '黑体')
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)

def add_h2(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, 14, True, '黑体')
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(5)

def add_body(text, indent=True):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, 12, False, '宋体')
    p.paragraph_format.line_spacing = 1.5
    if indent:
        p.paragraph_format.first_line_indent = Pt(24)
    return p

def add_table(headers, rows):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_run_font(r, 10.5, True, '宋体')
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.rows[ri+1].cells[ci]
            cell.text = ''
            p = cell.paragraphs[0]
            r = p.add_run(val)
            set_run_font(r, 10.5, False, '宋体')
    return t

# ========== 正文 ==========
add_title('Mech-Butterfly 飞控接线手册')

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('——基于蝴蝶飞控原理图与 ESP32-C3 开发板实物')
set_run_font(r, 12, False, '楷体')
p.paragraph_format.space_after = Pt(18)

# 一、系统供电总览
add_h1('一、系统供电总览')
add_body('整套系统由一块 2S 锂电池（7.4V）供电，分两路：')
add_body('第一路（7.4V 直供）：电池正极 → 左右舵机 VCC。舵机是高压版（PTK 7465W，额定 7.4V），直接吃电池电压，不经过降压。')
add_body('第二路（降压后供主控）：电池 7.4V → Mini360 降压模块 → 5V → 开发板 5V 引脚。开发板板载 3.3V 稳压，再转 3.3V 给 ESP32-C3。')
add_body('公共地：电池负极与开发板 GND 必须连在一起，否则 PWM 信号没有参考电平，舵机会乱抖。')

# 二、原理图 8 个模块
add_h1('二、原理图 8 个模块逐个讲')
add_body('打开《蝴蝶飞控原理图.pdf》，整张图共 8 个电路模块。每个模块的归属如下：')

add_table(
    ['编号', '模块名称（图中位置）', '功能', '归属'],
    [
        ['1', 'U5 ESP32-C3 模组（左上角）', '主控芯片，所有 GPIO 引出', '开发板板载，不用管'],
        ['2', 'U4 Micro-USB 座（上中）', '5V 供电 + USB 数据下载', '开发板板载，不用管'],
        ['3', 'P1 左舵机接口 / P2 右舵机接口（右上）', '各 3Pin：电池+、PWM 信号、GND', '洞洞板自己焊 2 个 3P 排针'],
        ['4', 'AMS1117-3V3 LDO（中右）', '5V 转 3.3V 给主控', '开发板板载，不用管'],
        ['5', 'LED 指示灯（左中）', '板载状态灯', '开发板板载（GPIO8），不用管'],
        ['6', 'R20/R21 电池分压（左中）', '300kΩ + 100kΩ 把电池电压降到 ADC 可读范围', '洞洞板自己焊 2 个电阻'],
        ['7', 'J1/J2 电池输入端子（左下）', '电池正、负接入点', '洞洞板自己焊 1 个 2P 插座'],
        ['8', 'D1/Q1 电源路径（右下）', 'USB 与外部供电自动切换', '开发板板载，不用管'],
    ]
)

# 三、开发板板载电路
add_h1('三、开发板板载电路（你什么都不用做）')
add_body('市售 ESP32-C3 开发板已经集成了以下电路，直接用，不用重复搭：')
add_table(
    ['板载电路', '作用', '你需要做什么'],
    [
        ['ESP32-C3 模组（含 WiFi、4MB Flash）', '跑固件、无线通信', '无'],
        ['3.3V LDO 稳压芯片', '把 5V 转成 3.3V 给主控', '无'],
        ['Micro-USB 座', '烧录固件、供电', '烧录时用一次'],
        ['BOOT 键、EN/RST 键', '进下载模式、复位', '烧录/重启时按'],
        ['板载 LED（GPIO8）', '状态指示', '无'],
        ['USB-Serial 电路', '串口下载', '无'],
    ]
)

# 四、洞洞板要焊的 4 样东西
add_h1('四、洞洞板要自己焊的 4 样东西')
add_body('除了开发板，剩下全部电路都在洞洞板上完成，一共 4 样：')
add_table(
    ['序号', '元件', '数量', '作用'],
    [
        ['1', 'Mini360 降压模块', '1', '把电池 7.4V 降到 5V 给开发板供电'],
        ['2', '电池 2P 插座（XH-2Y 母座）', '1', '电池主输出 2Pin 插头插这里，可拔插充电'],
        ['3', '舵机 3P 排针', '2 个', '左、右舵机 JR 母头分别插这里'],
        ['4', '分压电阻 300kΩ + 100kΩ', '各 1', '电池电压采样，中点接 GPIO1'],
    ]
)

# 五、端口对应详表
add_h1('五、端口对应详表（每个连接都写清楚）')
add_body('对着开发板实物丝印，每个引脚接什么、为什么这么接、代码依据，逐条列出：')

add_h2('5.1 舵机信号（2 根）')
add_table(
    ['开发板引脚', '连到哪里', '原因', '代码依据'],
    [
        ['右排丝印 "4"（GPIO4）', '左舵机信号线（橙线）', '固件里左舵机 PWM 从 GPIO4 输出',
         'Kang_Flyer_Ornithopter.py 第 21 行：Servo(pwm_pin=4)'],
        ['右排丝印 "5"（GPIO5）', '右舵机信号线（橙线）', '固件里右舵机 PWM 从 GPIO5 输出',
         'Kang_Flyer_Ornithopter.py 第 25 行：Servo(pwm_pin=5)'],
    ]
)

add_h2('5.2 电池电压采样（1 根信号线 + 2 个电阻）')
add_table(
    ['连接点', '怎么接', '原因', '代码依据'],
    [
        ['电池正极 → 300kΩ 电阻 → GPIO1',
         '300kΩ 一端接电池+，另一端同时接 GPIO1 和 100kΩ 电阻',
         '分压：8.4V 经 300k:100k 分压后，GPIO1 处电压 = 8.4 × 100/(300+100) = 2.1V，落在 ADC 0~3.3V 安全范围',
         'Kang_Battery.py 第 8 行：Battery(pin=1)；第 77 行：×4 还原电池电压'],
        ['100kΩ 电阻 → GND',
         '100kΩ 另一端接开发板 GND',
         '分压下臂，把采样点固定到 0V~2.1V',
         '同上'],
    ]
)

add_h2('5.3 供电（开发板本身的电）')
add_table(
    ['开发板引脚', '连到哪里', '原因'],
    [
        ['左排丝印 "5V"（任意一个）', 'Mini360 模块 OUT+（输出正极）',
         'Mini360 把 7.4V 降到 5V，喂给开发板 5V 脚供电'],
        ['任意 GND', 'Mini360 模块 OUT-（输出负极）+ 电池负极',
         '开发板、Mini360、电池三者共地'],
    ]
)

add_h2('5.4 舵机电源（不走开发板）')
add_table(
    ['连接点', '连到哪里', '原因'],
    [
        ['电池正极（红）', 'Mini360 IN+ + 左右舵机 VCC（红线）',
         '电池 7.4V 直供舵机（舵机额定 7.4V），同时给 Mini360 输入'],
        ['电池负极（黑）', 'Mini360 IN- + 左右舵机 GND（棕线）+ 开发板 GND',
         '公共地'],
    ]
)
add_body('特别注意：舵机的电源（VCC/GND）不经过开发板，直接从电池取电。只有信号线（橙色）才进开发板 GPIO4/GPIO5。原因是舵机电流大（峰值 1A），走开发板 5V 会烧板子。')

# 六、不能碰的引脚
add_h1('六、不能碰的引脚清单')
add_body('ESP32-C3 有几类"禁区脚"，接错会导致开不了机或功能冲突：')
add_table(
    ['引脚', '为什么不能碰'],
    [
        ['GPIO9', 'BOOT 键，上电时的电平决定是否进下载模式，接乱了开不了机'],
        ['GPIO8', '板载 LED 已经焊死在这里'],
        ['GPIO2', '启动 strapping 脚，上电时电平影响启动模式'],
        ['GPIO0', '启动 strapping 脚'],
        ['TX / RX', '烧录串口，被 USB 占用'],
        ['EN', '芯片使能脚'],
        ['3V3', '开发板 3.3V 输出，别接 7.4V 会烧主控'],
    ]
)
add_body('原因：固件已经编译进 firmware.bin，里面写死了 PWM 从 GPIO4/5 出、ADC 从 GPIO1 读。你把线插到别的脚，固件不知道，舵机不动；插到禁区脚，板子直接启动失败。')

# 七、完整接线清单
add_h1('七、完整接线清单（按焊接顺序）')
steps = [
    '第 1 步：洞洞板焊 Mini360 模块。',
    '第 2 步：焊电池 2P 插座（XH-2Y），插座正极接 Mini360 IN+，负极接 Mini360 IN-。',
    '第 3 步：焊 2 个舵机 3P 排针。每个排针：最靠近电池+的一脚接电池+（Vbus_2S），中间脚悬空，最外侧脚接 GND；信号脚分别接 GPIO4（左舵机）、GPIO5（右舵机）。',
    '第 4 步：焊分压电阻：电池+ → 300kΩ → GPIO1；GPIO1 → 100kΩ → GND。',
    '第 5 步：Mini360 OUT+ 接开发板 5V 脚，Mini360 OUT- 接开发板任意 GND。',
    '第 6 步：开发板插排母，插到洞洞板上。',
    '第 7 步：电池 2Pin 主输出插洞洞板电池插座；舵机 JR 母头插 3P 排针。',
    '第 8 步：上电前对照本手册第五节检查表，确认无短路、无接错。',
]
for s in steps:
    p = doc.add_paragraph()
    r = p.add_run(s)
    set_run_font(r, 12, False, '宋体')
    p.paragraph_format.line_spacing = 1.5

# 八、检查清单
add_h1('八、上电前检查清单')
checks = [
    '电池电压：7.4V 左右，用万用表量电池插座正负极，确认不是 0V 或超过 8.4V。',
    'Mini360 输出：先不接开发板，开电用万用表量 Mini360 OUT+ 和 OUT-，确认是 5.0V（误差 4.8~5.2V），不对就调 Mini360 上的小电位器。',
    '共地：电池负极、Mini360 OUT-、开发板 GND 三者通断。',
    'GPIO1 电压：电池接上后，量 GPIO1 对 GND，应该是电池电压 ÷ 4（例如 7.4V → 约 1.85V）。',
    '没有 7.4V 串到 5V/3V3 脚：用万用表通断档，确认电池+和开发板 5V 脚不通（除了 Mini360 之外）。',
    '舵机信号线：左舵机橙线接 GPIO4，右舵机橙线接 GPIO5。',
]
for c in checks:
    p = doc.add_paragraph()
    r = p.add_run('□ ' + c)
    set_run_font(r, 12, False, '宋体')
    p.paragraph_format.line_spacing = 1.5

# 保存
out = r'C:\大三上\宣怀\butterfly\飞控接线手册.docx'
doc.save(out)
print('saved:', out)
