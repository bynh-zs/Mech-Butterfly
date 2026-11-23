# GPIO 对照表

原理图网络名 ↔ 我的开发板排针丝印 ↔ 实际用途。接线时对着这张表接，不要对着原理图的机械脚号接。

## 一、核心对照（按开发板丝印顺序）

### 左列排针（从上到下）

| 丝印 | GPIO | 原理图网络名 | 用途 | 备注 |
|---|---|---|---|---|
| GND | — | GND | 公共地 | 与所有 GND 内部连通 |
| 3V3 | — | 3V3 | 3.3V 输出 | 板载 LDO 输出，**不要外灌电** |
| 3V3 | — | 3V3 | 同上 | 重复引出 |
| 2 | GPIO2 | IO2 | 通用 IO | 本项目未用 |
| 3 | GPIO3 | IO3 | 通用 IO | 本项目未用 |
| GND | — | GND | 地 | — |
| EN | — | EN/RST | 复位 | 板载已上拉，悬空即可 |
| GND | — | GND | 地 | — |
| 0 | GPIO0 | IO0 → ADC_BAT | 原理图口径的采样脚 | **与代码不符，别接这里** |
| 1 | GPIO1 | IO1 | 电池采样 | **接分压中点到这里** |
| 10 | GPIO10 | IO10 | 通用 IO | 本项目未用 |
| GND | — | GND | 地 | — |
| 5V | — | 5V（板载新增） | 5V 输入 | Mini360 输出接这里；模组本身不出 5V |
| 5V | — | 同上 | 5V 输入 | 重复引出 |
| GND | — | GND | 地 | — |

### 右列排针（从上到下）

| 丝印 | GPIO | 原理图网络名 | 用途 | 备注 |
|---|---|---|---|---|
| GND | — | GND | 地 | — |
| TX | GPIO21 | TXD0 | 串口发送 | 烧录/调试 |
| RX | GPIO20 | RXD0 | 串口接收 | 烧录/调试 |
| GND | — | GND | 地 | — |
| 9 | GPIO9 | IO9/BOOT | 下载模式 | 板载 BOOT 按键已接 |
| 8 | GPIO8 | IO8 → LED | 板载指示灯 | 可能是 RGB 灯，见踩坑 #8 |
| GND | — | GND | 地 | — |
| 7 | GPIO7 | IO7 | 通用 IO | 本项目未用 |
| 6 | GPIO6 | IO6 | 通用 IO | 本项目未用 |
| 5 | GPIO5 | IO5 → PWM1 | **右舵机信号** | 代码 `Servo(pwm_pin=5)` |
| 4 | GPIO4 | IO4 → PWM0 | **左舵机信号** | 代码 `Servo(pwm_pin=4)` |
| GND | — | GND | 地 | — |
| 18 | GPIO18 | IO18 → usb- | USB D- | 经 R5 0Ω，板内已连 |
| 19 | GPIO19 | IO19 → usb+ | USB D+ | 经 R4 0Ω，板内已连 |
| GND | — | GND | 地 | — |

## 二、本项目实际用到的只有 5 个脚

| 用途 | 接哪个丝印 | 依据 |
|---|---|---|
| 左舵机 PWM | `4` | `Kang_Flyer_Ornithopter.py:21` → `Servo(pwm_pin=4)` |
| 右舵机 PWM | `5` | `Kang_Flyer_Ornithopter.py:25` → `Servo(pwm_pin=5)` |
| 电池采样 | `1` | `Kang_Battery.py:8` → `def __init__(self, pin=1, ...)`（主程序 `Battery()` 未传参，用默认值） |
| 供电 | `5V` + `GND` | 板载只有 5V 输入和标准 GND |
| 调试 | USB 口 | 烧录 + 串口 |

其余 IO（2、3、6、7、10、18、19、TX、RX）本项目都用不到，悬空就行。

## 三、这张表是怎么对出来的（依据分层）

不是背出来的现成对照表，是靠四类证据交叉核对：

### ① 图面直读（硬件制图通用规则）

原理图上同名的网络就是连在一起的，不管画在哪。所以对应方式是**名字对名字**，不是脚号对脚号：

- 原理图 `IO4` → 板子丝印 `4`
- 原理图 `TXD0` → 板子丝印 `TX`
- 原理图 `EN/RST` → 板子丝印 `EN`
- 原理图 `3V3` / `GND` → 板子丝印 `3V3` / `GND`

**为什么脚号对不上**：原理图的 12、13、18 是模组**封装机械脚号**（第几个焊脚），板子印的是 **GPIO 编号**（芯片功能编号）。两套体系本来就不一样。

### ② 作者飞控源码（第一手依据）

代码里直接声明了每个脚干什么，这是"哪个脚接舵机/采样"最硬的依据：

| 代码 | 位置 | 结论 |
|---|---|---|
| `Servo(pwm_pin=4)` | `Kang_Flyer_Ornithopter.py:21` | 左舵机 = GPIO4 |
| `Servo(pwm_pin=5)` | `Kang_Flyer_Ornithopter.py:25` | 右舵机 = GPIO5 |
| `def __init__(self, pin=1, ...)` | `Kang_Battery.py:8` | 采样默认 GPIO1 |
| `Self.bat = Battery()` | `Kang_Flyer_Ornithopter.py:17` | 未传参，实际就用 GPIO1 |

### ③ 乐鑫官方资料（联网核实）

- [ESP32-C3-MINI-1 数据手册](https://www.espressif.com.cn/sites/default/files/documentation/esp32-c3-mini-1_datasheet_cn.pdf)：管脚定义、GPIO 复用功能
- [ESP32-C3-DevKitM-1 用户指南](https://docs.espressif.com/projects/esp-dev-kits/zh_CN/latest/esp32c3/esp32-c3-devkitm-1/user_guide.html)：排针位置、板载 LED/按键/USB 连接

核实到的关键点：
- GPIO1 = ADC1_CH1
- GPIO4 = ADC1_CH4
- GPIO5 = ADC2_CH0
- GPIO8 = 板载 RGB LED
- USB D+ = GPIO19，D− = GPIO18

### ④ 通用硬件常识

- NC 脚悬空不接；
- 多个 GND 内部并连，任选其一；
- DevKit 把模组 GPIO 按号引到排针；
- 复位/BOOT 按键、上拉电阻、滤波电容、USB 转串口出厂已焊好；
- 0Ω 电阻在 USB 差分线上起走线预留作用，正常直通。

### ⑤ 还需要自己实测的

我看不到这块第三方兼容板的完整原理图，所以丝印到模组脚的精确连通是靠"同名网络 + 官方同型号板惯例 + 照片丝印"推的。两处落地前要核实：

1. **采样脚口径**：原理图标 IO0，代码写 pin=1。以代码为准接 `1`；接好后万用表量中点电压，和串口读数×4 对比确认。
2. **排针顺序可能与官方 DevKitM-1 不同**：以板面丝印为准。想 100% 确认某根线通到哪个模组脚，用万用表**通断档**：一头戳排针，一头戳模组对应脚，响了就是连通。

## 四、一句话总结

> 规则是通用的（同名相连），骨架来自两张图的丝印，用途来自作者代码，复用功能由官方手册复核；拿不准的脚用万用表量通断，不靠猜。
