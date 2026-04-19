# TxTools

> Process Simulate (Tecnomatix) secondary development plugins for robotic welding workflows.
>
> 面向机器人焊接工艺的 Tecnomatix Process Simulate 二次开发插件集。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![.NET Framework](https://img.shields.io/badge/.NET-4.8-blue)](https://dotnet.microsoft.com/download/dotnet-framework)
[![C#](https://img.shields.io/badge/C%23-7.3-purple)](https://docs.microsoft.com/dotnet/csharp/)
[![Tecnomatix](https://img.shields.io/badge/Tecnomatix-2402-orange)](https://plm.sw.siemens.com/en-US/tecnomatix/)

---

**Language · 语言**： [English](#english) · [中文](#中文)

---

## English

### Overview

TxTools is a collection of Process Simulate (PS) plugins built on the `Tecnomatix.Engineering` SDK. Each plugin solves a concrete problem that arises in daily robotic welding simulation work — exporting data to CATIA, checking robot reachability, auto-planning welding paths, and more.

The plugins share a common foundation (`PsReader`-style data extraction, a consistent GUI pattern based on `TxForm` + `TxFlexGrid`, and disciplined COM interop with CATIA when needed).

### Plugins

| Plugin | Purpose | Status |
|---|---|---|
| **ExportGun** | Export welding gun + weld-point data from PS to CATIA (CGR placement) or Excel | Stable |
| **DotBall** | Generate spherical geometries at weld points for visualization in CATIA | Stable |
| **RobotReachabilityChecker** | Analyze robot reachability against welding operations, joint limits, and TCP margins | Stable |
| *More to come* | Additional plugins will be added over time | — |

### Requirements

- **Process Simulate**: 2402 (other versions may work but are untested)
- **.NET Framework**: 4.8
- **Language**: C# 7.3 (strictly — no later syntax due to PS hosting constraints)
- **Visual Studio**: 2019 or 2022
- **Optional**: CATIA V5 (R2021 or later) for plugins that bridge to CATIA

### Build

1. Clone this repository.
2. Open the `.sln` in Visual Studio.
3. Add a reference to `Tecnomatix.Engineering.dll` from your Process Simulate installation (e.g. `C:\Program Files\Tecnomatix_2402\eMPower\`).
4. Build in `Release / x64` mode.
5. Copy the resulting `.dll` (together with its folder) into `<Tecnomatix installation>\eMPower\DotNetCommands\<PluginName>\`.

### Installation in Process Simulate

Plugin DLLs live under `Tecnomatix_2402\eMPower\DotNetCommands\<PluginName>\`. Registration is done via the official `CommandReg.exe` utility shipped with Process Simulate.

1. Navigate to `<Tecnomatix installation>\eMPower\` and run `CommandReg.exe`.
2. In the `Register Command` dialog:
   - **Assembly**: click `Browse...` and select your compiled plugin DLL (e.g. `DotNetCommands\TxTools\TxTools.dll`).
   - **Class(es)**: the dialog lists all `TxButtonCommand` classes found in the assembly — tick the ones you want to register (e.g. `ExportGunCmd`, `DeviceZAlignerCmd`, `RobotReachabilityCheckerCmd`).
   - **Product(s)**: tick the PS products where the commands should appear (`Process Simulate` is usually enough; `eM-Review` and `Process Designer` are optional).
   - **File**: pick or create an XML file that will store the registration (e.g. `TxTools.xml`).
   - Click `Register`.
3. Launch Process Simulate.
4. Open `Customize` and find the newly registered commands — drag them onto any toolbar.

To remove commands, re-open `CommandReg.exe`, select the same XML file, and click `Unregister`.

### Coding Conventions

Contributors should read [CONTRIBUTING.md](./CONTRIBUTING.md) first. Key rules:

- **C# 7.3 only.** No `using var`, no default interface implementations, no target-typed `new`, no range operators.
- **Defensive API access.** PS SDK member names vary between versions — use `dynamic` + cascading `try/catch` fallbacks when in doubt.
- **Never guess API members.** Verify against actual PS documentation or runtime logs. Corrections from compiler errors are cheap; silent wrong behavior is expensive.
- **Prefer paraphrasing over verbatim.** Complete file rewrites are preferred over minimal patches when structural issues arise.

### License

[MIT](./LICENSE). See [LICENSE](./LICENSE) for full text.

### Disclaimer

This project is not affiliated with or endorsed by Siemens Digital Industries Software. *Tecnomatix* and *Process Simulate* are trademarks of their respective owners. Use at your own risk in a non-production environment first.

---

## 中文

### 项目简介

TxTools 是一组基于 `Tecnomatix.Engineering` SDK 开发的 Process Simulate (PS) 插件集合。每个插件都解决机器人焊接仿真日常工作中的具体问题——PS 数据导出到 CATIA、机器人可达性分析、焊接路径自动规划等。

各插件共享同一套基础设施（`PsReader` 风格的数据提取层、基于 `TxForm` + `TxFlexGrid` 的统一 GUI 模式，以及必要时对 CATIA 的严谨 COM 互操作）。

### 插件列表

| 插件 | 功能 | 状态 |
|---|---|---|
| **ExportGun（导插枪）** | 将焊枪和焊点数据从 PS 导出到 CATIA（CGR 放置）或 Excel | 稳定 |
| **DotBall（点球）** | 在焊点处生成球形几何体，用于 CATIA 可视化 | 稳定 |
| **RobotReachabilityChecker（可达性验证）** | 分析机器人对焊接操作的可达性、关节限位、TCP 余量 | 稳定 |
| *后续规划* | 将陆续加入更多插件 | — |

### 环境要求

- **Process Simulate**：2402（其他版本未测试）
- **.NET Framework**：4.8
- **语言**：C# 7.3（严格限定——PS 宿主环境约束，不得使用更高版本语法）
- **Visual Studio**：2019 或 2022
- **可选**：CATIA V5（R2021 及以上）——CATIA 相关插件需要

### 编译

1. Clone 本仓库
2. 在 Visual Studio 中打开 `.sln`
3. 从你的 Process Simulate 安装目录引用 `Tecnomatix.Engineering.dll`（例如 `C:\Program Files\Tecnomatix_2402\eMPower\`）
4. 以 `Release / x64` 模式编译
5. 把生成的 `.dll`（含其所属文件夹）放到 `<Tecnomatix 安装目录>\eMPower\DotNetCommands\<插件名>\`

### 在 Process Simulate 中注册插件

插件 DLL 一般位于 `Tecnomatix_2402\eMPower\DotNetCommands\<插件名>\`，注册通过 PS 随附的官方工具 `CommandReg.exe` 完成。

1. 进入 `<Tecnomatix 安装目录>\eMPower\`，运行 `CommandReg.exe`
2. 在 `Register Command` 对话框中：
   - **Assembly**：点 `Browse...` 选中你编译好的插件 DLL（例如 `DotNetCommands\TxTools\TxTools.dll`）
   - **Class(es)**：对话框会列出程序集里所有 `TxButtonCommand` 类——勾选你要注册的命令（例如 `ExportGunCmd`、`DeviceZAlignerCmd`、`RobotReachabilityCheckerCmd`）
   - **Product(s)**：勾选希望命令出现的 PS 产品（一般勾 `Process Simulate` 即可；`eM-Review` 和 `Process Designer` 可选）
   - **File**：选择或新建一个用于保存注册信息的 XML 文件（例如 `TxTools.xml`）
   - 点击 `Register`
3. 启动 Process Simulate
4. 打开"自定义"界面，找到新注册的命令，拖到任一工具栏

需要卸载命令时，重新打开 `CommandReg.exe`，选中同一个 XML 文件，点击 `Unregister` 即可。

### 编码规范

贡献代码前请先读 [CONTRIBUTING.md](./CONTRIBUTING.md)。核心规则：

- **仅允许 C# 7.3 语法**。不得使用 `using var`、默认接口实现、目标类型 `new`、范围运算符等
- **防御式 API 调用**。PS SDK 成员名在不同版本间可能变化——不确定时使用 `dynamic` + 级联 `try/catch` 兜底
- **绝不猜测 API 成员名**。请对照 PS 实际文档或运行时日志验证。编译器报错帮你纠正代价很小；静默的错误行为代价很大
- **优先完整重写，而非局部打补丁**。遇到结构性问题时，完整重写文件通常比多次小修更清晰

### 协议

[MIT](./LICENSE) 协议。完整条款见 [LICENSE](./LICENSE)。

### 免责声明

本项目与 Siemens Digital Industries Software 无任何隶属或背书关系。*Tecnomatix* 和 *Process Simulate* 是各自所有者的商标。请先在非生产环境下自行评估使用风险。

---

## Repository Structure · 仓库结构

```
TxTools/
├── ExportGun/                        # 导插枪插件
│   ├── PsReader.cs                   # PS 数据读取核心
│   ├── ExportGunForm.cs              # 主窗口
│   └── CatiaBridge.cs                # CATIA COM 互操作
├── DotBall/                          # 点球插件
│   └── ...
├── RobotReachabilityChecker/         # 可达性验证插件
│   └── ...
├── Common/                           # 共享工具类（如有）
├── docs/                             # 文档与截图
├── .gitignore
├── LICENSE
├── CONTRIBUTING.md
└── README.md
```

## Contact · 联系

Issues and pull requests are welcome.
欢迎提交 Issue 和 Pull Request。
