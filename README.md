# TxTools 使用与开发手册（中文）

版本：2026-06-22  
说明：本文件基于仓库现有代码与模块说明整理，面向插件使用者与二次开发者。建议将本文件保留在仓库根目录以便查阅。

---

## 一、项目概述
TxTools 是面向机器人焊接工艺的 Process Simulate（Tecnomatix）二次开发插件集合。基于 `Tecnomatix.Engineering` SDK，提供一组常用功能以辅助焊接仿真与后处理，包括：焊枪/焊点导出、焊点可视化（点球）、机器人可达性检查、由曲线生成实体、沿线生成围栏等。

设计要点：
- 在不同 PS 版本下采用防御式 API 探测（dynamic + 多路径 try/catch）。
- 统一的 GUI 风格：基于 TxForm + TxFlexGrid 的卡片式面板与日志。
- 批量操作包裹在 UndoScope 中，支持 Ctrl+Z 回滚。
- 对外部系统（如 CATIA）使用桥接层以集中管理 COM 互操作。

---

## 二、环境与依赖
- Process Simulate：推荐 2402（其它版本可能可用但未全面测试）。
- .NET Framework：4.8（与 PS 宿主匹配）。
- C# 语言：C# 7.3（严格限制，避免使用新语法）。
- Visual Studio：2019 / 2022。
- 可选：CATIA V5（用于 CATIA 相关功能的桥接）。

编译时常用引用：
- Tecnomatix.Engineering.dll
- Tecnomatix.Engineering.Ui.dll
- System.Windows.Forms、System.Drawing 等

编译目标：Release / x64，输出为 DLL 并连同资源一起部署到 PS 插件目录。

---

## 三、快速安装与注册
1. 克隆仓库并在 Visual Studio 中打开 `.sln`。  
2. 为项目添加 PS 安装目录下的 `Tecnomatix.Engineering.dll` 等引用（例如：C:\Program Files\Tecnomatix_2402\eMPower\）。  
3. 选择 Release / x64 并生成 DLL。  
4. 将生成的插件文件夹（DLL + Resources）复制到：  
   `<Tecnomatix 安装目录>\eMPower\DotNetCommands\<PluginName>\`  
5. 进入 `<Tecnomatix 安装目录>\eMPower\`，运行 `CommandReg.exe`：  
   - Assembly：选择插件 DLL。  
   - Class(es)：勾选要注册的 `TxButtonCommand` 类（如 ExportGunCmd、LineToSolidCommand 等）。  
   - Product(s)：通常勾选 Process Simulate。  
   - File：选择或新建 XML（例如 TxTools.xml），点击 Register。  
6. 启动 Process Simulate → Customize → 将新命令拖到工具栏。  
卸载：再次运行 CommandReg.exe，选择相同 XML，点击 Unregister。

---

## 四、主要插件与功能简介
- ExportGun（导插枪）  
  将焊枪与焊点数据从 PS 导出到 CATIA（CGR 放置）或 Excel，便于下游工艺校验与可视化。包含 PsReader 风格的数据提取与 CatiaBridge（CATIA COM 互操作）。

- DotBall（点球）  
  在焊点处生成球形几何，便于在 CATIA 等系统中直观显示焊点位置。

- RobotReachabilityChecker（可达性验证）  
  基于机器人模型、焊接操作点与关节限位，分析可达性、是否超限、是否满足 TCP 余量等。

- LineToSolid（曲线转实体）  
  将场景中的 Polyline / Line / Arc 拆分成段，并按用户指定截面（矩形/圆形）为每段生成独立 Solid（每段为一个零件）。支持圆弧自适应细分以满足最大弦高设置。

- FenceBuilder / AutoFance（围栏生成）  
  根据直线或折线基线按参数生成网片 + 立柱 + 可选底板。优先采用“单薄板 + 纹理”策略以显著降低 Solid 数量并保证可视效果，提供降级路径以适配不同 PS 版本的纹理 API。

仓库还包含：AutoRecorder、DeviceZAligner、WeldAnnotator、WeldSpotAllocator、SelectButton 等模块（详见各子目录）。

---

## 五、典型使用流程（示例）
LineToSolid：
1. 在 PS 中选中一个或多个曲线特征。  
2. 菜单 → TxTools → LineToSolid，点【从选择添加】加入特征列表。  
3. 选择截面类型（矩形/圆形）、输入尺寸（mm），如有圆弧调整“最大弦高”。  
4. 点击【生成几何体】，操作包裹在 UndoScope 中，支持 Ctrl+Z 撤销。

FenceBuilder（围栏）：
1. 在场景准备直线或多段线基线（首端 Z 作为地面）。  
2. 菜单 → TxTools → 围栏生成器，选择基线并调整参数（网片宽/高、立柱尺寸、间隙、底板、纹理）。  
3. 点击【生成围栏】；若不满意使用 Ctrl+Z 或【撤销上次】。

ExportGun：
1. 在插件窗体加载或选择焊点集合/焊枪对象。  
2. 选择导出目标（Excel / CATIA），配置输出选项并执行导出。

---

## 六、实现细节与重要约定
- 圆弧离散（LineToSolid）：给定半径 r 与最大弦高 s，分段角度 θ = 2·arccos(1 - s/r)，总扫掠角 / θ 向上取整得到段数 n，然后等分生成点与直线段。  
- 姿态对齐：段方向 dir = normalize(End - Start) 作为局部 Z；选取不共线的辅助轴以生成局部 X、Y，组装 4×4 变换赋给几何的 AbsoluteLocation。  
- 围栏几何约定（FenceBuilder）：强制水平投影、立柱两端放置与段内定间距生成、角点合并规则、网片宽度按中心距减去立柱与间隙计算等。  
- 纹理/颜色设置采用分级回退：若纹理 API 不可用则退为半透明纯色；颜色构造尝试多种构造形式以兼容不同 PS 版本接口。

---

## 七、常见问题与排查建议
1. 找不到或无法访问 PS API 成员：检查当前引用的 PS SDK 版本，使用 IntelliSense 确认实际类型/成员名；首次运行观察日志以确定实际生效的探测路径。  
2. 插件未在 PS 中出现：确认 CommandReg 注册成功、选择的 XML 与产品是否正确、PS 是否重启。  
3. TxTransformation 或几何姿态异常：尝试不同的矩阵构造路径（ctor、SetMatrix、Matrix 属性等），并观察日志中回退路径。  
4. 纹理无法加载或颜色异常：检查资源是否随 DLL 部署或嵌入资源是否正确解包，纹理 API 有多级降级逻辑。  
5. 圆弧顶点/属性读取失败：模块尝试多种候选属性名（Center、Radius、Start/End、Normal/Axis、SweepAngle 等），请以日志为准固化正确字段。  
调试建议：在小场景下快速验证几何与姿态，再扩展到大场景；启用并审阅内置日志以判断探测路径与 API 调用结果。

---

## 八、打包与发布建议
- 资源（如 mesh_pattern.png）可随 DLL 同目录下的 Resources 文件夹一起发布，或嵌入并在运行时正确解包。  
- 插件发布前在对应 PS 版本上做回归测试（测量 Solid 数量与生成耗时），以避免在生产场景中出现性能问题。  
- 对于 CATIA 交互功能，确保目标机器上已安装所需的 CATIA COM 组件并做好权限/COM 注册验证。

---

## 九、开发与贡献要点
- 仅使用 C# 7.3 语法，避免新的语言特性引发运行时不兼容。  
- 所有不确定的外部 API 都应优先用 dynamic + try/catch 多路径探测，稳定后可替换为强类型以提升性能。  
- 每个 `TxButtonCommand` 的 GUID 顶部注解请确保为合法 GUID（使用 VS Create GUID 工具生成）。  
- 贡献前请阅读并遵守仓库根目录的 CONTRIBUTING.md。欢迎 Issues 与 PR。

---

## 十、许可与联系方式
- 授权：MIT（详见 LICENSE 文件）。  
- 欢迎在仓库中提交 Issue 与 PR： https://github.com/ShuiYinMaster/TxTools

---

## 十一、附录：仓库中应关注的模块与文件（供开发者快速定位）
- ExportGun/：导出相关实现（PsReader、ExportGunForm、CatiaBridge 等）。  
- LineToSolid/：曲线离散、几何构建、姿态对齐的实现。  
- AutoFance/（FenceBuilder）：围栏布局、几何构建、纹理处理。  
- RobotReachabilityChecker/：可达性检查逻辑。  
- 共享工具与资源：Common/、Image/、SRC/、Resources 等。

---

（文档由仓库现有 README 与子模块说明整理而成）
