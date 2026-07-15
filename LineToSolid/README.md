# LineToSolid — Process Simulate 2402 插件

以场景中已有的曲线特征（**TxPolylineFeature / TxLineFeature / TxArcFeature**）为基线，按用户指定的截面参数（矩形宽×高 / 圆形直径）为每个直线段生成一个独立的长方体或圆柱体。几何体的长度方向自动与所在段方向对齐。

## 功能特性

- **多特征并列独立处理**：可一次添加多个曲线特征（包括独立的圈圈、线段、折线、圆弧），每个特征按自己的内部顺序产生段，**不做跨特征拼接**
- 自动识别 3 种特征类型：
  - **Polyline**：按节点拆为多段
  - **Line**：1 段
  - **Arc**：按"最大弦高"自适应细分为多段直线段（弦高越小，越接近真圆弧）
- 截面：**矩形（宽×高）** 或 **圆形（直径）**
- **长度方向 = 段方向**（无需用户选择，几何体局部 Z 轴对齐段方向）
- 每段一个独立零件，中心 = 段中点
- 整批生成包裹在 UndoScope 中，支持 Ctrl+Z 撤销

## 文件结构

```
LineToSolid/
├── LineToSolidCommand.cs   入口命令（TxButtonCommand）
├── LineToSolidForm.cs      主窗体（TxForm + TxToolStrip + 卡片 + 日志）
├── PolylineReader.cs       多特征识别 + 段提取 + 圆弧离散
├── GeometryBuilder.cs      几何体创建与姿态对齐
└── README.md
```

## 使用流程

1. 在 PS 中选中一个或多个曲线特征（Polyline / Line / Arc 都行）
2. 菜单 → TxTools → LineToSolid，打开窗体
3. 工具条点 **[从选择添加]**，特征会进入列表
4. 可以反复选中其他特征 → **[从选择添加]** 累加；点 **[清空列表]** 清空
5. 在"截面参数"卡片中：
   - 选择截面类型（矩形 / 圆形）
   - 输入尺寸（mm）
   - 如果列表里有圆弧，调整"圆弧最大弦高（mm）"控制细分密度（默认 0.5）
6. 点 **[生成几何体]**
7. 不满意可 **[撤销上次]** 或 Ctrl+Z

## 圆弧细分原理

设圆弧半径为 r，最大允许弦高为 s（用户输入），则每段所张的圆心角为：

```
θ = 2 · arccos(1 - s/r)
```

总扫掠角 / θ 向上取整得到段数 n，然后在圆弧平面内等分插值生成 n+1 个点、n 个直线段。

平面基底构造：
- 法线 normal 优先用特征自带的 Normal/Axis；缺失时用 `cross(start-center, end-center)` 推断
- 平面内基底：`e1 = normalize(start - center)`，`e2 = normalize(cross(normal, e1))`
- 圆弧上点：`p(t) = center + r·cos(sweep·t)·e1 + r·sin(sweep·t)·e2`，t ∈ [0, 1]

## 姿态对齐原理

对每个段 `seg`：
- 段方向 `dir = normalize(End - Start)` → 作为局部 Z 轴
- 取辅助轴 `aux`：当 `|dir · worldZ| > 0.99` 时用 `worldX`，否则用 `worldZ`
- 局部 X = `normalize(aux × dir)`，局部 Y = `normalize(dir × localX)`
- 平移 = 段中点
- 组装 4×4 齐次矩阵 → 构造 `TxTransformation` → 赋给几何体 `AbsoluteLocation`

矩形/圆形截面对绕长度轴的旋转不敏感（圆完全对称，矩形 180° 周期），所以 X/Y 的具体朝向不需要参数化。

## 编译配置

- 目标框架：与 PS 2402 SDK 匹配（通常 .NET Framework 4.7.2 / 4.8）
- 引用程序集：
  - `Tecnomatix.Engineering.dll`
  - `Tecnomatix.Engineering.Ui.dll`
  - `System.Windows.Forms`
- 输出：DLL，放入 PS 插件目录后由 PS 加载

## 编译时必看的几个 TODO

代码遵循"未验证就不写死"的既有约定，所有不确定 API 都用反射 + try/catch 多路径回退：

### 1. 多段线顶点属性名（PolylineReader）

按 `Vertices → Points → ControlPoints → GetVertices() → GetPoints()` 顺序探测。首次运行时看日志判断哪条路径生效，再固化。

### 2. 圆弧特征属性名（PolylineReader）

按以下候选名探测：
- 圆心：`Center`
- 半径：`Radius`
- 起止：`StartPoint`/`Start`、`EndPoint`/`End`
- 法线：`Normal`/`Axis`
- 扫掠角：`TotalAngle`/`Angle`/`SweepAngle`

PS 2402 的 `TxArcFeature` 实际属性需要在 IntelliSense 中确认；首次运行看日志即可锁定。

### 3. 几何体创建 API（GeometryBuilder）

两条主路径回退：
- `TxApplication.ActiveDocument.PolyhedronOperations.CreateBox / CreateCylinder`
- `TxApplication.ActiveDocument.PhysicalOperations.CreateBox / CreateCylinder`

配合数据类：
- `TxBoxCreationData`（属性 Name / SizeX / SizeY / SizeZ）
- `TxCylinderCreationData`（属性 Name / Radius / Height）

**强烈建议**：编译第一次失败时根据 IntelliSense 锁定 PS 2402 的实际类型名与方法签名，把 `GeometryBuilder` 中 dynamic 路径替换成强类型调用，以提升性能与可读性。

### 4. UndoManager 接口

按 `OpenScope` / `BeginScope` / `Begin` 顺序尝试。

### 5. LineToSolidCommand 的 GUID

`LineToSolidCommand.cs` 顶部 `[Guid("...")]` 是占位字符串（且包含非法字符），**必须用 VS Tools → Create GUID 重新生成**替换。

## 与既有 TxTools 的一致性

- 卡片化 GroupBox + FlowLayout/TableLayout
- `SystemFonts.DefaultFont` 统一字号
- `FlatStyle.Flat` 按钮（`TxToolStrip` 默认就是 Flat）
- 顶部 `TxToolStrip` 工具条
- 底部可折叠日志面板
- UndoScope 包裹整批操作，Ctrl+Z 一键回滚

## 已知限制 / 后续可扩展

- 不支持跨特征拼接（按确认范围"独立处理"）
- 圆弧法线在父变换下未做严格旋转重映射；若特征本身已经在世界坐标系下表达则无影响
- 截面恒定，不支持沿段变截面
- 圆弧只支持平面圆弧；如果是螺旋线一类的特征，需要扩展 PolylineReader
- 后续可加"颜色选择"下拉给生成的零件着色
- 后续可加"段过短跳过阈值"暴露到 UI（当前固定 1e-6）
