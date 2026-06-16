# FenceBuilder — Process Simulate 2402 围栏生成插件

根据场景中已有的直线/多段线特征(`TxLineFeature` / `TxPolylineFeature`),按用户指定的网片与立柱参数,沿线段生成围栏(网片 + 立柱 + 可选底板)。

## 设计取舍

经过方案对比,采用**实时生成 + 单薄板网片**:

| 方案 | 优点 | 缺点 | 取舍 |
|---|---|---|---|
| 纯密集圆柱网格 | 视觉最真实 | 5000+ Solid,PS 卡顿 | 否决 |
| 预制 JT 库 | 引用即用 | 长度离散,不能贴合任意线段 | 否决 |
| **单薄板 + 网格纹理** | **1 solid/网片,适配任意长度** | **依赖 PS 纹理 API,有降级风险** | **采用** |

实际几何成本:
- 单立柱:1 Solid(无底板) / 2 Solid(有底板)
- 单网片:4 Solid(外框上下左右) + 1 Solid(薄板)= 5 Solid
- 100m 围栏,2m 网片宽,50 片:50×5 + 51×1 ≈ **300 Solid**,完全可控

## 几何约定

1. **强制水平**:每条基线取首端 Z 作为地面高,所有点投影到该 Z 平面,XY 走线。线段两端 Z 不一致时直接忽略 Z 变化。
2. **坐标系**:每个长方体的局部 X = 沿线段方向,Y = 线段法线方向(XY 平面内左旋 90°),Z = 世界 Z(竖直向上)。
3. **立柱位置**:
   - 段两端各 1 根,中心从线段端点沿线段方向缩进 `PostWidth/2`,使立柱外沿贴齐端点。
   - 段内每 `MeshNominalWidth + PostWidth + 2×PostGap` 一根。
   - 拐角共享:相邻段端点立柱距离 < 0.5×PostWidth 时合并,方向取平分角。
4. **网片宽度**:`实际宽度 = 相邻立柱中心距 - PostWidth - 2×PostGap`。末片在线段无法整除时自动截断(末片宽度变小)。
5. **立柱/网片高度**:
   - 网片下沿 Z = `GroundZ + GroundClearance`
   - 网片上沿 Z = `GroundZ + GroundClearance + MeshHeight`
   - 立柱顶 = 网片上沿;立柱底 = 地面(无底板) / 地面 + 底板厚度(有底板)

## 文件结构

```
FenceBuilder/
├── FenceBuilderCommand.cs      入口命令(ITxButtonCommand)
├── FenceBuilderForm.cs         主窗体(TxForm + GroupBox 卡片 + 日志)
├── FenceBaselineReader.cs      基线读取(顶点反射探测 + 水平投影)
├── FenceLayoutPlanner.cs       布局算法(纯几何,无 SDK 依赖)
├── FenceGeometryBuilder.cs     几何创建(CreateSolidBox)
├── FenceParameters.cs          参数容器
├── AppearanceHelper.cs         颜色/透明度/纹理(三级降级反射)
├── Resources/
│   └── mesh_pattern.png        16×16 方格透明 PNG(随 DLL 打包)
└── README.md
```

## 编译时需要核对的几个 TODO

代码遵循"未验证就不写死"的既有约定。运行时观察日志,根据实际成功路径再固化。

### 1. `TxObjGridCtrl` 的对象读取 API(FenceBuilderForm.ReadGridObjects)

按 `Objects → Items → Entries → GetObjects() → GetItems() → GetAll()` 顺序探测。如果都失败,需要根据 PS 2402 SDK 实际接口名调整。

### 2. `TxTransformation` 的 4×4 矩阵构造(FenceGeometryBuilder.BuildXYAxisAlignedTransform)

按 `ctor(double[,])` → `SetMatrix(double[,])` → `Matrix` 属性 → 退化为 `(pos, zDir)` 默认朝向。**退化路径下立柱/网片不会按线段方向旋转**,首次运行务必观察日志,如果出现"姿态构造退化为默认朝向"警告,说明矩阵接口需要修复。

### 3. `TxAppearance` 纹理 API(AppearanceHelper.TrySetTexture)

按 `SetTexture(string)` → `ApplyTexture(string)` → `LoadTexture(string)` → `TexturePath` / `Texture` / `TextureFile` 属性 → `Appearance.<property>` 子对象路径探测。失败自动降级为半透明纯色。**如果你的项目对纹理无硬性要求,可以直接关闭"贴网格纹理"选项,所有降级路径均稳定**。

### 4. 颜色 API(AppearanceHelper.TrySetColor)

按 `solid.Color = TxColor` → `solid.Appearance.Color = TxColor`。`TxColor` 优先用 `(byte, byte, byte)` 构造,失败时回退到 `(int, int, int)`,再失败传 `System.Drawing.Color`。

### 5. `Polyline` 顶点读取

沿用 LineToSolid 的成熟探测路径:`Vertices → Points → ControlPoints → GetVertices() → GetPoints() → GetControlPoints()`。

## 使用流程

1. 在 PS 场景里画好基线特征(直线或多段线),特征本身的 Z 高度就是地面高
2. 菜单 → TxTools → 围栏生成器
3. 选中基线特征,点【从选择添加】
4. 调整参数:
   - 网片宽度、高度、外框方管尺寸、离地间隙
   - 立柱截面、立柱-网片间隙
   - 是否启用底板
   - 是否贴网格纹理(失败自动降级为半透明纯色)
5. 【生成围栏】
6. 不满意按 Ctrl+Z(整批撤销)或点【撤销上次】

## 性能与扩展

- 当前实现没有用任何复杂特征(无布尔运算、无圆角、无纹理坐标计算),`CreateSolidBox` 是 PS 最廉价的实体创建方式。100m 围栏约 300 Solid,生成时间预计 1-3 秒。
- 如果未来要支持更细的网格(双层菱形网、夹丝结构),可以在 `FenceGeometryBuilder.Build` 里给薄板那一层增加循环创建小圆柱的路径,作为"高保真模式"开关。
- 如果要做"标准件库"扩展,可以新增一个 `IFenceMeshSource` 接口,默认实现是当前的"实时 CreateSolidBox",另一个实现是"从 cojt 库 Insert",由 UI 切换。

## 编译配置

- 目标框架:与 PS 2402 SDK 匹配(通常 .NET Framework 4.7.2 / 4.8)
- 引用程序集:
  - `Tecnomatix.Engineering.dll`
  - `Tecnomatix.Engineering.Ui.dll`
  - `System.Windows.Forms`
  - `System.Drawing`
- 输出:DLL,放入 PS 插件目录;`Resources/mesh_pattern.png` 一起打包到 DLL 同级 `Resources` 子目录
