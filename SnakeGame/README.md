# TxTools · SnakeGame（贪吃蛇）

在 Process Simulate 3D 视图里跑一局贪吃蛇。所有蛇身、食物都是用 `ITxGeometryCreation.CreateSolidBox` 生成的实体（沿用 LineToSolid 的 absLoc = Z-minus 端面中心、单侧 +Z 延伸的语义）。

## 文件结构

| 文件 | 作用 |
| --- | --- |
| `SnakeGameCommand.cs` | Ribbon 命令入口（`TxButtonCommand`），打开窗体 |
| `SnakeGameEngine.cs` | 纯逻辑：网格坐标 / 方向 / 食物刷新 / 越界 / 自碰撞（无 PS 依赖） |
| `SnakeWorld.cs` | PS 几何：`CreateSolidBox` 创建蛇身与食物，`LocationRelativeToWorkingFrame` 移动；`TxCollisionRoot.CreatePair` 维护干涉集 |
| `SnakeGameForm.cs` | WinForms 窗体 + Timer 游戏循环 + 键盘输入 |

## 玩法

1. 在 PS 里打开一个 study，运行 `SnakeGame` 命令。
2. 点 **开始 / 重开**：场景原点生成一个正方体（蛇头），在 `[-10,+10]×[-10,+10]` 网格内随机位置生成一个较小的食物正方体。
3. 用 **方向键** 或 **WASD** 控制方向，**空格** 暂停/继续。
4. 蛇头吃到食物：食物移到新随机位置，尾部追加一个新蛇身正方体，并加入到干涉集的 `SecondList`。
5. 蛇头越界或撞到自己 → 游戏结束。

窗体标题栏下方有得分、长度、状态、速度实时显示；底部黑色区是运行日志（含每次吃食物、每次提速、以及所有 PS SDK 调用的失败原因）。

## 关键参数（`SnakeGameForm.cs` 顶部常量）

```csharp
GridHalfExtent   = 10      // 21×21 网格
CellSize         = 60.0    // 每格 60mm
InitialTickMs    = 400     // 起始每步 400ms
MinTickMs        = 120     // 最低 120ms
SpeedupEveryScore= 5       // 每 5 分加速一次
SpeedupStepMs    = 30      // 每次加速缩短 30ms
```

## 干涉集（TxCollisionPair）设计

- `FirstList`  = `[蛇头]`
- `SecondList` = `[食物, 蛇身1, 蛇身2, ...]`
- 每次吃到食物，新增蛇身 box 后立刻 `SecondList.Add(newBody)`。
- 启动时把 `TxCollisionRoot.CheckCollisions` 置 `true`，游戏结束/清理时恢复原值。

主碰撞判定仍然走网格坐标（可靠、无浮点误差），干涉集主要用于让 PS 界面里能直观看到干涉高亮。`SnakeWorld.IsPairColliding()` 也保留了读取 PS 干涉状态的能力，若您想改成"以 PS 干涉状态触发吃食物"可以直接接入。

## 与 LineToSolid 一致的几何模式

- `SetModelingScope()` → `CleanupPrevious()` → `CreateSolidBox(data)`；**不调用 `EndModeling`**（记忆：会 freeze，PS 会自动 finalize）。
- `TxBoxCreationData.AbsoluteLocation.Translation = (worldX, worldY, -size/2)`，使得 box 几何中心恰好落在 `(worldX, worldY, 0)`。
- 移动 box 用 `LocationRelativeToWorkingFrame`，退化到反射 `AbsoluteLocation`。

## 可能需要根据实际 PS 版本调整的点

代码对以下位置做了 `try/catch` + 反射兜底，若编译或运行时报错，看日志区提示对应改：

1. **命令基类** — 我用了 `TxButtonCommand`；若您的 TxTools 用别的（如 `TxRegistryCommand`），换基类即可，`Execute` 方法签名保持一致。
2. **`TxBoxCreationData` 属性名** — 尝试 `AbsoluteLocation`、`AbsLoc` 两个名字。
3. **`TxCollisionPairCreationData`** — 反射按简单名在 `Tecnomatix.Engineering` 装配里查找，找不到就跳过干涉集（不影响游戏本身）。
4. **`TxCollisionRoot.CreatePair`** — 依次尝试 `CreatePair` / `CreateCollisionPair` / `CreateChild`，最后扫描单参 `Create*` 方法兜底。
5. **删除对象** — 依次尝试 `doc.RemoveObject(obj)` / `doc.RemoveObjects(list)` / `obj.Delete()`。

日志区会打印每一次失败点，SHUIYIN 的常规迭代路径：先跑一次 → 看日志 → 定点改。
