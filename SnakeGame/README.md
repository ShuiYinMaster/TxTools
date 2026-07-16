# TxTools.SnakeGame

在 Process Simulate 场景内玩一局贪吃蛇：蛇头/蛇身/食物都是真实的 3D 实体（`TxSolid` 长方体），移动、吃食物、干涉高亮全部映射到 PS SDK 调用。

## 文件结构

| 文件 | 职责 |
| --- | --- |
| `SnakeGameCommand.cs` | Ribbon 入口，`TxButtonCommand`，点击后非模态弹出游戏窗体 |
| `SnakeGameEngine.cs` | 纯逻辑：网格坐标、方向、食物生成、越界/自碰撞判定。**不依赖 PS SDK**，可脱离 PS 单独跑单元测试 |
| `SnakeWorld.cs` | 全部 PS SDK 调用：建模（`TxComponent.CreateSolidBox`）、移动（`LocationRelativeToWorkingFrame`）、干涉集（`TxCollisionRoot` / `TxCollisionPair`） |
| `SnakeGameForm.cs` | WinForms 窗体：Timer 驱动游戏循环、方向键/WASD 输入、日志区 |

## 玩法

1. 打开一个 PS study，点击 Ribbon 上的「贪吃蛇小游戏」。
2. 点「开始 / 重开」：场景原点生成蛇头，随机位置生成食物。
3. 方向键 / WASD 控制方向，空格暂停。
4. 吃到食物：长度 +1，速度按分数阶梯式提升。
5. 撞到自己或出界：游戏结束，弹窗显示最终得分。
6. 「清除几何」：手动清空所有蛇身/食物实体和干涉对。

## 参数（`SnakeGameForm.cs` 顶部常量）

| 常量 | 默认值 | 说明 |
| --- | --- | --- |
| `GridHalfExtent` | 10 | 网格半边长，21×21 网格 `[-10,10]` |
| `CellSize` | 60.0 mm | 每格边长，全场约 1.26m × 1.26m |
| `InitialTickMs` | 200 | 初始移动间隔 |
| `MinTickMs` | 60 | 最快移动间隔（速度上限） |
| `SpeedupEveryScore` | 5 | 每吃 N 个食物加速一次 |
| `SpeedupStepMs` | 20 | 每次加速减少的间隔 |

蛇身 box 边长 = `CellSize × 0.9`（留视觉间隙），食物 box 边长 = `CellSize × 0.6`。

## 关键设计

### 建模：自建 Resource + `SetModelingScope`

`SnakeWorld.EnsureModelingComponent()` 不依赖 `doc.CurrentModelingWorkingSpace`（该属性只有在 PS 已有活跃建模上下文时才非 null，游戏刚启动时通常是 null）。改为参考 LineToSolid/GeometryBuilder 的做法：

```
PhysicalRoot.CreateResource(new TxResourceCreationData(resName))
  → 检查 CanOpenForModeling
  → comp.SetModelingScope()
  → 后续所有 CreateSolidBox 复用这个 comp
```

整局游戏只建一次 Resource（缓存在 `_modelingComponent`），`ClearAll()` 时连同 Resource 一起删除，下次开局自动重建。

### `TxBoxCreationData` 的 absLoc 语义

LineToSolid/FenceBuilder 已验证：absLoc 位置 = 长方体 **Z-minus 端面中心**（局部底面），沿 +Z 单侧延伸 `sizeZ`，X/Y 关于 absLoc 对称。所以要让 box 几何中心落在世界坐标 `(cell.X×CellSize, cell.Y×CellSize, 0)`，需要：

```
absLoc.Translation = (cell.X×CellSize, cell.Y×CellSize, -size/2)
absLoc.zDir = (0,0,1)
```

### 移动：不重建 box，只改 Location

每 tick 引擎更新 `SnakeCells` 后，UI 层调用 `MoveSnakeBoxTo(i, cell)` → `SnakeWorld.MoveBox`：首选 `ITxLocatableObject.LocationRelativeToWorkingFrame`，失败退化到 `AbsoluteLocation`，再退化到反射兜底。吃到食物才 `CreateSolidBox` 一个新蛇身 append 到末尾；食物 box 全程复用（移动而非重建）。

### 干涉集：强类型 API

参考 AutoPathPlanner/CollisionSetService v5.0+ 的经验，干涉集这块**改用强类型调用，不再反射**：

```
TxCollisionRoot root = TxApplication.ActiveDocument.CollisionRoot;
var data = new TxCollisionPairCreationData { FirstList = first, SecondList = second };
TxCollisionPair pair = root.CreateCollisionPair(data);
pair.Active = true;
root.CheckCollisions = true;   // 游戏结束/清除时恢复原值
```

- `FirstList` = 蛇头
- `SecondList` = 食物 + 全部蛇身（吃到食物新增蛇身时 `pair.SecondList.Add(newBox)`）
- 干涉对仅用于让 PS 3D 视图里看到干涉高亮；主碰撞判定仍走 `SnakeGameEngine` 内部的网格坐标比较（可靠、无浮点误差）
- `IsPairColliding()` 预留了读取干涉状态的接口，如果想改成"以 PS 干涉信号触发吃食物"可以从这里接入

## 已知的版本相关风险点

以下几处如果您的 PS 2402 SDK 版本行为不同，编译或运行时会报错，可定点修改：

1. **`TxResourceCreationData` 构造签名** —— 目前用单参 `(string name)`，如果 SDK 要求更多参数会在 `EnsureModelingComponent` 报编译错。
2. **`TxCollisionRoot.CreateCollisionPair` 方法名/签名** —— 如果不存在，日志会打印「未找到 CollisionRoot」或「CreateCollisionPair 返回 null」，不影响游戏本身运行，只是看不到干涉高亮。
3. **`TxCollisionPair.Active` / `.Name` 属性** —— 已用 try/catch 包裹，缺失也不影响主流程。
4. **`ITxLocatableObject.LocationRelativeToWorkingFrame` 可写性** —— 已提供 `AbsoluteLocation` 和反射两级兜底。

调试时看窗体内的日志区（黑底文本框），所有失败路径都会打印具体异常信息，按 SHUIYIN 一贯的"跑起来看日志再改"节奏即可。
