# MechArena  |  机器人竞技场（PDPS 插件）

一个基于 Tecnomatix Process Simulate 2402 的射击小游戏。玩家在场景原点操控一个绿色方块，
对抗场内所有 TxRobot —— 机器人通过关节 Sin 波挥舞进行攻击，>=3 台时自动合体成 Boss。

## 文件结构

```
TxTools.MechArena/
├── MechArenaCommand.cs      # Ribbon 命令入口（TxButtonCommand）
├── MechArenaForm.cs         # 主窗体 + 键盘输入 + 状态显示
├── MechArenaEngine.cs       # 游戏引擎 + 玩家 + 子弹 + 敌方机器人 + Boss 阵型
└── MechArenaHelpers.cs      # 几何工厂 + 机器人关节赋值三重防御
```

## 操作

| 按键 | 功能 |
|---|---|
| W / A / S / D | XY 平面移动 |
| Q / E | 上升 / 下降 |
| Space | 自动锁定最近敌人开火 |
| Esc | 停止游戏 |

## 潜在需要调整的点

1. **`TxColor(r, g, b)` 构造签名**  
   若报错，可能是 `TxColor(byte, byte, byte)` 或需要 alpha 参数，看编译错自行调。

2. **`TxJoint.CurrentValue` 赋值**  
   `MechArenaHelpers.MechArenaRobotHelper.SetJointValues` 有三重降级，
   跑一遍看哪条路生效 —— 直接把 `Console.WriteLine` 或
   `TxApplication.SystemRootContext.WriteLine` 打到日志确认。

3. **`Robot.TCPF`**  
   如果场景里机器人未装工具，TCPF 可能为 null，`GetAttackTip()` 已回退到基座。

4. **`comp.EndModelingScope()`**  
   SnakeGame 经验：可省略。反射调不到就跳过。

5. **速度/幅度/半径**  
   `MechArenaEngine.cs` 顶部常量区一键调优。

## 已知限制

- 每颗子弹独立 `TxComponent`，产量大时会影响性能。若卡顿，可加子弹池
  （预建 N 个组件循环复用，仅改位置）。
- 输入捕获必须窗体获焦，PS 3D 视图无法直接接收 WASD。
- Boss 阵型只做了"血量翻倍 + 相位错开 120°"，没做本体位置调整。
  想更 dramatic 可以在 `BossFormation.TryActivate` 里改机器人 `AbsoluteLocation`
  排成三角形围绕原点（注意 Dispose 时恢复）。
