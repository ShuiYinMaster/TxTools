# WeldSpotGrouper — 焊点自动分组插件

按焊点绑定的零件，把绑定信息完全一致（零件名 + 数量都相同）的焊点归到同一个**新建的空白焊接操作**里。

## 文件结构

```
WeldSpotGrouper/
├─ Core/GroupModels.cs       数据模型（SpotItem / SpotGroup / GroupOptions / GroupReport）
├─ Ps/SpotGrouper.cs         读取层：枚举焊点 → 读绑定零件 → 多重集指纹分组
├─ Ps/GroupWriter.cs         写入层：直接建空白焊接操作 + MoveLocationInto 移动焊点（Undo 包裹）
├─ Ui/GrouperForm.cs         GUI 窗体（TxForm + TxObjGridCtrl 原生拾取）
└─ WeldSpotGrouperCmd.cs     入口命令（打开窗体）
```

命名空间统一 `TxTools.WeldSpotGrouper`。目标 C# 7.3 / .NET 4.8 / PS 2402。

## 接入

1. 把 5 个 .cs 加进工程（旧式 .csproj 需手动 include 每个文件）。
2. 引用 `Tecnomatix.Engineering`、`Tecnomatix.Engineering.Ui`、`System.Windows.Forms`、`System.Drawing`。
3. 按 DotNetCommand 注册 `WeldSpotGrouperCmd`，挂到 TxTools 的「机器人/焊接」组。

## 用法

1. 在 PS 树里**拾取范围节点**（焊接/复合操作，或资源），窗体里点绿色拾取行可加多个。
2. 设选项：命名前缀、零件名是否忽略大小写、是否跳过无绑定的点。
3. **扫描预览** → 表里看各组的焊点数与零件指纹。
4. **执行分组** → 每组新建一个空白焊接操作（挂 OperationRoot 根），把该组焊点移入。

整个执行包在一个 Undo 事务里，出错回滚；PS 里也可 Ctrl+Z。

## 首跑要看的两处日志（`%TEMP%\WeldSpotGrouper.log`，窗体点「日志」也能看）

1. **`[建] ...`** — 建焊接操作走的哪条。SDK 命名规律：`ITxWeldLocationOperationCreation`
   建的是「焊点位(子)」；本插件要的是同族 `ITxWeldOperationCreation.CreateWeldOperation(TxWeldOperationCreationData)`
   建「焊接操作(容器)」。反射会自动定位；若出现「未匹配到创建方法」，日志会列出
   OperationRoot 上全部候选方法签名，贴回来即可锁死。
2. **`[移动] ...`** — `MoveLocationInto` 命中的分支（WeldSpotAllocator 已验证，大概率直接过）。

另需留意：`AssignedParts` 返回的是 `TxPart` 还是 `TxPartAppearance`（影响指纹用的名字）——
看预览表里每组的零件标签是否符合预期即可判断。
