# Contributing to TxTools · 贡献指南

[English](#english) · [中文](#中文)

---

## English

Thanks for your interest in contributing to TxTools. This document lists the conventions the project follows so that your pull requests integrate smoothly.

### Before You Start

- Open an issue to discuss non-trivial changes before coding — this avoids wasted effort on features that may not fit the project direction.
- For bug reports, include: Process Simulate version, a minimal reproduction, and the relevant log output.

### Development Environment

- **Visual Studio** 2019 or 2022
- **.NET Framework** 4.8 target
- **Process Simulate** 2402 recommended for testing
- Reference `Tecnomatix.Engineering.dll` from your PS installation — do **not** commit this DLL to the repository.

### Coding Rules

These rules are strict because the Process Simulate hosting environment imposes hard constraints:

#### 1. C# 7.3 only

The PS Script Editor and plugin runtime compile against C# 7.3. Do **not** use any newer syntax:

- ❌ `using var`  → ✅ `using (var x = ...)`
- ❌ Default interface implementations
- ❌ Target-typed `new` (e.g. `List<int> x = new();`)
- ❌ Range operators (`^`, `..`)
- ❌ Switch expressions
- ❌ `record` types
- ❌ Pattern matching enhancements from C# 8+
- ✅ `is Type var` patterns (C# 7.0 — OK)
- ✅ Default parameter values, tuple types, local functions (all C# 7.x — OK)

#### 2. Defensive SDK access

PS SDK member names may differ between versions. When calling anything that is not universally stable:

```csharp
try { dynamic d = obj; result = d.SomeProperty; } catch { }
if (result == null) try { ... } catch { }  // fallback
```

Never crash the entire plugin because one property name moved.

#### 3. Never guess API members

If you are unsure whether `TxXxx.SomeProperty` exists, **check** — either via official documentation, a runtime probe, or by asking in an issue. Silent wrong behavior is far worse than an honest compile error.

#### 4. Full rewrites over messy patches

When a file has structural issues, prefer rewriting it cleanly rather than stacking conditional patches. Readability wins over minimal-diff heroics.

#### 5. Logging style

- Use prefixed tags: `[APP]` for appearance binding, `[坐标]` for coordinate frame, `[PS]` for PS-layer parsing, `[调试]` for debug, etc.
- Keep logs concise by default; verbose diagnostics should be opt-in via a parameter or a documented debug switch.
- Never log inside tight loops without throttling.

### Pull Request Checklist

- [ ] Code compiles with no warnings against .NET Framework 4.8 and C# 7.3
- [ ] Tested on Process Simulate 2402
- [ ] No `Tecnomatix.Engineering.dll` or other PS install files committed
- [ ] No generated geometry files (`.cgr`, `.CATPart`, etc.) committed
- [ ] No bin/obj directories committed
- [ ] Commit messages describe intent, not just the diff
- [ ] Any new public API has a doc comment

### Commit Messages

Keep commit messages in English or Chinese — either is fine. A clear subject line helps:

```
[ExportGun] Fix CGR path resolution for jointly-stored tools
[Reachability] Add J7/J8 external axis exclusion
[Common] Refactor PsReader.FillPoints to L1/L2/L3/L4 layers
```

### License

By contributing, you agree that your contributions will be licensed under the MIT License, same as the rest of the project.

---

## 中文

感谢你对 TxTools 项目的关注。本文档列出了项目遵循的规范，帮助你的 PR 顺利合入。

### 开始之前

- 对于非琐碎的改动，**先开 Issue 讨论**再动手写代码——避免你花了精力做了不符合项目方向的功能
- 提 Bug 报告时请附：Process Simulate 版本、最小复现步骤、相关日志输出

### 开发环境

- **Visual Studio** 2019 或 2022
- **.NET Framework** 4.8 目标
- 推荐用 **Process Simulate 2402** 测试
- 从 PS 安装目录引用 `Tecnomatix.Engineering.dll`——**不要**把这个 DLL 提交到仓库

### 编码规则

以下规则都是硬性的，源于 Process Simulate 宿主环境的约束：

#### 1. 仅允许 C# 7.3 语法

PS Script Editor 和插件运行时按 C# 7.3 编译。**禁用**任何更新的语法：

- ❌ `using var`  → ✅ `using (var x = ...)`
- ❌ 默认接口实现
- ❌ 目标类型 `new`（如 `List<int> x = new();`）
- ❌ 范围运算符（`^`、`..`）
- ❌ Switch 表达式
- ❌ `record` 类型
- ❌ C# 8+ 的模式匹配增强
- ✅ `is Type var` 模式（C# 7.0 即支持，可用）
- ✅ 默认参数值、元组类型、本地函数（C# 7.x 均支持，可用）

#### 2. 防御式 SDK 调用

PS SDK 成员名在不同版本间可能不同。对不是绝对稳定的调用：

```csharp
try { dynamic d = obj; result = d.SomeProperty; } catch { }
if (result == null) try { ... } catch { }  // 兜底
```

不要因为某个属性名改位置就让整个插件崩溃。

#### 3. 绝不猜测 API 成员名

对 `TxXxx.SomeProperty` 是否存在不确定时，**去查**——通过官方文档、运行时探测或提 Issue 询问。静默的错误行为远比诚实的编译错误糟糕。

#### 4. 完整重写优于乱糟糟的补丁

文件出现结构性问题时，优先完整重写，而不是堆叠条件补丁。可读性比"最小 diff"重要。

#### 5. 日志风格

- 使用前缀标签：`[APP]` 外观绑定，`[坐标]` 参考坐标系，`[PS]` PS 层解析，`[调试]` 调试等
- 默认保持日志简洁；详细诊断应通过参数或文档化的调试开关显式启用
- 紧循环内的日志必须加节流，不要无限制输出

### PR 自检清单

- [ ] 代码在 .NET Framework 4.8 + C# 7.3 下无警告编译通过
- [ ] 已在 Process Simulate 2402 上测试
- [ ] 未提交 `Tecnomatix.Engineering.dll` 等 PS 安装文件
- [ ] 未提交生成的几何文件（`.cgr`、`.CATPart` 等）
- [ ] 未提交 bin/obj 目录
- [ ] Commit 信息描述的是**意图**，不只是 diff
- [ ] 新增的公开 API 有文档注释

### Commit 信息

中英文都可以。清晰的标题行很重要：

```
[ExportGun] 修复工具 CGR 路径解析（工具共存于同一目录时）
[Reachability] 新增 J7/J8 外部轴排除
[Common] 重构 PsReader.FillPoints 为 L1/L2/L3/L4 分层
```

### 许可

提交即视为同意你的贡献按 MIT 协议发布，与项目其余部分一致。
