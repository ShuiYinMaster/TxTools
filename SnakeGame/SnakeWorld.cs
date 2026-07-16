using System;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.DataTypes;

namespace TxTools.SnakeGame
{
    /// <summary>
    /// 贪吃蛇的 Process Simulate 世界层：
    ///   · 建模：参考 LineToSolid/GeometryBuilder — 自己创建 Resource + SetModelingScope，
    ///     不再依赖 CurrentModelingWorkingSpace（该属性多数情况下为 null）；
    ///   · 移动：用 ITxLocatableObject.LocationRelativeToWorkingFrame；
    ///   · 干涉集：参考 AutoPath/CollisionSetService — 强类型 API（TxCollisionRoot、
    ///     TxCollisionPairCreationData、TxCollisionPair），不再用反射拼凑。
    ///
    /// TxBoxCreationData 的 absLoc 语义（LineToSolid/FenceBuilder 已验证）：
    ///     absLoc 位置 = Z-minus 端面中心（局部底面）；box 沿 +Z 单侧延伸 sizeZ，X/Y 关于 absLoc 对称。
    ///     所以要让 box 几何中心位于 (worldX, worldY, 0)，absLoc.Translation = (worldX, worldY, -sizeZ/2)。
    /// </summary>
    public class SnakeWorld
    {
        private readonly double _cellSize;
        private readonly double _snakeSize;
        private readonly double _foodSize;
        private readonly Action<string> _log;

        public List<ITxObject> SnakeBoxes { get; private set; }
        public ITxObject FoodBox { get; private set; }

        // 干涉集 — 强类型 API（参考 AutoPath/CollisionSetService v5.0+）
        private TxCollisionPair _pair;
        private bool _prevCheckCollisions;
        private bool _hasPrevCheck;
        private bool _collisionSetup;

        // 建模资源 — 参考 LineToSolid/GeometryBuilder：自己创建 Resource 并打开建模作用域
        private TxComponent _modelingComponent;

        public SnakeWorld(double cellSize, Action<string> log)
        {
            _cellSize = cellSize;
            _snakeSize = cellSize * 0.9;      // 蛇身略小于格宽，视觉上有间隙
            _foodSize = cellSize * 0.6;       // 食物更小一点
            _log = log ?? new Action<string>(delegate { });
            SnakeBoxes = new List<ITxObject>();
        }

        // ========== 建模组件：自己创建 Resource（参考 LineToSolid/GeometryBuilder.BuildForSegments） ==========

        /// <summary>
        /// 确保有一个可用的建模组件。
        /// 不再依赖 doc.CurrentModelingWorkingSpace（该属性仅在 PS 已有活跃建模上下文时非 null）。
        /// 正确路径：PhysicalRoot.CreateResource → SetModelingScope → 强类型 CreateSolidBox。
        /// </summary>
        private TxComponent EnsureModelingComponent()
        {
            // 已有且仍然有效 → 复用
            if (_modelingComponent != null)
            {
                try { var _ = ((ITxObject)_modelingComponent).Name; return _modelingComponent; }
                catch { _modelingComponent = null; }
            }

            var doc = TxApplication.ActiveDocument;
            if (doc == null) { _log("ActiveDocument 为 null，无法建模。"); return null; }

            var root = doc.PhysicalRoot;
            if (root == null) { _log("PhysicalRoot 为 null，无法建模。"); return null; }

            try
            {
                string resName = "SnakeGame_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                ITxComponent comp = root.CreateResource(new TxResourceCreationData(resName));
                if (comp == null) { _log("CreateResource 返回 null，无法建模。"); return null; }

                if (!comp.CanOpenForModeling)
                {
                    _log("CanOpenForModeling=false，无法建模。");
                    return null;
                }

                comp.SetModelingScope();
                _modelingComponent = comp as TxComponent;
                _log("建模资源已创建：" + resName);
                return _modelingComponent;
            }
            catch (Exception ex)
            {
                _log("创建建模资源失败：" + ex.Message);
                return null;
            }
        }

        // ========== 蛇身几何 ==========

        public ITxObject CreateSnakeBox(Cell cell, string name)
        {
            var box = CreateBox(cell, _snakeSize, name);
            if (box != null) SnakeBoxes.Add(box);
            return box;
        }

        public void MoveSnakeBoxTo(int index, Cell cell)
        {
            if (index < 0 || index >= SnakeBoxes.Count) return;
            MoveBox(SnakeBoxes[index], cell, _snakeSize);
        }

        public void ClearSnake()
        {
            foreach (var b in SnakeBoxes) SafeDelete(b);
            SnakeBoxes.Clear();
        }

        // ========== 食物几何 ==========

        public void CreateFood(Cell cell)
        {
            SafeDelete(FoodBox);
            FoodBox = CreateBox(cell, _foodSize, "SnakeFood");
        }

        public void MoveFoodTo(Cell cell)
        {
            if (FoodBox == null) { CreateFood(cell); return; }
            MoveBox(FoodBox, cell, _foodSize);
        }

        public void RemoveFood()
        {
            SafeDelete(FoodBox);
            FoodBox = null;
        }

        public void ClearAll()
        {
            TeardownCollisionPair();
            ClearSnake();
            RemoveFood();

            // 清理建模资源，下次 CreateBox 时会自动重建
            try
            {
                if (_modelingComponent != null)
                {
                    ((ITxObject)_modelingComponent).Delete();
                }
            }
            catch { }
            _modelingComponent = null;
        }

        // ========== 通用几何操作 ==========

        /// <summary>
        /// 创建 Box：强类型 TxComponent.CreateSolidBox（参考 LineToSolid/GeometryBuilder.CreateBox）。
        /// </summary>
        private ITxObject CreateBox(Cell cell, double size, string name)
        {
            var comp = EnsureModelingComponent();
            if (comp == null) return null;

            try
            {
                var absLoc = BuildTransform(cell, size);
                var edgeSizes = new TxVector(size, size, size);
                var offset = new TxVector(0, 0, 0);
                var data = new TxBoxCreationData(
                    string.IsNullOrEmpty(name) ? "SnakeBox" : name,
                    absLoc,
                    edgeSizes,
                    offset);

                // 强类型调用，不再用反射（参考 LineToSolid 第 222 行）
                var solid = comp.CreateSolidBox(data);
                return solid as ITxObject;
            }
            catch (Exception ex)
            {
                _log("CreateBox 失败 (" + name + ")：" + ex.Message);
                return null;
            }
        }

        private void MoveBox(ITxObject box, Cell cell, double size)
        {
            if (box == null) return;
            try
            {
                var t = BuildTransform(cell, size);

                // 首选路径：ITxLocatableObject.LocationRelativeToWorkingFrame（可写）
                var loc = box as ITxLocatableObject;
                if (loc != null)
                {
                    try
                    {
                        loc.LocationRelativeToWorkingFrame = t;
                        return;
                    }
                    catch (Exception exLoc)
                    {
                        _log("LocationRelativeToWorkingFrame 写入失败，尝试 AbsoluteLocation：" + exLoc.Message);
                    }

                    // 备选：AbsoluteLocation
                    try
                    {
                        loc.AbsoluteLocation = t;
                        return;
                    }
                    catch { /* 继续退化 */ }
                }

                // 兜底：反射
                var prop = box.GetType().GetProperty("LocationRelativeToWorkingFrame");
                if (prop != null && prop.CanWrite) { prop.SetValue(box, t, null); return; }
                prop = box.GetType().GetProperty("AbsoluteLocation");
                if (prop != null && prop.CanWrite) { prop.SetValue(box, t, null); return; }

                _log("MoveBox：找不到可写的位置属性。");
            }
            catch (Exception ex)
            {
                _log("MoveBox 失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 构造 absLoc：位置 = (X·cell, Y·cell, -size/2)，zDir=+Z，方向对齐世界。
        /// 这样几何中心正好落在 (X·cell, Y·cell, 0)。
        /// </summary>
        private TxTransformation BuildTransform(Cell cell, double size)
        {
            var pos = new TxVector(
                cell.X * _cellSize,
                cell.Y * _cellSize,
                -size / 2.0);
            var zDir = new TxVector(0, 0, 1);
            return new TxTransformation(pos, zDir);
        }

        // ========== 干涉集 — 强类型 API（参考 AutoPath/CollisionSetService） ==========

        /// <summary>
        /// 建立 First=head、Second=food+body 的干涉对。
        /// 使用强类型 API：TxCollisionPairCreationData + CollisionRoot.CreateCollisionPair。
        /// </summary>
        public void SetupCollisionPair()
        {
            if (_collisionSetup) return;
            try
            {
                if (SnakeBoxes.Count == 0) return;

                var first = new TxObjectList();
                first.Add(SnakeBoxes[0]);

                var second = new TxObjectList();
                if (FoodBox != null) second.Add(FoodBox);
                for (int i = 1; i < SnakeBoxes.Count; i++) second.Add(SnakeBoxes[i]);

                TxCollisionRoot root = TxApplication.ActiveDocument.CollisionRoot;
                if (root == null)
                {
                    _log("未找到 CollisionRoot，跳过干涉集（不影响游戏运行）。");
                    return;
                }

                var data = new TxCollisionPairCreationData();
                data.FirstList = first;
                data.SecondList = second;

                _pair = root.CreateCollisionPair(data);
                if (_pair == null)
                {
                    _log("CreateCollisionPair 返回 null。");
                    return;
                }

                // 命名（可选，某些 SDK 版本可能不支持）
                try { _pair.Name = "SnakeGamePair"; } catch { }

                // 激活碰撞对
                try { _pair.Active = true; } catch { }

                // 确保全局碰撞检查开关开启
                try
                {
                    _prevCheckCollisions = root.CheckCollisions;
                    _hasPrevCheck = true;
                    if (!root.CheckCollisions)
                    {
                        root.CheckCollisions = true;
                        _log("CollisionRoot.CheckCollisions 已开启。");
                    }
                }
                catch (Exception exc) { _log("设置 CheckCollisions 失败：" + exc.Message); }

                _collisionSetup = true;
                _log("干涉集已建立：SnakeGamePair。");
            }
            catch (Exception ex)
            {
                _log("SetupCollisionPair 失败：" + ex.Message);
            }
        }

        /// <summary>将新蛇身加入干涉对的 SecondList（强类型）。</summary>
        public void AddToCollisionSecond(ITxObject obj)
        {
            if (_pair == null || obj == null) return;
            try
            {
                _pair.SecondList.Add(obj);
            }
            catch (Exception ex)
            {
                _log("AddToCollisionSecond 失败：" + ex.Message);
            }
        }

        /// <summary>从干涉对的 SecondList 中移除对象（强类型）。</summary>
        public void RemoveFromCollisionSecond(ITxObject obj)
        {
            if (_pair == null || obj == null) return;
            try
            {
                _pair.SecondList.Remove(obj);
            }
            catch (Exception ex)
            {
                _log("RemoveFromCollisionSecond 失败：" + ex.Message);
            }
        }

        /// <summary>拆除干涉对，恢复全局碰撞检查开关。</summary>
        public void TeardownCollisionPair()
        {
            // 恢复 CheckCollisions
            try
            {
                if (_hasPrevCheck)
                {
                    TxCollisionRoot root = TxApplication.ActiveDocument.CollisionRoot;
                    if (root != null)
                    {
                        root.CheckCollisions = _prevCheckCollisions;
                    }
                }
            }
            catch { }

            // 删除碰撞对（强类型 Delete）
            try { if (_pair != null) _pair.Delete(); } catch { }
            _pair = null;
            _hasPrevCheck = false;
            _collisionSetup = false;
        }

        /// <summary>查询干涉对当前是否处于干涉状态（可选，供上层实时读取）。</summary>
        public bool? IsPairColliding()
        {
            if (_pair == null) return null;
            try
            {
                // 尝试多种可能的属性名
                var t = _pair.GetType();
                var pStatus = t.GetProperty("IsColliding")
                              ?? t.GetProperty("Colliding");
                if (pStatus != null)
                    return (bool)pStatus.GetValue(_pair, null);

                // 兜底：State 属性
                var pState = t.GetProperty("State");
                if (pState != null)
                {
                    var v = pState.GetValue(_pair, null);
                    if (v != null)
                        return v.ToString().IndexOf("Colli", StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }
            catch { }
            return null;
        }

        // ========== 辅助 ==========

        private static void SafeDelete(ITxObject obj)
        {
            if (obj == null) return;
            try
            {
                var doc = TxApplication.ActiveDocument;
                if (doc == null) return;

                // 尝试 doc.RemoveObject
                var m = doc.GetType().GetMethod("RemoveObject", new[] { typeof(ITxObject) });
                if (m != null) { m.Invoke(doc, new object[] { obj }); return; }

                // 尝试 doc.RemoveObjects
                m = doc.GetType().GetMethod("RemoveObjects", new[] { typeof(TxObjectList) });
                if (m != null)
                {
                    var list = new TxObjectList();
                    list.Add(obj);
                    m.Invoke(doc, new object[] { list });
                    return;
                }

                // 兜底：obj.Delete
                m = obj.GetType().GetMethod("Delete", Type.EmptyTypes);
                if (m != null) m.Invoke(obj, null);
            }
            catch { /* 静默：删除失败不影响游戏 */ }
        }
    }
}
