// OpGridHost.cs  —  C# 7.3
// 强类型封装 TxObjGridCtrl。ListenToPick 常开：网格内自带拾取行，点它再去 PS 选对象即加入。
// 仅焊接操作硬限制：ObjectInserted 后立即剔除非 keep 对象。变更（增/删行）回调 onChanged 刷新计数。

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace MyPlugin.WeldSpotAllocator
{
    public sealed class OpGridHost
    {
        private TxObjGridCtrl _grid;
        private Func<ITxObject, bool> _keep;
        private Action<ITxObject> _onDrop;
        private Action _onChanged;
        private Action<string> _log;
        private bool _pruning;

        public void Init(Panel host, Func<ITxObject, bool> keep, Action<ITxObject> onDrop, Action onChanged, Action<string> log)
        {
            _keep = keep; _onDrop = onDrop; _onChanged = onChanged; _log = log ?? (s => { });
            try
            {
                _grid = new TxObjGridCtrl
                {
                    Dock = DockStyle.Fill,
                    ListenToPick = true,
                    EnableMultipleSelection = true,
                    EnableRecurringObjects = false
                };
                _grid.ObjectInserted += new TxObjGridCtrl_ObjectInsertedEventHandler(OnInserted);
                _grid.RowDeleted += new TxObjGridCtrl_RowDeletedEventHandler(OnRowDeleted);
                host.Controls.Add(_grid);
                _log("[Grid] 就绪（点网格拾取行→PS 中选对象；仅焊接操作，非操作自动剔除）");
            }
            catch (Exception ex) { _grid = null; _log("[Grid] 创建 TxObjGridCtrl 失败：" + ex.Message); }
        }

        private void OnInserted(object sender, TxObjGridCtrl_ObjectInsertedEventArgs e) { Prune(); _onChanged?.Invoke(); }
        private void OnRowDeleted(object sender, TxObjGridCtrl_RowDeletedEventArgs e) { _onChanged?.Invoke(); }

        // 插入后实时剔除非焊接操作
        private void Prune()
        {
            if (_pruning || _grid == null || _keep == null) return;
            _pruning = true;
            try
            {
                for (int i = _grid.Count - 1; i >= 0; i--)
                {
                    var o = _grid.GetObject(i) as ITxObject;
                    if (o != null && !_keep(o)) { _grid.DeleteRow(i); _onDrop?.Invoke(o); }
                }
            }
            catch { }
            finally { _pruning = false; }
        }

        public void Append(ITxObject o)
        {
            if (o == null || _grid == null) return;
            try { _grid.AppendObject(o); } catch (Exception ex) { _log("[Grid] AppendObject 失败：" + ex.Message); }
        }

        public void Clear()
        {
            if (_grid == null) return;
            try { for (int i = _grid.Count - 1; i >= 0; i--) _grid.DeleteRow(i); }
            catch (Exception ex) { _log("[Grid] 清空失败：" + ex.Message); }
        }

        public List<ITxObject> GetObjects()
        {
            var list = new List<ITxObject>();
            if (_grid == null) return list;
            try { int n = _grid.Count; for (int i = 0; i < n; i++) { var o = _grid.GetObject(i) as ITxObject; if (o != null) list.Add(o); } }
            catch (Exception ex) { _log("[Grid] 读对象失败：" + ex.Message); }
            return list;
        }

        public int Count { get { try { return _grid?.Count ?? 0; } catch { return 0; } } }
    }
}
