// GroupModels.cs — C# 7.3
// 焊点自动分组插件的数据模型。
//
// 设计要点：
//   · SpotItem 同时持有「焊点位操作」(LocOp，移动用) 与「焊点特征」(Wp，读零件/名字用)。
//     真正能被 MoveLocationInto 移动的是 LocOp（一条 weld location operation），
//     而绑定的零件挂在 Wp(TxWeldPoint) 上。两者缺一不可。
//   · Signature = 多重集指纹：把绑定零件名「排序 + 保留重复」后用 \u0001 连接。
//     这样「集合相同 且 数量相同」才会得到同一个指纹（满足你选的判定口径）。

using System;
using System.Collections.Generic;
using Tecnomatix.Engineering;

namespace TxTools.WeldSpotGrouper
{
    /// <summary>单个焊点的快照：移动句柄 + 读取句柄 + 绑定零件 + 指纹。</summary>
    public sealed class SpotItem
    {
        public ITxObject LocOp;                 // 焊点位操作（weld location operation）—— 移动的对象
        public TxWeldPoint Wp;                  // 焊点特征 —— 读零件/名字
        public string Name;                     // 焊点名（取自 Wp.Name，回退 LocOp.Name）
        public readonly List<string> PartNames = new List<string>(); // 绑定零件名（原始顺序，含重复）
        public string Signature;                // 多重集指纹；无绑定时为 ""

        public bool HasBinding { get { return PartNames.Count > 0; } }
    }

    /// <summary>一个分组：相同指纹的所有焊点。</summary>
    public sealed class SpotGroup
    {
        public string Signature;                // 指纹（内部 key）
        public readonly List<string> SamplePartNames = new List<string>(); // 该组的零件名（排序后，用于命名/展示）
        public readonly List<SpotItem> Spots = new List<SpotItem>();
        public string TargetOpName;             // 将要创建的焊接操作名（执行前填）
        public ITxObject CreatedOp;             // 执行后创建出的操作（回填）

        /// <summary>零件名拼出的可读标签，如 "PartA + PartB"。</summary>
        public string PartLabel
        {
            get { return SamplePartNames.Count == 0 ? "(无绑定)" : string.Join(" + ", SamplePartNames); }
        }
    }

    /// <summary>分组选项。</summary>
    public sealed class GroupOptions
    {
        public string NamePrefix = "焊点分组_";  // 新建操作名前缀
        public bool IgnoreCase = false;          // 零件名比对是否忽略大小写（PS 名字一般精确，默认 false）
        public bool SkipUnbound = true;          // 跳过无绑定零件的点（过渡点/游离点），不参与分组
        public bool DryRun = false;              // true=只预览不写入
    }

    /// <summary>执行报告。</summary>
    public sealed class GroupReport
    {
        public int ScannedSpots;     // 扫描到的焊点位总数
        public int BoundSpots;       // 有绑定的焊点数
        public int SkippedUnbound;   // 跳过的无绑定点
        public int GroupsCreated;    // 成功新建的操作数
        public int SpotsMoved;       // 成功移动的焊点数
        public int Failed;           // 失败计数
        public readonly List<string> Errors = new List<string>();

        public override string ToString()
        {
            return string.Format(
                "扫描 {0} 个焊点位 · 有绑定 {1} · 跳过无绑定 {2} · 新建操作 {3} · 移动焊点 {4} · 失败 {5}",
                ScannedSpots, BoundSpots, SkippedUnbound, GroupsCreated, SpotsMoved, Failed);
        }
    }
}
