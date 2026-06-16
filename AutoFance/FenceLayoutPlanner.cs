using System;
using System.Collections.Generic;
using Tecnomatix.Engineering;

namespace TxTools.FenceBuilder
{
    public class PostPlacement
    {
        public TxVector CenterXY;
        public TxVector DirAlong;
        public TxVector DirNormal;
        public string Note;
        public int SegmentIndex;  // -1 表示拐角立柱
    }

    public class MeshPanelPlacement
    {
        public TxVector LeftPostCenter;
        public TxVector RightPostCenter;
        public TxVector DirAlong;
        public TxVector DirNormal;
        public double ActualWidth;
        public bool IsTruncated;
        public int SegmentIndex;
        public int IndexInSegment;
    }

    public class FenceLayout
    {
        public string SourceChainName;
        public double GroundZ;
        public List<PostPlacement> Posts = new List<PostPlacement>();
        public List<MeshPanelPlacement> Panels = new List<MeshPanelPlacement>();
    }

    /// <summary>
    /// 围栏布局规划(纯几何,无 SDK 依赖)
    ///
    /// 布局约定:
    ///   - 端点立柱中心 = 线段端点(不缩进)
    ///   - 立柱步长 stride = MeshWidth + PostWidth + 2*PostGap
    ///     即沿基线: 立柱 + 间隙 + 网片(完整4框) + 间隙 + 立柱 + 间隙 + 网片 + ...
    ///   - 网片在两个相邻立柱之间,宽度 = 立柱中心距 - PostWidth - 2*PostGap
    ///   - 末片自动截断
    ///   - 拐角共享立柱(相邻段端点立柱距离 < 0.5*PostWidth)
    ///   - 每个网片是独立完整体(4 框 + 1 薄板),不共享立边
    /// </summary>
    public static class FenceLayoutPlanner
    {
        public static FenceLayout Plan(BaseSegmentChain chain, FenceParameters p, Action<string> log)
        {
            FenceLayout layout = new FenceLayout
            {
                SourceChainName = chain.FeatureName,
                GroundZ = chain.GroundZ
            };
            if (chain.Segments.Count == 0) return layout;

            // 1) 处理 IgnoreSegmentZ: 强制 GroundZ = 0 (世界地面)
            if (p.IgnoreSegmentZ)
            {
                if (Math.Abs(chain.GroundZ) > 1e-6)
                    log?.Invoke("[Planner] IgnoreSegmentZ: 原 GroundZ=" + chain.GroundZ.ToString("F2") + " → 0");
                layout.GroundZ = 0.0;
            }

            double stride = p.MeshNominalWidth + p.PostWidth + 2.0 * p.PostGap;
            List<List<PostPlacement>> perSegPosts = new List<List<PostPlacement>>();

            for (int si = 0; si < chain.Segments.Count; si++)
            {
                BaseSegment seg = chain.Segments[si];
                var posts = new List<PostPlacement>();

                // 2) ForceAxisAligned: 量化段方向到 ±X / ±Y 轴
                TxVector segStart = seg.Start;
                TxVector segEnd = seg.End;
                if (p.ForceAxisAligned)
                {
                    QuantizeToAxis(ref segStart, ref segEnd, log, si);
                }
                if (p.IgnoreSegmentZ)
                {
                    segStart = new TxVector(segStart.X, segStart.Y, 0);
                    segEnd = new TxVector(segEnd.X, segEnd.Y, 0);
                }

                double dx = segEnd.X - segStart.X, dy = segEnd.Y - segStart.Y;
                double segLen = Math.Sqrt(dx * dx + dy * dy);
                if (segLen < p.PostWidth + 1.0)
                {
                    log?.Invoke("[Planner] 段 #" + si + " 过短(" + segLen.ToString("F1") + "mm),跳过");
                    perSegPosts.Add(posts);
                    continue;
                }

                TxVector dir = new TxVector(dx / segLen, dy / segLen, 0);
                TxVector nrm = new TxVector(-dir.Y, dir.X, 0);

                TxVector startPostCenter = segStart;
                TxVector endPostCenter = segEnd;

                posts.Add(new PostPlacement
                {
                    CenterXY = startPostCenter,
                    DirAlong = dir, DirNormal = nrm,
                    Note = "段#" + si + "-起",
                    SegmentIndex = si
                });

                // 3) 先用临时列表收集中立柱与网片,稍后处理末片合并
                List<TxVector> midPostCenters = new List<TxVector>();
                List<MeshPanelPlacement> segPanels = new List<MeshPanelPlacement>();

                double cursor = 0;
                int panelIdx = 0;
                while (cursor + p.PostWidth < segLen - 0.5)
                {
                    double remain = segLen - cursor;
                    double panelStride = Math.Min(stride, remain);
                    double panelWidth = panelStride - p.PostWidth - 2.0 * p.PostGap;
                    bool truncated = (panelStride < stride - 0.5);

                    if (panelWidth < 10.0)
                    {
                        log?.Invoke("[Planner] 段 #" + si + " 末段残余过小(" + panelWidth.ToString("F1") + "mm),不放网片");
                        break;
                    }

                    TxVector leftPost = AddXY(startPostCenter, dir, cursor);
                    TxVector rightPost = AddXY(startPostCenter, dir, cursor + panelStride);

                    segPanels.Add(new MeshPanelPlacement
                    {
                        LeftPostCenter = leftPost,
                        RightPostCenter = rightPost,
                        DirAlong = dir,
                        DirNormal = nrm,
                        ActualWidth = panelWidth,
                        IsTruncated = truncated,
                        SegmentIndex = si,
                        IndexInSegment = panelIdx
                    });

                    cursor += panelStride;
                    panelIdx++;

                    // 仅记录中立柱位置(此时不加入 posts,以便后续合并末片时可移除最后一个中立柱)
                    if (!truncated && cursor + p.PostWidth < segLen - 0.5)
                    {
                        midPostCenters.Add(rightPost);
                    }
                }

                // 4) 末片合并: 若末片宽度 < MinLastPanelWidth, 把末片吸收到上一片
                if (p.MinLastPanelWidth > 0 && segPanels.Count >= 2)
                {
                    var lastPanel = segPanels[segPanels.Count - 1];
                    if (lastPanel.ActualWidth < p.MinLastPanelWidth)
                    {
                        var prevPanel = segPanels[segPanels.Count - 2];
                        // 上一片网片右扩到末片右立柱位置
                        // 新右立柱 = 原末片右立柱;新宽度 = 原宽 + (末片立柱中心距上一片右立柱中心 + 末片宽 + 间隙等) 
                        // 直接用几何重算:新宽 = 原宽 + PostWidth + 2*PostGap + 末片宽
                        double newWidth = prevPanel.ActualWidth + p.PostWidth + 2.0 * p.PostGap + lastPanel.ActualWidth;
                        prevPanel.RightPostCenter = lastPanel.RightPostCenter;
                        prevPanel.ActualWidth = newWidth;
                        prevPanel.IsTruncated = true;
                        // 移除末片
                        segPanels.RemoveAt(segPanels.Count - 1);
                        // 移除上一片与末片之间的中立柱(midPostCenters 的最后一个)
                        if (midPostCenters.Count > 0)
                            midPostCenters.RemoveAt(midPostCenters.Count - 1);
                        log?.Invoke("[Planner] 段 #" + si + " 末片(" + lastPanel.ActualWidth.ToString("F0") +
                                    "mm)< 阈值" + p.MinLastPanelWidth + ",已合并到上一片 → 新宽 " + newWidth.ToString("F0"));
                    }
                }

                // 5) 输出: 把中立柱加入 posts
                for (int mi = 0; mi < midPostCenters.Count; mi++)
                {
                    posts.Add(new PostPlacement
                    {
                        CenterXY = midPostCenters[mi],
                        DirAlong = dir, DirNormal = nrm,
                        Note = "段#" + si + "-中#" + (mi + 1),
                        SegmentIndex = si
                    });
                }

                layout.Panels.AddRange(segPanels);

                posts.Add(new PostPlacement
                {
                    CenterXY = endPostCenter,
                    DirAlong = dir, DirNormal = nrm,
                    Note = "段#" + si + "-终",
                    SegmentIndex = si
                });

                log?.Invoke("[Planner] 段 #" + si + " 长度=" + segLen.ToString("F1") +
                            "  网片=" + panelIdx + "  立柱=" + posts.Count);
                perSegPosts.Add(posts);
            }

            // 拐角合并
            double mergeTol = Math.Max(p.PostWidth, p.PostThickness) * 0.5;
            for (int si = 0; si < perSegPosts.Count; si++)
            {
                var current = perSegPosts[si];
                for (int pi = 0; pi < current.Count; pi++)
                {
                    PostPlacement post = current[pi];
                    if (p.ShareCornerPost &&
                        pi == current.Count - 1 &&
                        si < perSegPosts.Count - 1)
                    {
                        var next = perSegPosts[si + 1];
                        if (next.Count > 0 && Distance2D(post.CenterXY, next[0].CenterXY) < mergeTol)
                        {
                            TxVector mid = new TxVector(
                                (post.CenterXY.X + next[0].CenterXY.X) * 0.5,
                                (post.CenterXY.Y + next[0].CenterXY.Y) * 0.5,
                                post.CenterXY.Z);
                            TxVector avgDir = NormalizeXY(new TxVector(
                                post.DirAlong.X + next[0].DirAlong.X,
                                post.DirAlong.Y + next[0].DirAlong.Y, 0));
                            TxVector avgNrm = new TxVector(-avgDir.Y, avgDir.X, 0);

                            layout.Posts.Add(new PostPlacement
                            {
                                CenterXY = mid,
                                DirAlong = avgDir, DirNormal = avgNrm,
                                Note = "拐角#" + si + "-" + (si + 1),
                                SegmentIndex = -1
                            });
                            perSegPosts[si + 1].RemoveAt(0);
                            continue;
                        }
                    }
                    layout.Posts.Add(post);
                }
            }

            log?.Invoke("[Planner] 完成: 立柱=" + layout.Posts.Count + " 网片=" + layout.Panels.Count);
            return layout;
        }

        private static TxVector AddXY(TxVector p, TxVector dir, double dist)
        {
            return new TxVector(p.X + dir.X * dist, p.Y + dir.Y * dist, p.Z);
        }
        private static double Distance2D(TxVector a, TxVector b)
        {
            double dx = a.X - b.X, dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }
        private static TxVector NormalizeXY(TxVector v)
        {
            double L = Math.Sqrt(v.X * v.X + v.Y * v.Y);
            if (L < 1e-9) return new TxVector(1, 0, 0);
            return new TxVector(v.X / L, v.Y / L, 0);
        }

        /// <summary>
        /// 把段的两端点量化到坐标轴对齐:
        ///   - 比较 |dx| 与 |dy|,大的方向作为主轴方向,小的方向归零
        ///   - 段起点保留,终点 = 起点 + 主轴方向 * 段长(用 max(|dx|,|dy|) 作段长)
        /// 效果:斜着的线段会被强制变成纯水平或纯竖直,长度近似保持。
        /// </summary>
        private static void QuantizeToAxis(ref TxVector segStart, ref TxVector segEnd,
            Action<string> log, int si)
        {
            double dx = segEnd.X - segStart.X;
            double dy = segEnd.Y - segStart.Y;
            if (Math.Abs(dx) >= Math.Abs(dy))
            {
                // 主轴 X
                segEnd = new TxVector(segEnd.X, segStart.Y, segEnd.Z);
            }
            else
            {
                // 主轴 Y
                segEnd = new TxVector(segStart.X, segEnd.Y, segEnd.Z);
            }
            log?.Invoke("[Planner] ForceAxisAligned: 段#" + si + " 量化为 dx=" +
                        (segEnd.X - segStart.X).ToString("F1") + " dy=" + (segEnd.Y - segStart.Y).ToString("F1"));
        }
    }
}
