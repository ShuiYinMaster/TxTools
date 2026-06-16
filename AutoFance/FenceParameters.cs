using System.Drawing;

namespace TxTools.FenceBuilder
{
    public enum BaseplateMode
    {
        None,           // 立柱直接落地,无独立底板
        WithBaseplate   // 立柱底加法兰底板
    }

    public class FenceParameters
    {
        // ===== 网片参数 =====
        public double MeshNominalWidth = 1000.0;   // 标称网片宽度(mm) - 默认 1000
        public double MeshHeight = 1800.0;          // 网片高度
        public double MeshThickness = 2.0;          // 网片薄板物理厚度
        public double MeshFrameWidth = 40.0;        // 外框方管截面宽
        public double MeshFrameThickness = 40.0;    // 外框方管截面厚
        public double GroundClearance = 150.0;      // 网片下沿离地间隙 - 默认 150
        public bool EnableMeshTexture = true;
        // 默认黄色
        public Color MeshColor = Color.FromArgb(255, 200, 0);
        public Color FrameColor = Color.FromArgb(230, 180, 0);

        // ===== 立柱参数 =====
        public double PostWidth = 60.0;
        public double PostThickness = 60.0;
        public double PostGap = 5.0;                // 立柱与网片之间的水平间隙
        public double PostTopMargin = 100.0;        // 立柱顶高出网片顶的距离 - 默认 100
        public Color PostColor = Color.FromArgb(230, 180, 0);

        // ===== 底座参数 =====
        public BaseplateMode BaseplateMode = BaseplateMode.None;
        public double BaseplateWidth = 150.0;
        public double BaseplateLength = 150.0;
        public double BaseplateThickness = 10.0;
        public Color BaseplateColor = Color.FromArgb(200, 160, 0);

        // ===== 其他 =====
        public bool ShareCornerPost = true;
        public string GroupName = "Fence";

        // ===== 新增功能项 =====
        /// <summary>忽略线段 Z 向数值,围栏立柱底部强制贴地(GroundZ = 0)</summary>
        public bool IgnoreSegmentZ = true;

        /// <summary>末片合并阈值:若末片宽度 < 该值,则将这片合并到上一片(避免极窄末片)。0=不启用</summary>
        public double MinLastPanelWidth = 500.0;

        /// <summary>忽略线段之间夹角,强制每段方向量化到 ±X / ±Y 轴(横平竖直)</summary>
        public bool ForceAxisAligned = false;

        /// <summary>将生成的资源直接挂在当前建模资源下(而非 PhysicalRoot)。当前无建模资源时退化到 PhysicalRoot</summary>
        public bool CreateUnderActiveModeling = true;
    }
}
