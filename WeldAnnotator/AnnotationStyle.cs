using System.Collections.Generic;
using D = System.Drawing;

namespace TxTools.WeldAnnotator
{
    /// <summary>标注文字命名模式。</summary>
    public enum LabelNamingMode
    {
        Sequence   = 0,  // 按顺序号（1, 2, 3, ...）
        PointName  = 1,  // 按焊点名称
        Prefix     = 2,  // 前缀 + 序号
        Suffix     = 3,  // 序号 + 后缀
        SeqAndName = 4   // 序号 + 焊点名称
    }

    /// <summary>标注样式配置。</summary>
    public class AnnotationStyle
    {
        public D.Color DotColor         = D.Color.Red;
        public int     DotRadius        = 8;
        public D.Color LineColor        = D.Color.Red;
        public float   LineWidth        = 1.5f;
        public D.Color BoxBorderColor   = D.Color.Black;
        public D.Color BoxFillColor     = D.Color.White;
        public D.Color TextColor        = D.Color.Black;
        public D.Font  TextFont         = new D.Font("Arial", 9f, D.FontStyle.Regular);
        public int     BoxPadding       = 4;
        public int     OffsetX          = 40;
        public int     OffsetY          = 40;

        /// <summary>分类 → Excel MsoAutoShapeType 形状 ID。</summary>
        public Dictionary<string, int> CategoryShapes
            = new Dictionary<string, int>
            {
                { "二层板点焊",         9  },   // 圆
                { "二层板补焊",         9  },
                { "三层板及以上点焊",   7  },   // 三角形
                { "三层板及以上补焊",   7  },
                { "焊缝",              1  },   // 矩形
                { "CO2焊点",           1  },
                { "螺母",              10 },   // 六边形
                { "螺钉螺栓",          1  },
                { "胶",                5  },   // 圆角矩形
                { "强度校验点",         12 },   // 五角星
                { "重要特性",          12 },
                { "关键特性",          12 }
            };

        /// <summary>分类 → 是否实心（true=填充颜色，false=空心仅描边）。</summary>
        public Dictionary<string, bool> CategoryFilled
            = new Dictionary<string, bool>
            {
                { "二层板点焊",         true  },
                { "二层板补焊",         false },
                { "三层板及以上点焊",   true  },
                { "三层板及以上补焊",   false },
                { "焊缝",              true  },
                { "CO2焊点",           true  },
                { "螺母",              true  },
                { "螺钉螺栓",          false },
                { "胶",                false },
                { "强度校验点",         true  },
                { "重要特性",          false },
                { "关键特性",          true  }
            };
    }
}
