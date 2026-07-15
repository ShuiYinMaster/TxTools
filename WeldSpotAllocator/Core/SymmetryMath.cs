// SymmetryMath.cs  —  C# 7.3
// 关于对称中心 XZ 平面（Y 翻转）的镜像。复用 PsReader 的 TxTransformation 工具，不自造矩阵库。
//
// 【焊钳坐标约定（实测）】 X=前后  Y=左右  Z=上下
//   左右对称 ⇒ 对称面 = XZ 平面（法向 = Y 左右）⇒ 反射矩阵 M = diag(1,-1,1,1)，位置 Y→-Y。✔
//
// 世界系镜像公式：
//   局部系反射矩阵 M = diag(1,-1,1,1)
//   物体位姿在对称中心局部系 T_local = C⁻¹·T_world
//   反射后（物体本身也被镜像）  T'_local = M·T_local·M
//   回世界系                    T'_world = C·T'_local = C·M·C⁻¹·T_world·M
//   → M·R·M 行列式 +1，仍是合法旋转；位置 (x,y,z)→(x,-y,z)。
//
// 焊钳手性：镜像后工具系手性翻转，实际枪型常需再绕工具自身某轴翻 180°（FlipAxis，UI 实测选）。
//   按上面坐标约定：绕 X=前后翻 / 绕 Y=左右翻 / 绕 Z=上下翻，见 ApplyToolFlip。

using Tecnomatix.Engineering;
using TxTools.ExportGun;

namespace TxTools.WeldSpotAllocator
{
    public static class SymmetryMath
    {
        // 关于局部 XZ 平面的反射：diag(1,-1,1,1)
        private static readonly double[] MIRROR_XZ = { 1,0,0,0,  0,-1,0,0,  0,0,1,0,  0,0,0,1 };
        private static readonly TxTransformation M = PsReader.ArrToTxPublic(MIRROR_XZ);

        /// <summary>世界系位姿 T 关于中心 C 的 XZ 平面镜像。flip 用于焊枪手性修正。</summary>
        public static TxTransformation MirrorWorld(TxTransformation T, TxTransformation C, FlipAxis flip = FlipAxis.None)
        {
            var local  = TxTransformation.Multiply(C.Inverse, T);               // C⁻¹·T
            var mLocal = TxTransformation.Multiply(TxTransformation.Multiply(M, local), M); // M·local·M
            var world  = TxTransformation.Multiply(C, mLocal);                  // C·(…)
            return flip == FlipAxis.None ? world : ApplyToolFlip(world, flip);
        }

        /// <summary>double[16] 版本：返回镜像后的完整位姿 double[16]。</summary>
        public static double[] MirrorWorld(double[] worldMat, double[] centerMat, FlipAxis flip = FlipAxis.None)
        {
            var w = MirrorWorld(PsReader.ArrToTxPublic(worldMat), PsReader.ArrToTxPublic(centerMat), flip);
            return PsReader.TxToArr(w);
        }

        /// <summary>镜像后绕工具自身某轴翻 180°（右乘工具系旋转）。轴向：X=前后 Y=左右 Z=上下。</summary>
        private static TxTransformation ApplyToolFlip(TxTransformation w, FlipAxis flip)
        {
            double[] rot;
            switch (flip)
            {
                case FlipAxis.X: rot = new double[] { 1,0,0,0,  0,-1,0,0,  0,0,-1,0,  0,0,0,1 }; break; // 绕X(前后)翻：左右↔、上下↔
                case FlipAxis.Y: rot = new double[] { -1,0,0,0, 0,1,0,0,  0,0,-1,0,  0,0,0,1 }; break;  // 绕Y(左右)翻：前后↔、上下↔
                case FlipAxis.Z: rot = new double[] { -1,0,0,0, 0,-1,0,0, 0,0,1,0,   0,0,0,1 }; break;  // 绕Z(上下)翻：前后↔、左右↔
                default: return w;
            }
            return TxTransformation.Multiply(w, PsReader.ArrToTxPublic(rot));
        }

        /// <summary>把参考焊点解析成“用于匹配/姿态复制”的位姿。统一 B/C × 是否不同参考系。
        /// B 同系→原样；B 不同系→变到参考车件局部(对齐车身在原点的目标)；
        /// C 同系→世界镜像；C 不同系→局部镜像。</summary>
        public static double[] ResolveRef(double[] refWorld, double[] center, AllocMode mode, bool diffFrame, FlipAxis flip)
        {
            var T = PsReader.ArrToTxPublic(refWorld);
            if (mode == AllocMode.Symmetric)
                return diffFrame ? PsReader.TxToArr(MirrorLocal(T, center, flip))
                                 : MirrorWorld(refWorld, center, flip);
            // NewSpot
            return diffFrame ? PsReader.TxToArr(ToLocal(T, center)) : refWorld;
        }

        // C⁻¹·T：变到车件局部坐标系
        private static TxTransformation ToLocal(TxTransformation T, double[] center)
        {
            var C = PsReader.ArrToTxPublic(center);
            return TxTransformation.Multiply(C.Inverse, T);
        }

        // 局部系内镜像：M·(C⁻¹·T)·M，再按 flip 修正工具轴
        private static TxTransformation MirrorLocal(TxTransformation T, double[] center, FlipAxis flip)
        {
            var local = ToLocal(T, center);
            var m = TxTransformation.Multiply(TxTransformation.Multiply(M, local), M);
            return flip == FlipAxis.None ? m : ApplyToolFlip(m, flip);
        }

        /// <summary>欧氏距离平方（避免开方，比较够用）。</summary>
        public static double Dist2(double[] a, double[] b)
        {
            double dx = a[0]-b[0], dy = a[1]-b[1], dz = a[2]-b[2];
            return dx*dx + dy*dy + dz*dz;
        }
    }
}
