using System;

namespace TxTools.AutoPathPlanner
{
    /// <summary>轻量三维向量 — 纯算法层，不依赖SDK类型</summary>
    public struct Vec3
    {
        public double X, Y, Z;
        public Vec3(double x, double y, double z) { X = x; Y = y; Z = z; }

        public static Vec3 operator +(Vec3 a, Vec3 b) { return new Vec3(a.X + b.X, a.Y + b.Y, a.Z + b.Z); }
        public static Vec3 operator -(Vec3 a, Vec3 b) { return new Vec3(a.X - b.X, a.Y - b.Y, a.Z - b.Z); }
        public static Vec3 operator *(Vec3 a, double s) { return new Vec3(a.X * s, a.Y * s, a.Z * s); }

        public double Length() { return Math.Sqrt(X * X + Y * Y + Z * Z); }

        public Vec3 Normalized()
        {
            double len = Length();
            return len < 1e-9 ? new Vec3(0, 0, 1) : this * (1.0 / len);
        }

        public static double Distance(Vec3 a, Vec3 b) { return (a - b).Length(); }

        public static Vec3 Lerp(Vec3 a, Vec3 b, double t)
        {
            return new Vec3(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t, a.Z + (b.Z - a.Z) * t);
        }

        public override string ToString() { return string.Format("({0:F1},{1:F1},{2:F1})", X, Y, Z); }
    }

    /// <summary>轴对齐包围盒 — RRT 采样空间</summary>
    public struct Aabb
    {
        public Vec3 Min, Max;

        public static Aabb FromPoints(Vec3 a, Vec3 b)
        {
            return new Aabb
            {
                Min = new Vec3(Math.Min(a.X, b.X), Math.Min(a.Y, b.Y), Math.Min(a.Z, b.Z)),
                Max = new Vec3(Math.Max(a.X, b.X), Math.Max(a.Y, b.Y), Math.Max(a.Z, b.Z))
            };
        }

        public Aabb Inflate(double xy, double zUp, double zDown)
        {
            return new Aabb
            {
                Min = new Vec3(Min.X - xy, Min.Y - xy, Min.Z - zDown),
                Max = new Vec3(Max.X + xy, Max.Y + xy, Max.Z + zUp)
            };
        }

        public Vec3 Sample(Random rng)
        {
            return new Vec3(
                Min.X + rng.NextDouble() * (Max.X - Min.X),
                Min.Y + rng.NextDouble() * (Max.Y - Min.Y),
                Min.Z + rng.NextDouble() * (Max.Z - Min.Z));
        }
    }
}
