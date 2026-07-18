using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Tecnomatix.Engineering;

using Keys = System.Windows.Forms.Keys;

namespace TxTools.MechArena
{
    // =========================================================================
    //  游戏引擎
    // =========================================================================
    public class MechArenaEngine : IDisposable
    {
        public PlayerEntity Player { get; private set; }
        public readonly List<ProjectileEntity> Projectiles = new List<ProjectileEntity>();
        public readonly List<EnemyRobotAgent> Enemies = new List<EnemyRobotAgent>();
        public BossFormation Boss { get; private set; }

        public int Score;
        public double Time;
        public bool GameOver;
        public bool Victory;

        // ---------- 调参 ----------
        private const double TICK_SECONDS = 0.05;
        private const double SHOT_COOLDOWN = 0.20;
        private const double BULLET_SPEED = 300;
        private const int BULLET_LIFETIME = 80;
        private const double PLAYER_HIT_RADIUS = 700;
        private const int BULLET_DAMAGE = 15;
        public const int ENEMY_MAX_HP = 100;
        public const int BOSS_MEMBER_HP = 200;

        // 视口裁剪：玩家周围此范围外的机器人跳过 UpdateAttack
        private const double UPDATE_RANGE = 15000;  // mm（15m）

        // ---------- 相机（鼠标轨道 + 焦点锁定玩家）----------
        private const double CAM_YAW_SENS = 0.008;   // 鼠标水平灵敏度（rad/px）
        private const double CAM_PITCH_SENS = 0.006;   // 鼠标垂直灵敏度（rad/px），取负可反转
        private const double CAM_PITCH_MIN = 0.08;    // 最低俯角（rad）
        private const double CAM_PITCH_MAX = 1.45;    // 最高俯角（rad）
        private const double CAM_DIST_MIN = 2500;
        private const double CAM_DIST_MAX = 25000;
        private const double CAM_SMOOTH = 0.35;    // 焦点平滑系数 0..1

        /// <summary>视角是否锁定在玩家身上。Esc 解除，F 重新锁定。</summary>
        public bool CameraFocusLocked { get; private set; } = true;

        private double _camYaw = Math.PI / 4;   // 初始 45°
        private double _camPitch = 0.7;           // 初始约 40°
        private double _camDist = 10000;
        private TxVector _camCenter;
        private bool _camInit;

        // ---------- 干涉集 ----------
        private MechArenaCollisionService _collision;
        public string CollisionStatus
        {
            get { return _collision != null ? _collision.StatusText : "未启用"; }
        }

        private double _lastShotTime = -10;

        // =====================================================================
        public void Initialize()
        {
            Player = new PlayerEntity();
            Player.Body = MechArenaGeometry.CreateBox(
                "MechArena_Player_" + Guid.NewGuid().ToString("N").Substring(0, 6),
                300, 300, 300, new TxColor(60, 220, 80));
            Player.SetPosition(0, 0, 500);
            Player.Health = 100;

            var doc = TxApplication.ActiveDocument;
            if (doc == null) throw new Exception("当前没有活动 study");

            var filter = new TxTypeFilter(typeof(TxRobot));
            var all = doc.PhysicalRoot.GetAllDescendants(filter);
            foreach (ITxObject o in all)
            {
                var r = o as TxRobot;
                if (r == null) continue;
                var agent = new EnemyRobotAgent { Robot = r, Health = ENEMY_MAX_HP };
                agent.Init();
                Enemies.Add(agent);
            }

            Boss = new BossFormation();

            // 干涉集：玩家×机器人 + 子弹×机器人（子弹对懒创建）
            try
            {
                _collision = new MechArenaCollisionService();
                _collision.Initialize(Player.Body, Enemies.Select(e => e.Robot));
            }
            catch { _collision = null; }

            // 保存原相机（Dispose 时还原），锁定焦点到玩家
            MechArenaCamera.SaveCurrentCamera();
            CameraFocusLocked = true;
            _camInit = false;
            UpdateFollowCamera(force: true);
        }

        /// <summary>Esc 解除视角锁定 / F 重新锁定（Form 调用）。</summary>
        public void SetCameraFocus(bool locked)
        {
            CameraFocusLocked = locked;
            if (locked) _camInit = false;   // 重新锁定时直接对准玩家
        }

        /// <summary>
        /// 轨道相机：焦点始终为玩家，yaw/pitch 由鼠标增量控制，距离键控。
        /// 未锁定时完全不碰相机（交还 PS 原生鼠标操作）。
        /// </summary>
        private void UpdateFollowCamera(bool force = false)
        {
            if (!CameraFocusLocked) return;
            try
            {
                var target = new TxVector(Player.X, Player.Y, Player.Z);
                if (force || !_camInit)
                {
                    _camCenter = target;
                    _camInit = true;
                }
                else
                {
                    _camCenter = new TxVector(
                        _camCenter.X + (target.X - _camCenter.X) * CAM_SMOOTH,
                        _camCenter.Y + (target.Y - _camCenter.Y) * CAM_SMOOTH,
                        _camCenter.Z + (target.Z - _camCenter.Z) * CAM_SMOOTH);
                }

                double cp = Math.Cos(_camPitch), sp = Math.Sin(_camPitch);
                var camPos = new TxVector(
                    _camCenter.X + _camDist * cp * Math.Cos(_camYaw),
                    _camCenter.Y + _camDist * cp * Math.Sin(_camYaw),
                    _camCenter.Z + _camDist * sp);

                // refresh:false —— Update 末尾统一 RefreshDisplay
                MechArenaCamera.SetLookAtCamera(_camCenter, camPos, refresh: false);
            }
            catch { }
        }

        // =====================================================================
        public void Update(HashSet<Keys> keys, bool fire, int mouseDx, int mouseDy)
        {
            if (GameOver || Victory) return;
            Time += TICK_SECONDS;

            // ---- 相机输入（即使不移动也生效）----
            _camYaw += mouseDx * CAM_YAW_SENS;
            _camPitch = Clamp(_camPitch + mouseDy * CAM_PITCH_SENS,
                CAM_PITCH_MIN, CAM_PITCH_MAX);
            if (keys.Contains(Keys.PageUp) || keys.Contains(Keys.Add) ||
                keys.Contains(Keys.Oemplus))
                _camDist = Math.Max(CAM_DIST_MIN, _camDist * 0.95);
            if (keys.Contains(Keys.PageDown) || keys.Contains(Keys.Subtract) ||
                keys.Contains(Keys.OemMinus))
                _camDist = Math.Min(CAM_DIST_MAX, _camDist * 1.05);

            Player.MoveByInput(keys);

            if (fire && Time - _lastShotTime > SHOT_COOLDOWN)
            {
                FireProjectile();
                _lastShotTime = Time;
            }

            for (int i = Projectiles.Count - 1; i >= 0; i--)
            {
                var p = Projectiles[i];
                p.Update();
                if (p.TimeToLive <= 0) RemoveProjectileAt(i);
            }

            var pp = new TxVector(Player.X, Player.Y, Player.Z);
            foreach (var e in Enemies)
            {
                if (e.Health <= 0) continue;

                double dx = e.BasePos.X - Player.X;
                double dy = e.BasePos.Y - Player.Y;
                double dz = e.BasePos.Z - Player.Z;
                double dist = Math.Sqrt(dx * dx + dy * dy + dz * dz);

                if (dist <= UPDATE_RANGE)
                {
                    e.UpdateAttack(Time, pp, Boss.Active);
                }
                else if (!e.WasOutOfViewport)
                {
                    e.RestoreOriginalPose();
                }
                e.WasOutOfViewport = (dist > UPDATE_RANGE);
            }

            if (!Boss.Active) Boss.TryActivate(Enemies);
            Boss.UpdateBoss();

            CheckCollisions(pp);

            UpdateFollowCamera();

            try { TxApplication.RefreshDisplay(); } catch { }

            if (Player.Health <= 0) GameOver = true;
            if (Enemies.Count > 0 && Enemies.All(x => x.Health <= 0)) Victory = true;
        }

        private static double Clamp(double v, double lo, double hi)
        {
            return v < lo ? lo : (v > hi ? hi : v);
        }

        // =====================================================================
        private void FireProjectile()
        {
            EnemyRobotAgent target = null;
            double minDist = double.MaxValue;
            foreach (var e in Enemies)
            {
                if (e.Health <= 0) continue;
                var c = e.AimCenter;
                double d = Dist(Player.X, Player.Y, Player.Z, c.X, c.Y, c.Z);
                if (d < minDist) { minDist = d; target = e; }
            }
            if (target == null) return;

            var t = target.AimCenter;
            double vx = t.X - Player.X;
            double vy = t.Y - Player.Y;
            double vz = t.Z - Player.Z;
            double len = Math.Sqrt(vx * vx + vy * vy + vz * vz);
            if (len < 1e-3) return;
            double s = BULLET_SPEED / len;
            vx *= s; vy *= s; vz *= s;

            TxComponent body;
            try
            {
                body = MechArenaGeometry.CreateBox(
                    "MechArena_Bullet_" + Guid.NewGuid().ToString("N").Substring(0, 6),
                    150, 150, 150, new TxColor(230, 80, 40));
            }
            catch { return; }

            var p = new ProjectileEntity
            {
                Body = body,
                Position = new TxVector(Player.X, Player.Y, Player.Z),
                Velocity = new TxVector(vx, vy, vz),
                TimeToLive = BULLET_LIFETIME,
                Damage = BULLET_DAMAGE
            };
            p.Apply();
            Projectiles.Add(p);

            // 子弹加入干涉集
            try { _collision?.AddBullet(body); } catch { }
        }

        /// <summary>统一的子弹移除：先出干涉对，再删几何。</summary>
        private void RemoveProjectileAt(int idx)
        {
            var p = Projectiles[idx];
            try { _collision?.RemoveBullet(p.Body); } catch { }
            try { p.Dispose(); } catch { }
            Projectiles.RemoveAt(idx);
        }

        // =====================================================================
        //  碰撞判定：干涉集优先，查询不可用时退回 AABB/距离数学检测
        // =====================================================================
        private void CheckCollisions(TxVector playerPos)
        {
            List<MechArenaCollisionService.CollidingPair> colls = null;
            bool useCS = _collision != null && _collision.QueryUsable;
            if (useCS)
            {
                colls = _collision.Query();
                useCS = _collision.QueryUsable;   // 查询过程中可能自我判废
            }

            // ---- 子弹 × 机器人 ----
            if (useCS)
            {
                var deadBullets = new HashSet<ProjectileEntity>();
                foreach (var c in colls)
                {
                    var proj = ResolveProjectile(c.A);
                    var other = c.B;
                    if (proj == null) { proj = ResolveProjectile(c.B); other = c.A; }
                    if (proj == null || deadBullets.Contains(proj)) continue;

                    var enemy = ResolveEnemy(other);
                    if (enemy == null || enemy.Health <= 0) continue;

                    DamageEnemy(enemy, proj.Damage);
                    deadBullets.Add(proj);
                }
                for (int i = Projectiles.Count - 1; i >= 0; i--)
                    if (deadBullets.Contains(Projectiles[i]))
                        RemoveProjectileAt(i);
            }
            else
            {
                // AABB 兜底（虚拟包围盒，无实体几何）
                for (int i = Projectiles.Count - 1; i >= 0; i--)
                {
                    var b = Projectiles[i];
                    bool hit = false;
                    foreach (var e in Enemies)
                    {
                        if (e.Health <= 0) continue;
                        if (e.CheckBulletHit(b.Position))
                        {
                            DamageEnemy(e, b.Damage);
                            hit = true;
                            break;
                        }
                    }
                    if (hit) RemoveProjectileAt(i);
                }
            }

            // ---- 机器人 × 玩家（仅 Strike 状态造成伤害）----
            foreach (var e in Enemies)
            {
                if (e.Health <= 0) continue;
                if (e.State != EnemyRobotAgent.AttackState.Strike) continue;

                bool hit = Dist(playerPos, e.GetAttackTip()) < PLAYER_HIT_RADIUS;
                if (!hit && useCS)
                {
                    // 干涉集补充：机器人本体与玩家方块真实接触也算命中
                    foreach (var c in colls)
                    {
                        bool aIsPlayer = IsPlayerObject(c.A);
                        bool bIsPlayer = IsPlayerObject(c.B);
                        if (!aIsPlayer && !bIsPlayer) continue;
                        var enemy = ResolveEnemy(aIsPlayer ? c.B : c.A);
                        if (enemy == e) { hit = true; break; }
                    }
                }
                if (hit)
                    Player.Health -= Boss.Active ? 8 : 4;
            }
        }

        private void DamageEnemy(EnemyRobotAgent e, int dmg)
        {
            e.Health -= dmg;
            Score += 10;
            if (e.Health <= 0)
            {
                Score += 100;
                e.OnKilled();
                try { _collision?.RemoveRobot(e.Robot); } catch { }
            }
        }

        // ---------- 对象归属解析 ----------
        private ProjectileEntity ResolveProjectile(ITxObject o)
        {
            if (o == null) return null;
            string n = null;
            try { n = o.Name; } catch { }
            foreach (var p in Projectiles)
            {
                if (p.Body == null) continue;
                try { if (p.Body.Equals(o)) return p; } catch { }
                try { if (n != null && p.Body.Name == n) return p; } catch { }
            }
            return null;
        }

        private bool IsPlayerObject(ITxObject o)
        {
            if (o == null || Player?.Body == null) return false;
            try { if (Player.Body.Equals(o)) return true; } catch { }
            try { return o.Name == Player.Body.Name; } catch { }
            return false;
        }

        /// <summary>
        /// 干涉结果中的机器人对象可能是机器人本身也可能是其子 link，
        /// Equals 失败时用位置最近兜底（PS 会返回不同 CLR 代理对象）。
        /// </summary>
        private EnemyRobotAgent ResolveEnemy(ITxObject o)
        {
            if (o == null) return null;
            foreach (var e in Enemies)
            {
                try { if (e.Robot != null && e.Robot.Equals(o)) return e; } catch { }
            }
            try
            {
                var lo = o as ITxLocatableObject;
                if (lo != null)
                {
                    var pos = lo.AbsoluteLocation.Translation;
                    EnemyRobotAgent best = null;
                    double bd = 5000;   // 5m 内最近
                    foreach (var e in Enemies)
                    {
                        if (e.Health <= 0) continue;
                        double d = Dist(pos, e.BasePos);
                        if (d < bd) { bd = d; best = e; }
                    }
                    return best;
                }
            }
            catch { }
            return null;
        }

        // =====================================================================
        /// <summary>
        /// 清理游戏对象 —— 干涉集 → 实体 → 相机 → 刷新。
        /// </summary>
        public void Dispose()
        {
            try { _collision?.Dispose(); } catch { }
            _collision = null;

            try { Player?.Dispose(); } catch { }
            foreach (var p in Projectiles) { try { p.Dispose(); } catch { } }
            Projectiles.Clear();

            foreach (var e in Enemies)
            {
                try { e.RestoreOriginalPose(); } catch { }
                try { e.RestoreOriginalLocation(); } catch { }
                try { MechArenaVisibility.TrySetVisible(e.Robot, true); } catch { }
                try { e.Dispose(); } catch { }
            }

            try { MechArenaCamera.RestoreSavedCamera(); } catch { }
            try { TxApplication.RefreshDisplay(); } catch { }
        }

        private static double Dist(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            double dx = x1 - x2, dy = y1 - y2, dz = z1 - z2;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }
        private static double Dist(TxVector a, TxVector b) { return Dist(a.X, a.Y, a.Z, b.X, b.Y, b.Z); }
    }

    // =========================================================================
    //  玩家（已取消 Q/E 升降，Z 固定）
    // =========================================================================
    public class PlayerEntity : IDisposable
    {
        public TxComponent Body;
        public double X, Y, Z;
        public int Health;
        public double MoveSpeed = 150;

        public void SetPosition(double x, double y, double z)
        {
            X = x; Y = y; Z = z;
            ApplyPosition();
        }

        public void MoveByInput(HashSet<Keys> keys)
        {
            bool moved = false;
            if (keys.Contains(Keys.W)) { X += MoveSpeed; moved = true; }
            if (keys.Contains(Keys.S)) { X -= MoveSpeed; moved = true; }
            if (keys.Contains(Keys.A)) { Y += MoveSpeed; moved = true; }
            if (keys.Contains(Keys.D)) { Y -= MoveSpeed; moved = true; }
            if (moved) ApplyPosition();
        }

        public void ApplyPosition()
        {
            if (Body == null) return;
            try
            {
                var t = new TxTransformation();
                t.Translation = new TxVector(X, Y, Z);
                Body.AbsoluteLocation = t;
            }
            catch { }
        }

        public void Dispose()
        {
            try { Body?.Delete(); } catch { }
            Body = null;
        }
    }

    // =========================================================================
    //  子弹
    // =========================================================================
    public class ProjectileEntity : IDisposable
    {
        public TxComponent Body;
        public TxVector Position;
        public TxVector Velocity;
        public int TimeToLive;
        public int Damage;

        public void Update()
        {
            Position = new TxVector(
                Position.X + Velocity.X,
                Position.Y + Velocity.Y,
                Position.Z + Velocity.Z);
            Apply();
            TimeToLive--;
        }

        public void Apply()
        {
            if (Body == null) return;
            try
            {
                var t = new TxTransformation();
                t.Translation = Position;
                Body.AbsoluteLocation = t;
            }
            catch { }
        }

        public void Dispose()
        {
            try { Body?.Delete(); } catch { }
            Body = null;
        }
    }

    // =========================================================================
    //  敌方机器人（实体碰撞盒已移除；AimCenter 为瞄准点/虚拟包围盒中心）
    // =========================================================================
    public class EnemyRobotAgent : IDisposable
    {
        public enum AttackState { Track, WindUp, Strike, Recover }

        public TxRobot Robot;
        public int Health;
        public double PhaseOffset;
        public AttackState State { get; private set; } = AttackState.Track;

        // 显示用名字（Form 里的血量列表要用）
        public string RobotName
        {
            get { try { return Robot?.Name ?? "?"; } catch { return "?"; } }
        }

        // 瞄准点 / 虚拟包围盒（仅数学，无实体几何；干涉集失效时的 AABB 兜底）
        public TxVector AimCenter;
        public const double AIM_BOX_SX = 2000;
        public const double AIM_BOX_SY = 2000;
        public const double AIM_BOX_SZ = 2500;

        // 内部
        private double[] _baseJoints;
        private int _jointCount;
        private TxVector _basePos;
        private TxVector _originalBasePos;                 // 恢复位置用
        private double _stateStartTime;
        private double[] _stateStartJoints;
        private double[] _stateTargetJoints;
        private TxVector _lockedTarget;
        private double _cooldownUntil;
        private bool _wasOutOfViewport;                    // 上一次是否在视口外
        private static readonly Random _rnd = new Random();

        // ---------- 移动候选梯队（本体 → 逐级父容器） ----------
        // 直接写 Robot.AbsoluteLocation 在部分场景无效（机器人被挂载/锁定），
        // 此时改为移动其所在容器（Equipment）。策略：
        //   候选0 = 机器人本体；候选1..n = 沿父链向上的每一级 ITxLocatableObject。
        //   首次移动做"写后读回"验证——机器人绝对位置没真动就升级到上一级候选，
        //   验证通过后锁定该候选，后续不再验证。
        private readonly List<ITxLocatableObject> _moveCandidates = new List<ITxLocatableObject>();
        private readonly List<TxTransformation> _moveOriginalLocs = new List<TxTransformation>();
        private readonly List<string> _moveCandidateTags = new List<string>();
        private int _moveIdx;
        private bool _moveVerified;

        /// <summary>当前生效的移动目标描述（调试用）。</summary>
        public string MoveTargetInfo
        {
            get
            {
                if (_moveCandidates.Count == 0) return "(无候选)";
                if (_moveIdx >= _moveCandidates.Count) return "(全部失效)";
                return _moveCandidateTags[_moveIdx] + (_moveVerified ? " ✓" : " ?");
            }
        }

        /// <summary>机器人基座世界坐标（供视口裁剪用）</summary>
        public TxVector BasePos { get { return _basePos; } }

        /// <summary>上一次是否在视口外（供视口裁剪状态切换用）</summary>
        public bool WasOutOfViewport
        {
            get { return _wasOutOfViewport; }
            set { _wasOutOfViewport = value; }
        }

        // 攻击时间
        private const double STRIKE_RANGE = 3500;
        private const double COOLDOWN_TIME = 1.5;
        private const double WINDUP_TIME = 0.4;
        private const double STRIKE_TIME = 0.25;
        private const double RECOVER_TIME = 0.5;

        // 移动参数
        private const double PURSUE_RANGE = 8000;         // 追击范围（超过不动）
        private const double KEEP_DISTANCE = 2200;         // 保持距离
        private const double MOVE_SPEED = 60;           // mm/tick ≈ 1.2 m/s

        // =====================================================================
        public void Init()
        {
            PhaseOffset = _rnd.NextDouble() * Math.PI * 2;

            // 读取关节：主路径 TxRobot.Joints 强类型（IkSolver.cs 已验证），
            // 回退 TxPoseData.JointValues（CollisionWorld 已验证）
            try
            {
                var joints = Robot.Joints;
                if (joints != null && joints.Count > 0)
                {
                    _jointCount = joints.Count;
                    _baseJoints = new double[_jointCount];
                    for (int i = 0; i < _jointCount; i++)
                    {
                        try { dynamic j = joints[i]; _baseJoints[i] = (double)j.CurrentValue; }
                        catch { _baseJoints[i] = 0.0; }
                    }
                }
                else throw new Exception("Joints 为空");
            }
            catch
            {
                try
                {
                    TxPoseData pose = Robot.CurrentPose;
                    ArrayList jv = pose.JointValues;
                    if (jv != null && jv.Count > 0)
                    {
                        _jointCount = jv.Count;
                        _baseJoints = new double[_jointCount];
                        for (int i = 0; i < _jointCount; i++)
                            _baseJoints[i] = Convert.ToDouble(jv[i]);
                    }
                    else throw new Exception("JointValues 为空");
                }
                catch
                {
                    _jointCount = 0;
                    _baseJoints = new double[0];
                }
            }

            try { _basePos = Robot.AbsoluteLocation.Translation; }
            catch { _basePos = new TxVector(0, 0, 0); }
            _originalBasePos = _basePos;

            BuildMoveCandidates();
            SyncAimCenter();
        }

        // =====================================================================
        //  移动候选梯队构建 / 移动 / 恢复
        // =====================================================================
        private void BuildMoveCandidates()
        {
            _moveCandidates.Clear();
            _moveOriginalLocs.Clear();
            _moveCandidateTags.Clear();
            _moveIdx = 0;
            _moveVerified = false;

            // 候选0：机器人本体
            AddMoveCandidate(Robot as ITxLocatableObject, "本体");

            // 候选1..n：逐级父容器（到 PhysicalRoot 为止）
            try
            {
                ITxObject root = null;
                try { root = TxApplication.ActiveDocument.PhysicalRoot; } catch { }

                ITxObject cur = Robot;
                for (int guard = 0; guard < 8; guard++)
                {
                    ITxObject parent = GetParentObject(cur);
                    if (parent == null) break;
                    if (parent is TxPhysicalRoot) break;
                    if (root != null && ReferenceEquals(parent, root)) break;

                    var loc = parent as ITxLocatableObject;
                    if (loc != null)
                    {
                        string tag = "父级";
                        try
                        {
                            var c = parent as ITxComponent;
                            if (c != null && c.IsEquipment) tag = "容器";
                        }
                        catch { }
                        string name = "?";
                        try { name = parent.Name; } catch { }
                        AddMoveCandidate(loc, tag + ":" + name);
                    }
                    cur = parent;
                }
            }
            catch { }
        }

        private void AddMoveCandidate(ITxLocatableObject obj, string tag)
        {
            if (obj == null) return;
            TxTransformation original = null;
            try { original = obj.AbsoluteLocation; } catch { }
            if (original == null) return;
            _moveCandidates.Add(obj);
            _moveOriginalLocs.Add(original);
            _moveCandidateTags.Add(tag);
        }

        /// <summary>取父对象：先试 Collection（PS 对象树标准父链），再试 Parent。</summary>
        private static ITxObject GetParentObject(ITxObject obj)
        {
            if (obj == null) return null;
            try
            {
                var p = ((dynamic)obj).Collection as ITxObject;
                if (p != null && !ReferenceEquals(p, obj)) return p;
            }
            catch { }
            try
            {
                var p = ((dynamic)obj).Parent as ITxObject;
                if (p != null && !ReferenceEquals(p, obj)) return p;
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 把机器人移动到 newPos（世界坐标）。
        /// 对当前候选施加平移增量；未验证时"写后读回"检查机器人绝对位置是否真动了，
        /// 没动则升级到上一级容器候选重试。
        /// </summary>
        private bool TryMoveRobotTo(TxVector newPos)
        {
            if (_moveCandidates.Count == 0) return false;

            TxVector before;
            try { before = Robot.AbsoluteLocation.Translation; }
            catch { before = _basePos; }

            double wdx = newPos.X - before.X;
            double wdy = newPos.Y - before.Y;
            double wdz = newPos.Z - before.Z;
            double wantDist = Math.Sqrt(wdx * wdx + wdy * wdy + wdz * wdz);
            if (wantDist < 0.01) return true;   // 已在目标位置

            while (_moveIdx < _moveCandidates.Count)
            {
                var target = _moveCandidates[_moveIdx];
                bool wrote = false;
                try
                {
                    var loc = target.AbsoluteLocation;
                    var t = loc.Translation;
                    loc.Translation = new TxVector(t.X + wdx, t.Y + wdy, t.Z + wdz);
                    target.AbsoluteLocation = loc;
                    wrote = true;
                }
                catch { }

                if (wrote)
                {
                    if (_moveVerified) return true;

                    // 首次验证：机器人绝对位置是否真的动了
                    TxVector after;
                    try { after = Robot.AbsoluteLocation.Translation; }
                    catch { after = before; }
                    double moved = Dist3(after, before);
                    if (moved > Math.Max(1.0, wantDist * 0.2))
                    {
                        _moveVerified = true;   // 该候选有效，锁定
                        return true;
                    }
                }

                // 该候选无效（写失败或写了没动）→ 升级到上一级容器
                _moveIdx++;
            }
            return false;   // 所有候选都失效
        }

        private static double Dist3(TxVector a, TxVector b)
        {
            double dx = a.X - b.X, dy = a.Y - b.Y, dz = a.Z - b.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        // =====================================================================
        //  瞄准点同步（原碰撞盒中心公式保留为数学包围盒）
        // =====================================================================
        private void SyncAimCenter()
        {
            AimCenter = new TxVector(
                _basePos.X, _basePos.Y,
                _basePos.Z + AIM_BOX_SZ / 2.0 + 200);
        }

        /// <summary>AABB 兜底命中判定（干涉集查询不可用时使用）。</summary>
        public bool CheckBulletHit(TxVector bulletPos)
        {
            if (Health <= 0) return false;
            double dx = Math.Abs(bulletPos.X - AimCenter.X);
            double dy = Math.Abs(bulletPos.Y - AimCenter.Y);
            double dz = Math.Abs(bulletPos.Z - AimCenter.Z);
            return dx < AIM_BOX_SX / 2 && dy < AIM_BOX_SY / 2 && dz < AIM_BOX_SZ / 2;
        }

        public void OnKilled()
        {
            try { RestoreOriginalPose(); } catch { }
            MechArenaVisibility.TrySetVisible(Robot, false);
        }

        // =====================================================================
        //  Update
        // =====================================================================
        public void UpdateAttack(double time, TxVector playerPos, bool bossMode)
        {
            if (_jointCount == 0 || Health <= 0) return;

            double dx = playerPos.X - _basePos.X;
            double dy = playerPos.Y - _basePos.Y;
            double dz = playerPos.Z - _basePos.Z;
            double dist = Math.Sqrt(dx * dx + dy * dy + dz * dz);
            double j1Target = Math.Atan2(dy, dx);

            double strikeRange = STRIKE_RANGE * (bossMode ? 1.35 : 1.0);
            double windUpDur = WINDUP_TIME * (bossMode ? 0.7 : 1.0);
            double strikeDur = STRIKE_TIME * (bossMode ? 0.7 : 1.0);
            double recoverDur = RECOVER_TIME * (bossMode ? 0.7 : 1.0);
            double cooldown = COOLDOWN_TIME * (bossMode ? 0.5 : 1.0);

            double sTime = time - _stateStartTime;

            switch (State)
            {
                case AttackState.Track:
                    // 追击 —— 攻击时不动，只有 Track 移动
                    UpdateBodyMovement(playerPos, bossMode);
                    MechArenaRobotHelper.SetJointValues(Robot, ComputeTrackPose(j1Target, time));
                    if (dist < strikeRange && time >= _cooldownUntil)
                        EnterWindUp(time, playerPos);
                    break;

                case AttackState.WindUp:
                    {
                        double t = Math.Min(1.0, sTime / windUpDur);
                        double e = EaseInQuad(t);
                        MechArenaRobotHelper.SetJointValues(Robot,
                            InterpolateJoints(_stateStartJoints, _stateTargetJoints, e));
                        if (t >= 1.0) EnterStrike(time);
                        break;
                    }

                case AttackState.Strike:
                    {
                        double t = Math.Min(1.0, sTime / strikeDur);
                        double e = EaseOutCubic(t);
                        MechArenaRobotHelper.SetJointValues(Robot,
                            InterpolateJoints(_stateStartJoints, _stateTargetJoints, e));
                        if (t >= 1.0) EnterRecover(time);
                        break;
                    }

                case AttackState.Recover:
                    {
                        double t = Math.Min(1.0, sTime / recoverDur);
                        double e = EaseInOutQuad(t);
                        MechArenaRobotHelper.SetJointValues(Robot,
                            InterpolateJoints(_stateStartJoints, _baseJoints, e));
                        if (t >= 1.0)
                        {
                            State = AttackState.Track;
                            _stateStartTime = time;
                            _cooldownUntil = time + cooldown;
                        }
                        break;
                    }
            }
        }

        // =====================================================================
        //  机器人本体移动 —— 追击玩家（Track 状态）
        // =====================================================================
        private void UpdateBodyMovement(TxVector playerPos, bool bossMode)
        {
            double dx = playerPos.X - _basePos.X;
            double dy = playerPos.Y - _basePos.Y;
            double planarDist = Math.Sqrt(dx * dx + dy * dy);

            if (planarDist > PURSUE_RANGE) return;         // 玩家太远，站桩
            if (planarDist < KEEP_DISTANCE) return;        // 太近，交给攻击

            double moveSpeed = MOVE_SPEED * (bossMode ? 1.6 : 1.0);
            double step = Math.Min(moveSpeed, planarDist - KEEP_DISTANCE);
            double vx = dx / planarDist * step;
            double vy = dy / planarDist * step;

            var newPos = new TxVector(_basePos.X + vx, _basePos.Y + vy, _basePos.Z);

            // 通过候选梯队移动（本体失效自动改移父容器）
            if (TryMoveRobotTo(newPos))
            {
                // 以读回的真实位置为准更新 _basePos
                try { _basePos = Robot.AbsoluteLocation.Translation; }
                catch { _basePos = newPos; }
                SyncAimCenter();
            }
        }

        // ---------- 状态转换 ----------
        private void EnterWindUp(double time, TxVector playerPos)
        {
            _lockedTarget = playerPos;
            State = AttackState.WindUp;
            _stateStartTime = time;
            _stateStartJoints = ReadCurrentJoints();
            _stateTargetJoints = ComputeWindUpPose(_lockedTarget);
        }

        private void EnterStrike(double time)
        {
            State = AttackState.Strike;
            _stateStartTime = time;
            _stateStartJoints = ReadCurrentJoints();
            _stateTargetJoints = ComputeStrikePose(_lockedTarget);
        }

        private void EnterRecover(double time)
        {
            State = AttackState.Recover;
            _stateStartTime = time;
            _stateStartJoints = ReadCurrentJoints();
        }

        private double[] ReadCurrentJoints()
        {
            var v = new double[_jointCount];
            Array.Copy(_baseJoints, v, _jointCount);
            try
            {
                // 主路径：强类型 Robot.Joints
                var joints = Robot.Joints;
                if (joints != null)
                {
                    int n = Math.Min(_jointCount, joints.Count);
                    for (int i = 0; i < n; i++)
                    {
                        try { dynamic j = joints[i]; v[i] = (double)j.CurrentValue; } catch { }
                    }
                    return v;
                }
            }
            catch { }
            try
            {
                // 回退：Robot.CurrentPose.JointValues
                ArrayList jv = Robot.CurrentPose.JointValues;
                if (jv != null)
                {
                    int n = Math.Min(_jointCount, jv.Count);
                    for (int i = 0; i < n; i++)
                        v[i] = Convert.ToDouble(jv[i]);
                }
            }
            catch { }
            return v;
        }

        // ---------- 姿态计算 ----------
        private double[] ComputeTrackPose(double j1Target, double time)
        {
            var cur = ReadCurrentJoints();
            var v = new double[_jointCount];
            Array.Copy(_baseJoints, v, _jointCount);

            if (_jointCount > 0)
            {
                double diff = j1Target - cur[0];
                while (diff > Math.PI) diff -= 2 * Math.PI;
                while (diff < -Math.PI) diff += 2 * Math.PI;
                const double maxRate = 0.08;
                v[0] = cur[0] + Math.Sign(diff) * Math.Min(Math.Abs(diff), maxRate);
            }

            for (int i = 1; i < _jointCount; i++)
            {
                v[i] = _baseJoints[i] + Math.Sin(time * 1.2 + PhaseOffset + i * 0.7) * (Math.PI / 24);
            }
            return v;
        }

        private double[] ComputeWindUpPose(TxVector playerPos)
        {
            var above = new TxVector(
                _basePos.X + (playerPos.X - _basePos.X) * 0.5,
                _basePos.Y + (playerPos.Y - _basePos.Y) * 0.5,
                _basePos.Z + 2500);
            var ik = TrySolveIK(above);
            if (ik != null) return ik;

            double dx = playerPos.X - _basePos.X;
            double dy = playerPos.Y - _basePos.Y;
            double j1 = Math.Atan2(dy, dx);
            var v = (double[])_baseJoints.Clone();
            if (_jointCount > 0) v[0] = j1;
            if (_jointCount > 1) v[1] = _baseJoints[1] - Math.PI / 3;
            if (_jointCount > 2) v[2] = _baseJoints[2] + Math.PI / 3;
            if (_jointCount > 4) v[4] = _baseJoints[4] - Math.PI / 4;
            return v;
        }

        private double[] ComputeStrikePose(TxVector playerPos)
        {
            var ik = TrySolveIK(playerPos);
            if (ik != null) return ik;

            double dx = playerPos.X - _basePos.X;
            double dy = playerPos.Y - _basePos.Y;
            double dz = playerPos.Z - _basePos.Z;
            double planar = Math.Sqrt(dx * dx + dy * dy);
            double j1 = Math.Atan2(dy, dx);
            double pitch = Math.Atan2(-dz, planar);
            var v = (double[])_baseJoints.Clone();
            if (_jointCount > 0) v[0] = j1;
            if (_jointCount > 1) v[1] = _baseJoints[1] + Math.PI / 4 + pitch * 0.3;
            if (_jointCount > 2) v[2] = _baseJoints[2] - Math.PI / 3;
            if (_jointCount > 4) v[4] = _baseJoints[4] + Math.PI / 4 + pitch * 0.5;
            return v;
        }

        private double[] TrySolveIK(TxVector targetPos)
        {
            try
            {
                var target = new TxTransformation();
                target.Translation = targetPos;

                var invData = new TxRobotInverseData();
                invData.Destination = target;
                invData.InverseType = TxRobotInverseData.TxInverseType.InverseFullReach;

                ArrayList solutions = Robot.CalcInverseSolutions(invData);
                if (solutions == null || solutions.Count == 0) return null;

                // 取第一个解（参考 IkSolver / CollisionWorld：TxPoseData.JointValues = ArrayList）
                var pose = solutions[0] as TxPoseData;
                if (pose == null) return null;

                ArrayList jv = pose.JointValues;
                if (jv == null || jv.Count == 0) return null;

                var r = new double[_jointCount];
                Array.Copy(_baseJoints, r, _jointCount);
                int lim = Math.Min(_jointCount, jv.Count);
                for (int i = 0; i < lim; i++)
                    r[i] = Convert.ToDouble(jv[i]);
                return r;
            }
            catch { }
            return null;
        }

        // ---------- 工具 ----------
        public TxVector GetAttackTip()
        {
            try
            {
                var tcpf = Robot.TCPF;
                if (tcpf != null) return tcpf.AbsoluteLocation.Translation;
            }
            catch { }
            try { return Robot.AbsoluteLocation.Translation; }
            catch { return _basePos; }
        }

        public void RestoreOriginalPose()
        {
            if (_jointCount == 0) return;
            MechArenaRobotHelper.SetJointValues(Robot, _baseJoints);
        }

        /// <summary>
        /// 恢复所有移动候选（本体 + 各级父容器）到 Init 时缓存的原始位姿。
        /// 没被动过的候选恢复等于无操作，安全。
        /// </summary>
        public void RestoreOriginalLocation()
        {
            int n = Math.Min(_moveCandidates.Count, _moveOriginalLocs.Count);
            for (int i = 0; i < n; i++)
            {
                try { _moveCandidates[i].AbsoluteLocation = _moveOriginalLocs[i]; }
                catch { }
            }
            _basePos = _originalBasePos;
        }

        public void Dispose()
        {
            // 实体碰撞盒已移除，无需清理几何
        }

        private static double[] InterpolateJoints(double[] a, double[] b, double t)
        {
            int n = Math.Min(a.Length, b.Length);
            var r = new double[n];
            for (int i = 0; i < n; i++) r[i] = a[i] + (b[i] - a[i]) * t;
            return r;
        }

        private static double EaseInQuad(double t) => t * t;
        private static double EaseOutCubic(double t) { double u = 1 - t; return 1 - u * u * u; }
        private static double EaseInOutQuad(double t) => t < 0.5 ? 2 * t * t : 1 - Math.Pow(-2 * t + 2, 2) / 2;
    }

    // =========================================================================
    //  Boss 阵型
    // =========================================================================
    public class BossFormation
    {
        public bool Active { get; private set; }
        public readonly List<EnemyRobotAgent> Members = new List<EnemyRobotAgent>();

        public void TryActivate(List<EnemyRobotAgent> pool)
        {
            if (Active) return;
            var alive = pool.Where(x => x.Health > 0).ToList();
            if (alive.Count < 3) return;

            Members.Clear();
            Members.AddRange(alive.Take(3));
            Active = true;

            for (int i = 0; i < Members.Count; i++)
            {
                var m = Members[i];
                m.Health = Math.Max(m.Health, MechArenaEngine.BOSS_MEMBER_HP);
                m.PhaseOffset = i * (Math.PI * 2.0 / 3.0);
            }
        }

        public void UpdateBoss()
        {
            if (!Active) return;
            if (Members.All(m => m.Health <= 0)) Active = false;
        }
    }
}