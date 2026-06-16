using System;
using System.Collections.Generic;
using System.Threading;
using Tecnomatix.Engineering;

namespace TxTools.AutoRecorder
{
    public enum RecordingState
    {
        Idle,
        Preparing,
        Recording,
        Completed,
        Cancelled,
        Failed,
    }

    public enum LogLevel { Detail, Info, Success, Warn, Error }

    /// <summary>关键帧：到达第 N 个 location 时切换到指定视角</summary>
    public class CameraKeyframe
    {
        public int LocationIndex;   // 1-based：到达第 N 个 location 触发
        public TxCamera Camera;
    }

    /// <summary>一个 operation 的视角脚本</summary>
    public class OperationViewSchedule
    {
        /// <summary>op 开播时套用（对应 LocationIndex=0）。null = 不强制起始视角</summary>
        public TxCamera InitialCamera;
        /// <summary>关键帧列表（推荐按 LocationIndex 升序）</summary>
        public List<CameraKeyframe> Keyframes = new List<CameraKeyframe>();

        public bool IsEmpty
        {
            get { return InitialCamera == null && (Keyframes == null || Keyframes.Count == 0); }
        }
    }

    /// <summary>单个录制任务（一个操作 → 一个视频文件）</summary>
    public class RecordingJob
    {
        public ITxObject Operation;
        public string FilePath;
        public string OperationName;
        /// <summary>该任务的视角脚本。null = 走 ZoomToObjects 兜底；非 null = 起始视角 + 关键帧切换</summary>
        public OperationViewSchedule Schedule;
    }

    /// <summary>批量任务共享的录制参数</summary>
    public class RecordingOptions
    {
        public TxVideoCodec Codec = TxVideoCodec.MPEG4_AVC_H264;
        public int FrameRate = 30;
        public int Compression = 70;
        public int? ResolutionWidth;
        public int? ResolutionHeight;
        public TxMovieTimeSource TimeSource = TxMovieTimeSource.SimulationTime;
        public bool FocusOnOperation = true;
        /// <summary>录像后加速倍数。1.0 = 不处理；>1.0 = 用 ffmpeg 加速</summary>
        public double SpeedupFactor = 1.0;
        /// <summary>ffmpeg 可执行文件路径（null = 自动查找 PATH / 插件目录）</summary>
        public string FfmpegPath;
    }

    /// <summary>
    /// 批量录像编排：依次录制多个 RecordingJob。
    /// player.Ended 回调通过 SyncContext 切回主线程，保证所有 PS SDK 调用都在主线程。
    /// 单个任务失败不中断批次，只记录失败计数。
    /// </summary>
    public class RecordingService
    {
        // ===== 对外事件 =====
        public event Action<RecordingState> StateChanged;
        public event Action<string, LogLevel> Log;
        /// <summary>job 进度：(当前序号 1-based, 总数, 当前操作名)</summary>
        public event Action<int, int, string> JobProgress;

        // ===== 对外状态 =====
        public RecordingState State { get; private set; }
        public int TotalJobs { get; private set; }
        public int CurrentJobIndex { get; private set; }   // 1-based, 0 = 未开始
        public int SucceededCount { get; private set; }
        public int FailedCount { get; private set; }
        public List<string> ProducedFiles { get; private set; }
        public double LastDurationSeconds { get; private set; }
        /// <summary>最近一次成功完成的文件路径（用于"完成后打开"）</summary>
        public string LastFilePath { get; private set; }

        // ===== 内部状态 =====
        private readonly SynchronizationContext _ctx;

        private List<RecordingJob> _jobs;
        private RecordingOptions _options;
        private TxGraphicViewer _viewer;

        private TxSimulationPlayer _player;
        private TxMovieRecorder _recorder;
        private TxViewerRecordingSettings _settings;
        private RecordingJob _currentJob;
        private bool _isCancelling;

        // 当前任务的 Ended 事件是否已处理。
        // 用 Interlocked.CompareExchange 抢占，保证同一任务的多次事件触发只处理一次。
        // 每次启动新任务时重置为 0。
        // 0 = 未处理，1 = 已处理
        private int _jobEndHandled;

        // 当前任务已到达的 destination 计数（关键帧切换索引），原子操作
        private int _destCount;
        // 是否已订阅 OperationReachedDestination
        private bool _hasDestSub;

        public RecordingService() : this(null) { }

        public RecordingService(SynchronizationContext ctx)
        {
            _ctx = ctx ?? SynchronizationContext.Current ?? new SynchronizationContext();
            State = RecordingState.Idle;
            ProducedFiles = new List<string>();
        }

        // ============================================================
        // 主入口：启动批量录制
        // ============================================================
        public void Start(IList<RecordingJob> jobs, RecordingOptions options)
        {
            if (State == RecordingState.Recording || State == RecordingState.Preparing)
            {
                EmitLog("当前已有录制任务进行中", LogLevel.Warn);
                return;
            }
            if (jobs == null || jobs.Count == 0) { Fail("未指定任何录制任务"); return; }
            if (options == null) { Fail("未指定录制参数"); return; }

            // 取主视口
            _viewer = PsReader.GetGraphicViewer();
            if (_viewer == null) { Fail("无法获取主视口"); return; }

            _jobs = new List<RecordingJob>(jobs);
            _options = options;
            TotalJobs = _jobs.Count;
            CurrentJobIndex = 0;
            SucceededCount = 0;
            FailedCount = 0;
            ProducedFiles.Clear();
            LastFilePath = null;
            LastDurationSeconds = 0;
            _isCancelling = false;

            EmitLog(string.Format("批量录制启动，共 {0} 个任务", TotalJobs), LogLevel.Info);
            SetState(RecordingState.Preparing);

            StartNextJob();
        }

        // ============================================================
        // 用户主动取消（终止当前任务，丢弃剩余任务）
        // ============================================================
        public void Cancel()
        {
            if (State != RecordingState.Recording && State != RecordingState.Preparing)
                return;

            _isCancelling = true;
            Interlocked.Exchange(ref _jobEndHandled, 1);  // 让残留 Ended 事件立即失效
            EmitLog("正在取消批量录制...", LogLevel.Warn);

            try { if (_player != null) _player.Stop(); }
            catch (Exception ex) { EmitLog("Stop simulation 异常：" + ex.Message, LogLevel.Warn); }

            try { if (_recorder != null) _recorder.Terminate(); }
            catch (Exception ex) { EmitLog("Terminate recorder 异常：" + ex.Message, LogLevel.Warn); }

            CleanupCurrent();
            SetState(RecordingState.Cancelled);

            int remaining = (_jobs != null) ? _jobs.Count - CurrentJobIndex : 0;
            EmitLog(string.Format("已取消，完成 {0}/{1}（剩余 {2} 个未执行）",
                    SucceededCount, TotalJobs, remaining), LogLevel.Warn);
        }

        // ============================================================
        // 启动下一个任务（或收尾）
        // ============================================================
        private void StartNextJob()
        {
            if (_isCancelling) return;
            if (CurrentJobIndex >= _jobs.Count)
            {
                FinishBatch();
                return;
            }

            CurrentJobIndex++;
            _currentJob = _jobs[CurrentJobIndex - 1];

            EmitLog(string.Format("▶ [{0}/{1}] 开始：{2}",
                    CurrentJobIndex, TotalJobs, _currentJob.OperationName), LogLevel.Info);
            EmitJobProgress(CurrentJobIndex, TotalJobs, _currentJob.OperationName);

            try
            {
                // 1. 应用起始相机 / 聚焦视角
                if (_options.FocusOnOperation)
                {
                    var sched = _currentJob.Schedule;
                    if (sched != null && sched.InitialCamera != null)
                    {
                        EmitLog("  应用起始视角", LogLevel.Detail);
                        PsReader.SetCurrentCamera(_viewer, sched.InitialCamera);
                    }
                    else
                    {
                        var objs = PsReader.CollectOperationObjects(_currentJob.Operation);
                        EmitLog("  收集相关对象 " + objs.Count + " 个，ZoomToSelection 兜底", LogLevel.Detail);
                        PsReader.ZoomToObjects(objs);
                    }
                }

                // 2. 决定分辨率
                uint width, height;
                if (_options.ResolutionWidth.HasValue && _options.ResolutionHeight.HasValue)
                {
                    width = (uint)RoundTo4(_options.ResolutionWidth.Value);
                    height = (uint)RoundTo4(_options.ResolutionHeight.Value);
                }
                else
                {
                    int vw, vh;
                    if (PsReader.TryGetViewerSize(_viewer, out vw, out vh))
                    {
                        width = (uint)RoundTo4(vw);
                        height = (uint)RoundTo4(vh);
                    }
                    else { width = 1920; height = 1080; }
                }

                // 3. 配置录像参数
                _settings = PsReader.CreateViewerSettings(_viewer, _currentJob.FilePath, width, height);
                _settings.Codec = _options.Codec;
                _settings.FrameRate = (uint)_options.FrameRate;
                _settings.Compression = (uint)_options.Compression;
                try { _settings.TimeSource = _options.TimeSource; }
                catch (Exception ex) { EmitLog("  TimeSource 设置失败（已忽略）：" + ex.Message, LogLevel.Warn); }

                // 4. recorder + player
                _recorder = PsReader.CreateMovieRecorder(_viewer);
                _player = new TxSimulationPlayer();
                TrySetOperation(_player, _currentJob.Operation);

                // 只订阅 Ended：自然完成的唯一可靠信号
                // 不订阅 SimulationStopped —— PS 在自然结束时会同时触发两者，制造重复处理；
                // 用户主动取消由 Cancel() 方法独立处理，不依赖该事件
                Interlocked.Exchange(ref _jobEndHandled, 0);
                _player.Ended += OnSimulationEnded;

                // 视角关键帧事件订阅：每到一个 location 计数 +1，找匹配的 keyframe
                Interlocked.Exchange(ref _destCount, 0);
                if (_currentJob.Schedule != null
                    && _currentJob.Schedule.Keyframes != null
                    && _currentJob.Schedule.Keyframes.Count > 0)
                {
                    try
                    {
                        _player.OperationReachedDestination += OnDestinationReached;
                        _hasDestSub = true;
                        EmitLog("  关键帧订阅成功（" + _currentJob.Schedule.Keyframes.Count
                              + " 个切换点）", LogLevel.Detail);
                    }
                    catch (Exception ex)
                    {
                        EmitLog("  关键帧订阅失败：" + ex.Message, LogLevel.Warn);
                        _hasDestSub = false;
                    }
                }

                // 5. 倒带 → 启录 → 播放
                _player.Rewind();
                _recorder.Start(_settings);
                SetState(RecordingState.Recording);
                _player.Play();
            }
            catch (Exception ex)
            {
                FailedCount++;
                EmitLog(string.Format("✗ [{0}/{1}] 启动失败：{2}，跳到下一个",
                        CurrentJobIndex, TotalJobs, ex.Message), LogLevel.Error);
                CleanupCurrent();
                // 用 ctx.Post 避免栈递归
                _ctx.Post(_ => StartNextJob(), null);
            }
        }

        // ============================================================
        // 到达 location —— 关键帧切换触发点
        // 计数 +1 → 在主线程查找匹配关键帧并切换相机
        // ============================================================
        private void OnDestinationReached(object sender, EventArgs e)
        {
            int idx = Interlocked.Increment(ref _destCount);
            _ctx.Post(_ =>
            {
                if (_isCancelling) return;
                if (_currentJob == null || _currentJob.Schedule == null) return;
                var kfs = _currentJob.Schedule.Keyframes;
                if (kfs == null || kfs.Count == 0) return;

                CameraKeyframe match = null;
                for (int i = 0; i < kfs.Count; i++)
                {
                    if (kfs[i] != null && kfs[i].LocationIndex == idx && kfs[i].Camera != null)
                    { match = kfs[i]; break; }
                }
                if (match == null) return;

                try
                {
                    PsReader.SetCurrentCamera(_viewer, match.Camera);
                    EmitLog(string.Format("    ▷ 到达第 {0} 点，切换视角", idx), LogLevel.Detail);
                }
                catch (Exception ex)
                {
                    EmitLog("    切换视角异常：" + ex.Message, LogLevel.Warn);
                }
            }, null);
        }

        // ============================================================
        // 仿真自然结束 —— 抢占 flag 后 Post 到主线程处理
        // 注意：worker 线程不做任何 SDK 调用（含 unsubscribe），避免 PS 事件系统死锁
        // ============================================================
        private void OnSimulationEnded(object sender, EventArgs e)
        {
            // 抢占：只有第一个进入的事件能继续；同一任务的重复触发会被丢弃
            if (Interlocked.CompareExchange(ref _jobEndHandled, 1, 0) != 0) return;
            _ctx.Post(_ => OnSimulationEndedUi(), null);
        }

        private void OnSimulationEndedUi()
        {
            if (_isCancelling) return;
            if (_recorder == null) return;

            try
            {
                double dur = 0;
                try { dur = Convert.ToDouble(_recorder.Duration); } catch { }

                _recorder.Stop();

                LastDurationSeconds = dur;
                LastFilePath = _settings != null ? _settings.FilePath : null;

                // 录后加速（如果配置 > 1.0）
                if (LastFilePath != null && _options.SpeedupFactor > 1.01)
                {
                    var sped = TryApplyFfmpegSpeedup(LastFilePath, _options.SpeedupFactor);
                    if (sped) EmitLog(string.Format("  ▷ 已加速 {0:F2}x", _options.SpeedupFactor),
                                      LogLevel.Detail);
                    // 失败不打断流程，原文件保留
                }

                if (LastFilePath != null) ProducedFiles.Add(LastFilePath);
                SucceededCount++;

                EmitLog(string.Format("✓ [{0}/{1}] 完成（{2:F1}s）：{3}",
                        CurrentJobIndex, TotalJobs, dur, LastFilePath), LogLevel.Success);
            }
            catch (Exception ex)
            {
                try { _recorder.Terminate(); } catch { }
                FailedCount++;
                EmitLog(string.Format("✗ [{0}/{1}] Stop 异常：{2}",
                        CurrentJobIndex, TotalJobs, ex.Message), LogLevel.Error);
            }

            CleanupCurrent();
            StartNextJob();   // 已经在 UI 线程，直接递推
        }

        // ============================================================
        // 批次收尾
        // ============================================================
        private void FinishBatch()
        {
            // 已经全部跑完
            if (FailedCount == 0)
            {
                SetState(RecordingState.Completed);
                EmitLog(string.Format("✓ 批量录制全部完成（成功 {0}/{1}）",
                        SucceededCount, TotalJobs), LogLevel.Success);
            }
            else if (SucceededCount == 0)
            {
                SetState(RecordingState.Failed);
                EmitLog(string.Format("✗ 批量录制全部失败（{0} 个任务均出错）",
                        FailedCount), LogLevel.Error);
            }
            else
            {
                SetState(RecordingState.Completed);
                EmitLog(string.Format("⚠ 批量录制部分完成：成功 {0} / 失败 {1} / 共 {2}",
                        SucceededCount, FailedCount, TotalJobs), LogLevel.Warn);
            }
        }

        // ============================================================
        // 资源清理（当前任务）
        // ============================================================
        private void CleanupCurrent()
        {
            if (_player != null)
            {
                // 在 UI 线程上 unsubscribe（worker 线程做这个可能跟 PS 事件系统冲突）
                try { _player.Ended -= OnSimulationEnded; } catch { }
                if (_hasDestSub)
                {
                    try { _player.OperationReachedDestination -= OnDestinationReached; } catch { }
                    _hasDestSub = false;
                }
                _player = null;
            }
            _recorder = null;
            _settings = null;
            _currentJob = null;
        }

        // ============================================================
        // 工具方法
        // ============================================================
        private static void TrySetOperation(TxSimulationPlayer player, ITxObject op)
        {
            try
            {
                ((dynamic)player).SetOperation((dynamic)op);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "player.SetOperation 失败（操作类型可能不受支持）：" + ex.Message);
            }
        }

        private static int RoundTo4(int v) { return Math.Max(4, (v / 4) * 4); }

        private void SetState(RecordingState s)
        {
            State = s;
            var h = StateChanged;
            if (h != null) { try { h(s); } catch { } }
        }

        private void EmitLog(string msg, LogLevel level)
        {
            var h = Log;
            if (h != null) { try { h(msg, level); } catch { } }
        }

        private void EmitJobProgress(int current, int total, string opName)
        {
            var h = JobProgress;
            if (h != null) { try { h(current, total, opName); } catch { } }
        }

        private void Fail(string msg)
        {
            EmitLog("✗ " + msg, LogLevel.Error);
            SetState(RecordingState.Failed);
        }

        // ============================================================
        // 录后加速 —— 调用 ffmpeg.exe
        // ============================================================
        private bool TryApplyFfmpegSpeedup(string videoPath, double factor)
        {
            try
            {
                if (string.IsNullOrEmpty(videoPath) || !System.IO.File.Exists(videoPath))
                {
                    EmitLog("  加速跳过：源文件不存在", LogLevel.Warn);
                    return false;
                }

                string ffmpeg = LocateFfmpeg(_options.FfmpegPath);
                if (ffmpeg == null)
                {
                    EmitLog("  ⚠ 未找到 ffmpeg.exe（PATH / 插件目录 / 自定义路径都没有），"
                          + "已下载请放到插件目录或加入 PATH。本次录像保留原速。",
                            LogLevel.Warn);
                    return false;
                }

                // ffmpeg 命令：setpts 时间戳压缩 → 视频加速；用 H.264 重编码（无音频）
                string tmpPath = videoPath + ".speedup.tmp.mp4";
                if (System.IO.File.Exists(tmpPath))
                {
                    try { System.IO.File.Delete(tmpPath); } catch { }
                }

                // CRF 23 是 H.264 视觉无损常用档；preset veryfast 平衡速度和压缩
                string args = string.Format(System.Globalization.CultureInfo.InvariantCulture,
                    "-y -i \"{0}\" -filter:v \"setpts=PTS/{1:F4}\" -an "
                  + "-c:v libx264 -preset veryfast -crf 23 \"{2}\"",
                    videoPath, factor, tmpPath);

                EmitLog("  调用 ffmpeg 加速中…", LogLevel.Detail);
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = ffmpeg,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                };
                using (var proc = System.Diagnostics.Process.Start(psi))
                {
                    // 同步等待 —— 加速通常几秒到几十秒，调用方已在 UI 线程
                    // 可在未来改成异步 + 进度回调
                    string stderr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();
                    if (proc.ExitCode != 0)
                    {
                        EmitLog("  ⚠ ffmpeg 退出码 " + proc.ExitCode + "，加速失败。原文件保留。",
                                LogLevel.Warn);
                        if (!string.IsNullOrEmpty(stderr))
                        {
                            // 截取最后 200 字符避免日志爆炸
                            int n = stderr.Length;
                            string tail = n > 200 ? stderr.Substring(n - 200) : stderr;
                            EmitLog("  ffmpeg stderr: " + tail.Replace("\n", " "), LogLevel.Detail);
                        }
                        try { System.IO.File.Delete(tmpPath); } catch { }
                        return false;
                    }
                }

                // 替换原文件
                System.IO.File.Delete(videoPath);
                System.IO.File.Move(tmpPath, videoPath);
                return true;
            }
            catch (Exception ex)
            {
                EmitLog("  ⚠ 加速处理异常：" + ex.Message + "。原文件保留。", LogLevel.Warn);
                return false;
            }
        }

        /// <summary>
        /// 查找 ffmpeg.exe：
        /// 1) 用户自定义路径（RecordingOptions.FfmpegPath）
        /// 2) 插件 DLL 所在目录
        /// 3) 系统 PATH
        /// </summary>
        private static string LocateFfmpeg(string customPath)
        {
            // 1) 自定义路径
            if (!string.IsNullOrEmpty(customPath) && System.IO.File.Exists(customPath))
                return customPath;

            // 2) 插件目录
            try
            {
                string asmDir = System.IO.Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(asmDir))
                {
                    string candidate = System.IO.Path.Combine(asmDir, "ffmpeg.exe");
                    if (System.IO.File.Exists(candidate)) return candidate;
                }
            }
            catch { }

            // 3) PATH
            try
            {
                string path = Environment.GetEnvironmentVariable("PATH") ?? "";
                foreach (var dir in path.Split(System.IO.Path.PathSeparator))
                {
                    if (string.IsNullOrWhiteSpace(dir)) continue;
                    string candidate;
                    try { candidate = System.IO.Path.Combine(dir.Trim(), "ffmpeg.exe"); }
                    catch { continue; }
                    if (System.IO.File.Exists(candidate)) return candidate;
                }
            }
            catch { }

            return null;
        }
    }
}