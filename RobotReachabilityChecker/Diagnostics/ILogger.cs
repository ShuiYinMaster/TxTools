// ============================================================================
// ILogger.cs
//
// 日志抽象接口。Services 层（核心业务逻辑）依赖此接口而非主窗体的 Log 方法，
// 这样：
//   1) 业务逻辑不耦合 UI；
//   2) 单元测试可注入空实现或捕获实现；
//   3) 日志后端可替换（文件、Console、PS Output 窗口等）。
//
// 主窗体 ReachabilityCheckerForm 实现此接口，将日志写入 RichTextBox。
// 服务的方法签名形如：
//     public Foo(..., ILogger log)
// 方便服务在不同上下文（带/不带 UI）下复用。
// ============================================================================
namespace TxTools.RobotReachabilityChecker.Diagnostics
{
    public interface ILogger
    {
        /// <summary>
        /// 写入一条日志。
        /// </summary>
        /// <param name="message">日志正文</param>
        /// <param name="level">"INFO" / "WARN" / "ERR" / "OK" / "DEBUG"</param>
        void Log(string message, string level = "INFO");
    }

    /// <summary>静默实现：丢弃所有日志，主要用于单元测试或不需要日志的场景。</summary>
    public sealed class NullLogger : ILogger
    {
        public static readonly NullLogger Instance = new NullLogger();
        public void Log(string message, string level = "INFO") { /* no-op */ }
    }
}
