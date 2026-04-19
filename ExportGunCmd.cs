// ExportGunCmd.cs  —  C# 7.3
//
// [重构适配] ExportGunForm 现在继承 TxForm，构造函数只需 SynchronizationContext。
// TxForm 由 PS 框架托管，自动处理窗口层级，无需 P/Invoke 或独立 STA 线程。
//
// 关键改动：
//   - ExportGunForm 构造函数：(SynchronizationContext psCtx)，移除了 IntPtr psHwnd
//   - TxForm 必须在 PS 主线程上创建和显示，不能放到独立 STA 线程
//   - 移除 Application.Run() 和独立线程逻辑
//   - 使用 form.Show() 非模态显示，不阻塞 PS 主线程

using System;
using System.Threading;
using Tecnomatix.Engineering;

namespace MyPlugin.ExportGun
{
    public class ExportGunCmd : TxButtonCommand
    {
        public override string Name { get { return "MyPlugin.ExportGun"; } }
        public override string Category { get { return "My Plugins"; } }
        public override string Tooltip { get { return "导出插枪和点云到 Catia"; } }
        public override string Description { get { return "将插枪及焊点云数据导出到 Catia"; } }

        public override void Execute(object cmdParams)
        {
            // 在 PS 主线程捕获 SynchronizationContext
            SynchronizationContext psCtx = SynchronizationContext.Current
                                          ?? new SynchronizationContext();

            // TxForm 在 PS 主线程上创建和显示
            // 不使用独立 STA 线程，TxForm 由 PS 框架托管窗口层级
            ExportGunForm form = new ExportGunForm(psCtx);
            form.Show();  // 非模态，不阻塞 PS

            try { TxApplication.StatusBarMessage = "导出插枪插件已启动"; } catch { }
        }
    }
}