// AllocatorCommand.cs  —  C# 7.3
// 焊点分配插件入口。TxButtonCommand 真抽象成员仅 Name / Category / Execute(object)。

using System;
using System.Threading;
using System.Windows.Forms;
using Tecnomatix.Engineering;

namespace MyPlugin.WeldSpotAllocator
{
    public class AllocatorCommand : TxButtonCommand
    {
        public override string Name => "焊点分配 / 更新";
        public override string Category => "TxTools";
        public override string Description => "打开焊点分配界面，根据已有焊点来分配新的焊点";
        public override string Bitmap => "image.WeldSpotAllocator.png";
        public override string LargeBitmap { get { return "WeldSpotAllocator.png"; } }

        public override void Execute(object cmdParams)
        {
            try
            {
                // 捕获 PS 主线程上下文（如后续要把耗时任务放后台、SDK 调用 Send 回主线程）
                var ctx = SynchronizationContext.Current;
                AllocatorForm.ShowSingleton(ctx);
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动失败：" + ex.Message, "焊点分配", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
