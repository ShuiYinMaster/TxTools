using Tecnomatix.Engineering;
using TxTools.WeldGunDefiner.UI;

namespace TxTools.WeldGunDefiner
{
    public class WeldGunDefinerCommand : TxButtonCommand
    {
        public override string Name => ".快速创建 / X枪运动学";
        public override string Category => "TxTools";
        public override string Description => "快速创建X枪运动学";
        public override string LargeBitmap => "image.Xgun.png";

        public override void Execute(object param)
        {
            WeldGunWizardForm.ShowSingleton();
        }
    }
}
