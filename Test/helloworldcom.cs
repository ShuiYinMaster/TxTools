using System.Collections.Generic;
using Tecnomatix.Engineering;

namespace TxTools.HelloMulti
{
    public class HelloMultiCmd : TxButtonCommand
    {
        private static readonly List<HelloTxForm> _open = new List<HelloTxForm>();
        private static int _counter = 0;

        public override string Name => "HelloMultiTest";
        public override string Category => "TxTools";

        public override void Execute(object cmdParams)
        {
            _counter++;
            var f = new HelloTxForm("Hello World " + _counter, _counter);
            _open.Add(f);
            f.FormClosed += (s, e) => _open.Remove(f);
            f.Show();   // 非半模态的 TxForm,等同普通 modeless,不锁命令
        }
    }
}