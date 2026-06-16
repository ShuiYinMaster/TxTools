using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Tecnomatix.Engineering;

namespace TxTools.FenceBuilder
{
    /// <summary>
    /// 给 TxSolid 设置外观:颜色、透明度、纹理贴图(三级降级)。
    /// 所有 API 走反射 + try/catch,以适配 PS 2402 不同小版本。
    /// </summary>
    public static class AppearanceHelper
    {
        // 纹理失败仅记录一次,避免每片网片都刷日志
        private static bool _textureFailureLogged = false;
        public static void ResetTextureLogState() { _textureFailureLogged = false; }

        /// <summary>
        /// 设置纯色外观。返回是否成功。
        /// </summary>
        public static bool TrySetColor(TxSolid solid, Color color, Action<string> log)
        {
            if (solid == null) return false;

            // 路径 1: solid.Color = TxColor
            try
            {
                PropertyInfo p = solid.GetType().GetProperty("Color");
                if (p != null && p.CanWrite)
                {
                    object txColor = TryMakeTxColor(color);
                    if (txColor != null)
                    {
                        p.SetValue(solid, txColor, null);
                        return true;
                    }
                }
            }
            catch (Exception ex) { log?.Invoke("[Appearance] Color 属性设置失败: " + ex.Message); }

            // 路径 2: solid.Appearance.Color = TxColor
            try
            {
                PropertyInfo pa = solid.GetType().GetProperty("Appearance");
                if (pa != null)
                {
                    object app = pa.GetValue(solid, null);
                    if (app != null)
                    {
                        PropertyInfo pc = app.GetType().GetProperty("Color");
                        if (pc != null && pc.CanWrite)
                        {
                            object txColor = TryMakeTxColor(color);
                            if (txColor != null)
                            {
                                pc.SetValue(app, txColor, null);
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { log?.Invoke("[Appearance] Appearance.Color 设置失败: " + ex.Message); }

            return false;
        }

        /// <summary>
        /// 设置半透明度(0..1, 1=完全透明)。返回是否成功。
        /// PS 不同版本对透明度的暴露不一致,做多路径尝试。
        /// </summary>
        public static bool TrySetTransparency(TxSolid solid, double t, Action<string> log)
        {
            if (solid == null) return false;
            string[] candidates = { "Transparency", "Opacity" };
            foreach (string name in candidates)
            {
                try
                {
                    PropertyInfo p = solid.GetType().GetProperty(name);
                    if (p != null && p.CanWrite)
                    {
                        // Opacity 是 1-t
                        double val = (name == "Opacity") ? (1.0 - t) : t;
                        p.SetValue(solid, val, null);
                        return true;
                    }
                }
                catch (Exception ex) { log?.Invoke("[Appearance] " + name + " 失败: " + ex.Message); }

                try
                {
                    PropertyInfo pa = solid.GetType().GetProperty("Appearance");
                    if (pa != null)
                    {
                        object app = pa.GetValue(solid, null);
                        if (app != null)
                        {
                            PropertyInfo pt = app.GetType().GetProperty(name);
                            if (pt != null && pt.CanWrite)
                            {
                                double val = (name == "Opacity") ? (1.0 - t) : t;
                                pt.SetValue(app, val, null);
                                return true;
                            }
                        }
                    }
                }
                catch { }
            }
            return false;
        }

        /// <summary>
        /// 尝试给薄板贴纹理(网格 PNG)。
        /// PS SDK 程序化设纹理覆盖度有限,本方法用反射做最佳努力,失败时返回 false。
        /// 每次失败都记录日志,以便诊断到底哪条路径不通。
        /// </summary>
        public static bool TrySetTexture(TxSolid solid, string texturePath, Action<string> log)
        {
            if (solid == null || string.IsNullOrEmpty(texturePath) || !File.Exists(texturePath))
                return false;

            Type tSolid = solid.GetType();
            object appearance = null;
            try
            {
                PropertyInfo pa = tSolid.GetProperty("Appearance");
                if (pa != null) appearance = pa.GetValue(solid, null);
            }
            catch { }

            // 探测方法 (solid + appearance 都试)
            string[] methodNames = { "SetTexture", "ApplyTexture", "LoadTexture", "SetTextureFromFile" };
            foreach (string mname in methodNames)
            {
                if (TryInvokeStringMethod(solid, mname, texturePath))
                {
                    log?.Invoke("[Appearance] 纹理 OK (Solid." + mname + ")");
                    return true;
                }
                if (appearance != null && TryInvokeStringMethod(appearance, mname, texturePath))
                {
                    log?.Invoke("[Appearance] 纹理 OK (Appearance." + mname + ")");
                    return true;
                }
            }

            // 探测属性 (solid + appearance 都试)
            string[] propNames = { "TexturePath", "Texture", "TextureFile", "TextureFileName", "ImageFile" };
            foreach (string pname in propNames)
            {
                if (TrySetStringProperty(solid, pname, texturePath))
                {
                    log?.Invoke("[Appearance] 纹理 OK (Solid." + pname + ")");
                    return true;
                }
                if (appearance != null && TrySetStringProperty(appearance, pname, texturePath))
                {
                    log?.Invoke("[Appearance] 纹理 OK (Appearance." + pname + ")");
                    return true;
                }
            }

            // 全部失败:首次记录,避免后续每片都刷
            if (!_textureFailureLogged)
            {
                log?.Invoke("[Appearance] 纹理设置失败:TxSolid 上无 SetTexture/Texture 等已知 API,所有网片将降级为半透明纯色(此提示仅出现一次)");
                _textureFailureLogged = true;
            }
            return false;
        }

        private static bool TryInvokeStringMethod(object obj, string methodName, string arg)
        {
            try
            {
                MethodInfo mi = obj.GetType().GetMethod(methodName, new[] { typeof(string) });
                if (mi == null) return false;
                mi.Invoke(obj, new object[] { arg });
                return true;
            }
            catch { return false; }
        }

        private static bool TrySetStringProperty(object obj, string propName, string value)
        {
            try
            {
                PropertyInfo p = obj.GetType().GetProperty(propName);
                if (p == null || !p.CanWrite) return false;
                p.SetValue(obj, value, null);
                return true;
            }
            catch { return false; }
        }

        private static object TryMakeTxColor(Color c)
        {
            // 1) TxColor(byte r, byte g, byte b)
            try
            {
                Type tcType = Type.GetType("Tecnomatix.Engineering.TxColor, Tecnomatix.Engineering");
                if (tcType != null)
                {
                    ConstructorInfo ctor = tcType.GetConstructor(new[] { typeof(byte), typeof(byte), typeof(byte) });
                    if (ctor != null)
                    {
                        return ctor.Invoke(new object[] { c.R, c.G, c.B });
                    }
                    ctor = tcType.GetConstructor(new[] { typeof(int), typeof(int), typeof(int) });
                    if (ctor != null)
                    {
                        return ctor.Invoke(new object[] { (int)c.R, (int)c.G, (int)c.B });
                    }
                }
            }
            catch { }

            // 2) System.Drawing.Color 直接传(某些版本接受)
            return c;
        }
    }
}