using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using TxTools.Common;
using LineToSolid;

namespace TxTools.FenceBuilder
{
    /// <summary>
    /// 围栏生成器主窗体。
    /// 风格参考 WeldPointAnnotator: 左侧 270px 控制栏 (GroupBox 卡片 + FlowLayoutPanel + AutoSize),
    /// 右侧主区域 (基线列表 + 操作按钮 + 日志面板)。
    /// 窗体非模态,允许用户在窗口打开时继续在 PS 选择对象。
    /// </summary>
    public class FenceBuilderForm : TxForm
    {
        // 参数控件
        private NumericUpDown _nudMeshWidth, _nudMeshHeight, _nudMeshThk;
        private NumericUpDown _nudFrameWidth, _nudGroundClear;
        private CheckBox _chkMeshTexture;
        private Panel _pnlMeshColor, _pnlFrameColor;

        private NumericUpDown _nudPostWidth, _nudPostGap, _nudPostTopMargin;
        private CheckBox _chkBaseplate;
        private NumericUpDown _nudBpW, _nudBpL, _nudBpT;
        private Panel _pnlPostColor, _pnlBpColor;

        private CheckBox _chkShareCorner;
        private CheckBox _chkIgnoreZ, _chkForceAxis, _chkUseActiveModeling;
        private NumericUpDown _nudMinLastPanel;

        // 颜色状态
        private Color _meshColor = Color.FromArgb(255, 200, 0);
        private Color _frameColor = Color.FromArgb(230, 180, 0);
        private Color _postColor = Color.FromArgb(230, 180, 0);
        private Color _bpColor = Color.FromArgb(200, 160, 0);

        // 基线列表
        private TxObjGridCtrl _baselineGrid;
        private Button _btnAddSel, _btnClearList;

        // 操作按钮
        private Button _btnGenerate, _btnUndo, _btnToggleLog;

        // 日志
        private Panel _logPanel;
        private RichTextBox _logBox;
        private bool _logVisible = true;

        // 撤销记忆
        private List<TxSolid> _lastCreatedSolids = new List<TxSolid>();
        private TxComponent _lastCreatedContainer;

        // 纹理
        private string _texturePath;

        // 字体（统一走 FormUiKit，OS 已 DPI 缩放）
        private readonly Font _baseFont = FormUiKit.BaseFont;
        private readonly Font _boldFont = FormUiKit.BoldFont;
        private static readonly Size _designSize = new Size(1100, 720);
        private bool _dpiApplied;

        public FenceBuilderForm()
        {
            SemiModal = false;
            InitForm();
            BuildUI();
            ResolveTexturePath();


        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            FormUiKit.ApplyDpiScaling(this, ref _dpiApplied, _designSize);
        }

        public override void OnInitTxForm()
        {
            try
            {
                if (_baselineGrid != null)
                {
                    try { _baselineGrid.ListenToPick = true; } catch { }
                    try { _baselineGrid.EnableMultipleSelection = true; } catch { }
                    try { _baselineGrid.EnableRecurringObjects = false; } catch { }
                }
            }
            catch { }
        }

        private void InitForm()
        {
            // 统一窗体规范 + DPI（套件唯一一处缩放设置，详见 FormUiKit）
            FormUiKit.InitStandardForm(this, "围栏生成器 - FenceBuilder",
                _designSize, new Size(960, 600));
        }

        private void BuildUI()
        {
            // ===== 总体骨架: TableLayoutPanel 1行2列 (左侧控制栏固定 320px, 右侧 Fill) =====
            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(4),
                BackColor = SystemColors.Control
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            this.Controls.Add(root);

            // ===== 左侧: 用 FlowLayoutPanel(纵向滚动) 装多张 GroupBox 卡片 =====
            FlowLayoutPanel sidePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(0)
            };
            sidePanel.Controls.Add(BuildMeshCard());
            sidePanel.Controls.Add(BuildPostCard());
            sidePanel.Controls.Add(BuildOptionCard());
            root.Controls.Add(sidePanel, 0, 0);

            // ===== 右侧: 3 行 1 列 TableLayoutPanel (基线 170固定 / 操作 44固定 / 日志 Fill) =====
            TableLayoutPanel rightLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(8, 0, 0, 0)
            };
            rightLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 210));  // 基线
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));   // 操作
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // 日志
            root.Controls.Add(rightLayout, 1, 0);

            // --- 右侧第 1 行: 基线列表卡片 ---
            GroupBox grpBase = new GroupBox
            {
                Text = "基线 (Line / Polyline 特征)",
                Font = _boldFont,
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 20, 8, 8)
            };

            // 基线卡片内部: TableLayoutPanel 1行2列 (Grid Fill / 按钮列 140宽)
            TableLayoutPanel baseInner = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };
            baseInner.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            baseInner.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            baseInner.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            _baselineGrid = new TxObjGridCtrl { Dock = DockStyle.Fill };
            baseInner.Controls.Add(_baselineGrid, 0, 0);

            FlowLayoutPanel baseBtns = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(4, 0, 0, 0)
            };
            _btnAddSel = NewFlatButton("从选择添加", 130);
            _btnAddSel.Click += OnAddFromSelection;
            baseBtns.Controls.Add(_btnAddSel);
            _btnClearList = NewFlatButton("清空列表", 130);
            _btnClearList.Click += OnClearList;
            baseBtns.Controls.Add(_btnClearList);
            baseInner.Controls.Add(baseBtns, 1, 0);

            grpBase.Controls.Add(baseInner);
            rightLayout.Controls.Add(grpBase, 0, 0);

            // --- 右侧第 2 行: 操作按钮条 ---
            FlowLayoutPanel actFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 6, 0, 6)
            };
            _btnGenerate = NewFlatButton("生成围栏", 140, true);
            _btnGenerate.Click += OnGenerate;
            actFlow.Controls.Add(_btnGenerate);

            _btnUndo = NewFlatButton("撤销上次", 120);
            _btnUndo.Click += OnUndoLast;
            actFlow.Controls.Add(_btnUndo);

            _btnToggleLog = NewFlatButton("隐藏日志", 120);
            _btnToggleLog.Click += (s, e) =>
            {
                _logVisible = !_logVisible;
                _logPanel.Visible = _logVisible;
                _btnToggleLog.Text = _logVisible ? "隐藏日志" : "显示日志";
            };
            actFlow.Controls.Add(_btnToggleLog);
            rightLayout.Controls.Add(actFlow, 0, 1);

            // --- 右侧第 3 行: 日志面板 ---
            _logPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle
            };
            _logBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.LightGray,
                ReadOnly = true,
                BorderStyle = BorderStyle.None
            };
            _logPanel.Controls.Add(_logBox);
            rightLayout.Controls.Add(_logPanel, 0, 2);
        }

        // ===================== 卡片构建 =====================

        private GroupBox BuildMeshCard()
        {
            GroupBox g = NewCard("网片参数");
            g.Padding = new Padding(6, 14, 6, 6);
            FlowLayoutPanel f = NewCardFlow();

            AddParamRow(f, "标称网片宽度 (mm)", out _nudMeshWidth, 200, 6000, 1000);
            AddParamRow(f, "网片高度 (mm)", out _nudMeshHeight, 500, 4000, 1800);
            AddParamRow(f, "网片薄板厚度 (mm)", out _nudMeshThk, 1, 20, 2);
            AddParamRow(f, "外框方管截面 (mm)", out _nudFrameWidth, 10, 200, 40);
            AddParamRow(f, "离地间隙 (mm)", out _nudGroundClear, 0, 1000, 150);

            _chkMeshTexture = new CheckBox
            {
                Text = "贴网格纹理(失败自动降级)",
                Checked = true,
                AutoSize = true,
                Font = _baseFont
            };
            f.Controls.Add(_chkMeshTexture);

            // 颜色按钮
            AddColorRow(f, "网片薄板颜色", _meshColor, c => _meshColor = c, out _pnlMeshColor);
            AddColorRow(f, "外框方管颜色", _frameColor, c => _frameColor = c, out _pnlFrameColor);

            g.Controls.Add(f);
            return g;
        }

        private GroupBox BuildPostCard()
        {
            GroupBox g = NewCard("立柱与底座");
            g.Padding = new Padding(6, 14, 6, 6);
            FlowLayoutPanel f = NewCardFlow();

            AddParamRow(f, "立柱截面 (mm)", out _nudPostWidth, 20, 300, 60);
            AddParamRow(f, "立柱-网片间隙 (mm)", out _nudPostGap, 0, 200, 5);
            AddParamRow(f, "立柱顶上边距 (mm)", out _nudPostTopMargin, 0, 500, 100);
            AddColorRow(f, "立柱颜色", _postColor, c => _postColor = c, out _pnlPostColor);

            _chkBaseplate = new CheckBox
            {
                Text = "启用底板法兰",
                Checked = false,
                AutoSize = true,
                Font = _baseFont,
                Margin = new Padding(0, 4, 0, 4)
            };
            _chkBaseplate.CheckedChanged += (s, e) =>
            {
                bool en = _chkBaseplate.Checked;
                _nudBpW.Enabled = en; _nudBpL.Enabled = en; _nudBpT.Enabled = en;
                _pnlBpColor.Enabled = en;
            };
            f.Controls.Add(_chkBaseplate);

            AddParamRow(f, "底板宽 (mm)", out _nudBpW, 50, 500, 150);
            AddParamRow(f, "底板长 (mm)", out _nudBpL, 50, 500, 150);
            AddParamRow(f, "底板厚 (mm)", out _nudBpT, 2, 50, 10);
            AddColorRow(f, "底板颜色", _bpColor, c => _bpColor = c, out _pnlBpColor);
            _nudBpW.Enabled = false; _nudBpL.Enabled = false; _nudBpT.Enabled = false;
            _pnlBpColor.Enabled = false;

            g.Controls.Add(f);
            return g;
        }

        private GroupBox BuildOptionCard()
        {
            GroupBox g = NewCard("通用选项");
            g.Padding = new Padding(6, 14, 6, 6);
            FlowLayoutPanel f = NewCardFlow();

            _chkShareCorner = new CheckBox
            {
                Text = "拐角共享立柱",
                Checked = true,
                AutoSize = true,
                Font = _baseFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            f.Controls.Add(_chkShareCorner);

            _chkIgnoreZ = new CheckBox
            {
                Text = "忽略线段 Z 向(立柱底贴地)",
                Checked = true,
                AutoSize = true,
                Font = _baseFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            f.Controls.Add(_chkIgnoreZ);

            _chkForceAxis = new CheckBox
            {
                Text = "忽略线段夹角(强制横平竖直)",
                Checked = false,
                AutoSize = true,
                Font = _baseFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            f.Controls.Add(_chkForceAxis);

            _chkUseActiveModeling = new CheckBox
            {
                Text = "在当前建模资源下生成",
                Checked = true,
                AutoSize = true,
                Font = _baseFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            f.Controls.Add(_chkUseActiveModeling);

            AddParamRow(f, "末片最小宽度 (mm)", out _nudMinLastPanel, 0, 5000, 500);

            g.Controls.Add(f);
            return g;
        }

        private GroupBox NewCard(string title)
        {
            return new FormUiKit.ColoredGroupBox
            {
                Text = title,
                Font = _boldFont,
                Width = 290,
                MinimumSize = new Size(290, 0),
                Margin = new Padding(0, 0, 0, 8),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowOnly,
                HeaderColor = Color.FromArgb(60, 90, 150),
                ForeColor = Color.FromArgb(60, 90, 150)
            };
        }

        private FlowLayoutPanel NewCardFlow()
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
        }

        private void AddParamRow(FlowLayoutPanel parent, string label,
            out NumericUpDown nud, decimal min, decimal max, decimal val)
        {
            FlowLayoutPanel row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = new Padding(0, 2, 0, 2)
            };

            Label lbl = new Label
            {
                Text = label,
                AutoSize = false,
                Width = 160,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = _baseFont,
                Height = 24
            };
            nud = new NumericUpDown
            {
                Width = 95,
                Height = 24,
                Minimum = min,
                Maximum = max,
                Value = val,
                DecimalPlaces = (max < 100 ? 1 : 0),
                Increment = (max < 100 ? 0.5m : 10),
                Font = _baseFont
            };
            row.Controls.Add(lbl);
            row.Controls.Add(nud);
            parent.Controls.Add(row);
        }

        private void AddColorRow(FlowLayoutPanel parent, string label,
            Color initialColor, Action<Color> onColorChanged, out Panel colorPanel)
        {
            FlowLayoutPanel row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = new Padding(0, 2, 0, 2)
            };

            Label lbl = new Label
            {
                Text = label,
                AutoSize = false,
                Width = 160,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = _baseFont,
                Height = 26
            };

            // 颜色控件用 Panel 而不是 Button: Panel 的 BackColor 变化立即重绘,
            // 不会被 Windows 按钮系统渲染覆盖
            Panel pnl = new Panel
            {
                Width = 95,
                Height = 24,
                BackColor = initialColor,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 1, 0, 1),
                Cursor = Cursors.Hand
            };

            pnl.Click += (s, e) =>
            {
                if (!pnl.Enabled) return;
                using (ColorDialog cd = new ColorDialog
                {
                    Color = pnl.BackColor,
                    FullOpen = true,
                    AnyColor = true
                })
                {
                    if (cd.ShowDialog() == DialogResult.OK)
                    {
                        pnl.BackColor = cd.Color;
                        pnl.Invalidate();
                        pnl.Update();
                        // 通过 callback 同步外部字段(替代之前用 Tag 反向查找)
                        onColorChanged?.Invoke(cd.Color);
                    }
                }
            };

            row.Controls.Add(lbl);
            row.Controls.Add(pnl);
            parent.Controls.Add(row);
            colorPanel = pnl;
        }

        private Button NewFlatButton(string text, int width, bool primary = false)
        {
            FormUiKit.FlatColorButton b = new FormUiKit.FlatColorButton
            {
                Text = text,
                Width = width,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                BgColor = primary ? Color.FromArgb(0, 122, 204) : SystemColors.ControlLight,
                ForeColor = primary ? Color.White : SystemColors.ControlText,
                BorderColor = primary ? Color.FromArgb(0, 122, 204) : SystemColors.ControlDark,
                Font = primary ? _boldFont : _baseFont,
                Margin = new Padding(0, 0, 4, 0)
            };
            return b;
        }

        // ===================== 操作 =====================

        private void OnAddFromSelection(object sender, EventArgs e)
        {
            try
            {
                TxObjectList sel = TxApplication.ActiveSelection.GetItems();
                if (sel == null || sel.Count == 0)
                {
                    Log("[UI] 当前 PS 中没有选中对象");
                    return;
                }

                // 用 LineToSolid 同款 PolylineReader 展开曲线（替代原 FenceBaselineReader）
                var features = new List<ITxObject>();
                foreach (ITxObject obj in sel)
                {
                    if (obj == null) continue;
                    if (PolylineReader.ClassifyFeature(obj) != LineToSolid.FeatureKind.Unknown)
                    {
                        if (!features.Contains(obj)) features.Add(obj);
                    }
                    else
                    {
                        int before = features.Count;
                        PolylineReader.CollectCurvesRecursive(obj, features);
                        Log("[UI] 容器展开，新增 " + (features.Count - before) + " 条曲线");
                    }
                }

                if (features.Count == 0)
                {
                    Log("[UI] WARN 选中对象不是曲线，且容器内未找到任何曲线特征");
                    return;
                }
                Log("[UI] 选中 " + sel.Count + " 项，展开后曲线特征 " + features.Count + " 个");

                int added = 0;
                foreach (ITxObject feat in features)
                {
                    try
                    {
                        _baselineGrid.AppendObject(feat);
                        added++;
                    }
                    catch (Exception ex) { Log("[UI] AppendObject 失败: " + ex.Message); }
                }
                Log("[UI] 已添加 " + added + " 个曲线特征到基线列表");
            }
            catch (Exception ex)
            {
                Log("[UI] ERR 从选择添加失败: " + ex.Message);
            }
        }

        private void OnClearList(object sender, EventArgs e)
        {
            try
            {
                _baselineGrid.Objects = new TxObjectList();
                Log("[UI] 列表已清空");
            }
            catch (Exception ex)
            {
                Log("[UI] ERR 清空列表: " + ex.Message);
            }
        }

        private void OnGenerate(object sender, EventArgs e)
        {
            try
            {
                _btnGenerate.Enabled = false;
                _lastCreatedSolids.Clear();
                _lastCreatedContainer = null;

                List<ITxObject> raw = ReadGridObjects();
                if (raw.Count == 0)
                {
                    Log("[UI] 基线列表为空,无法生成");
                    return;
                }

                // 展开容器(资源/零件等) → 提取全部曲线特征（与 OnAddFromSelection 同款逻辑）
                var baselines = new List<ITxObject>();
                foreach (ITxObject obj in raw)
                {
                    if (PolylineReader.ClassifyFeature(obj) != LineToSolid.FeatureKind.Unknown)
                        baselines.Add(obj);
                    else
                        PolylineReader.CollectCurvesRecursive(obj, baselines);
                }
                Log("[Gen] 网格 " + raw.Count + " 项 → 展开后曲线特征 " + baselines.Count + " 个");

                FenceParameters p = CollectParameters();
                Log("[Gen] 开始: 基线=" + baselines.Count +
                    "  网片=" + p.MeshNominalWidth + "×" + p.MeshHeight +
                    "  立柱=" + p.PostWidth + "  底板=" + p.BaseplateMode);

                int totalSolids = 0;
                foreach (ITxObject obj in baselines)
                {
                    BaseSegmentChain chain = FenceBaselineReader.ReadAsChain(obj, Log);
                    if (chain == null) continue;
                    FenceLayout layout = FenceLayoutPlanner.Plan(chain, p, Log);
                    BuildResult br = FenceGeometryBuilder.Build(layout, p, _texturePath, Log);
                    _lastCreatedSolids.AddRange(br.CreatedSolids);
                    if (br.Container != null) _lastCreatedContainer = br.Container;
                    totalSolids += br.CreatedSolids.Count;
                }
                Log("[Gen] 全部完成,本次共创建 " + totalSolids + " 个 Solid");
            }
            catch (Exception ex)
            {
                Log("[Gen] ERR " + ex.Message + "\r\n" + ex.StackTrace);
            }
            finally
            {
                _btnGenerate.Enabled = true;
            }
        }

        private void OnUndoLast(object sender, EventArgs e)
        {
            // 优先删除整个 Part 容器
            if (_lastCreatedContainer != null)
            {
                try
                {
                    MethodInfo mi = _lastCreatedContainer.GetType().GetMethod("Delete", Type.EmptyTypes);
                    if (mi != null)
                    {
                        mi.Invoke(_lastCreatedContainer, null);
                        Log("[UI] 已删除整个容器 Part");
                        _lastCreatedContainer = null;
                        _lastCreatedSolids.Clear();
                        return;
                    }
                }
                catch (Exception ex) { Log("[UI] 删除容器失败: " + ex.Message); }
            }

            // 退化:用 PS 原生 Undo
            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc != null)
                {
                    PropertyInfo pi = doc.GetType().GetProperty("UndoManager");
                    if (pi != null)
                    {
                        object mgr = pi.GetValue(doc, null);
                        if (mgr != null)
                        {
                            MethodInfo mi = mgr.GetType().GetMethod("Undo", Type.EmptyTypes);
                            if (mi != null) { mi.Invoke(mgr, null); Log("[UI] 已调用 PS Undo"); return; }
                        }
                    }
                }
            }
            catch (Exception ex) { Log("[UI] PS Undo 失败: " + ex.Message); }

            Log("[UI] WARN 撤销未能执行,请手动 Ctrl+Z");
        }

        private List<ITxObject> ReadGridObjects()
        {
            var result = new List<ITxObject>();
            try
            {
                int n = 0;
                try { n = _baselineGrid.Count; }
                catch (Exception ex) { Log("[UI] Count 失败: " + ex.Message); }

                for (int i = 0; i < n; i++)
                {
                    try
                    {
                        ITxObject o = _baselineGrid.GetObject(i) as ITxObject;
                        if (o != null) result.Add(o);
                    }
                    catch (Exception ex) { Log("[UI] GetObject(" + i + ") 失败: " + ex.Message); }
                }
            }
            catch (Exception ex) { Log("[UI] ReadGridObjects 异常: " + ex.Message); }
            return result;
        }

        private FenceParameters CollectParameters()
        {
            return new FenceParameters
            {
                MeshNominalWidth = (double)_nudMeshWidth.Value,
                MeshHeight = (double)_nudMeshHeight.Value,
                MeshThickness = (double)_nudMeshThk.Value,
                MeshFrameWidth = (double)_nudFrameWidth.Value,
                MeshFrameThickness = (double)_nudFrameWidth.Value, // 同宽同厚,正方形截面
                GroundClearance = (double)_nudGroundClear.Value,
                EnableMeshTexture = _chkMeshTexture.Checked,
                MeshColor = _meshColor,
                FrameColor = _frameColor,

                PostWidth = (double)_nudPostWidth.Value,
                PostThickness = (double)_nudPostWidth.Value, // 同宽同厚
                PostGap = (double)_nudPostGap.Value,
                PostTopMargin = (double)_nudPostTopMargin.Value,
                PostColor = _postColor,

                BaseplateMode = _chkBaseplate.Checked ? BaseplateMode.WithBaseplate : BaseplateMode.None,
                BaseplateWidth = (double)_nudBpW.Value,
                BaseplateLength = (double)_nudBpL.Value,
                BaseplateThickness = (double)_nudBpT.Value,
                BaseplateColor = _bpColor,

                ShareCornerPost = _chkShareCorner.Checked,
                IgnoreSegmentZ = _chkIgnoreZ.Checked,
                ForceAxisAligned = _chkForceAxis.Checked,
                CreateUnderActiveModeling = _chkUseActiveModeling.Checked,
                MinLastPanelWidth = (double)_nudMinLastPanel.Value
            };
        }

        private void ResolveTexturePath()
        {
            // 优先级:
            //   1. 嵌入资源(DLL 自带,无需外部部署): 从 manifest stream 读出释放到临时目录
            //   2. DLL 同目录下 Resources\mesh_pattern.png(用户手动部署的)
            //   3. 都没有 → 走降级(半透明纯色)
            try
            {
                // ---- 路径 1: 嵌入资源 ----
                Assembly asm = Assembly.GetExecutingAssembly();
                // 资源名通常是 "<DefaultNamespace>.Resources.mesh_pattern.png"
                // 用反射扫描所有嵌入资源,匹配文件名结尾
                string resName = null;
                foreach (string n in asm.GetManifestResourceNames())
                {
                    if (n.EndsWith("mesh_pattern.png", StringComparison.OrdinalIgnoreCase))
                    {
                        resName = n; break;
                    }
                }
                if (resName != null)
                {
                    string tempPath = Path.Combine(Path.GetTempPath(),
                        "FenceBuilder_mesh_pattern_" + asm.GetName().Version + ".png");
                    if (!File.Exists(tempPath))
                    {
                        using (var stream = asm.GetManifestResourceStream(resName))
                        using (var fs = File.Create(tempPath))
                        {
                            stream.CopyTo(fs);
                        }
                    }
                    _texturePath = tempPath;
                    Log("[Init] 纹理就绪(嵌入资源): " + tempPath);
                    return;
                }

                // ---- 路径 2: DLL 同目录的 Resources\mesh_pattern.png ----
                string dllDir = Path.GetDirectoryName(asm.Location);
                string candidate = Path.Combine(dllDir, "Resources", "mesh_pattern.png");
                if (File.Exists(candidate))
                {
                    _texturePath = candidate;
                    Log("[Init] 纹理就绪(同目录): " + _texturePath);
                    return;
                }

                Log("[Init] WARN 未找到嵌入资源,也未找到 " + candidate + ",将使用降级方案");
            }
            catch (Exception ex) { Log("[Init] 纹理定位异常: " + ex.Message); }
        }

        private void Log(string msg)
        {
            try
            {
                if (_logBox == null) return;
                if (_logBox.InvokeRequired)
                {
                    _logBox.BeginInvoke(new Action<string>(Log), msg);
                    return;
                }
                Color c = Color.LightGray;
                if (msg.Contains("ERR")) c = Color.IndianRed;
                else if (msg.Contains("WARN")) c = Color.Gold;

                int start = _logBox.TextLength;
                string line = "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + Environment.NewLine;
                _logBox.AppendText(line);
                _logBox.Select(start, line.Length);
                _logBox.SelectionColor = c;
                _logBox.Select(_logBox.TextLength, 0);
                _logBox.ScrollToCaret();
            }
            catch { }
        }
    }
}