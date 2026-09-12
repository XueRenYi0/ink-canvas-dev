using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：截图与图片功能（自绘遮罩框选版）。
    ///
    /// 【为什么弃用系统截图（ms-screenclip:）】
    /// 系统截图是"黑盒"：没有完成/取消回调，只能靠监听剪贴板+超时去猜结果——
    /// 用户半途取消、期间复制了别的图，都会造成误判（莫名插入旧图/界面卡在隐藏状态）。
    ///
    /// 【本方案（方案 B）】
    /// 自己弹全屏半透明遮罩窗口（ScreenshotMaskWindow），用户在上面拖框选区：
    /// - 完成 = 松手（选区够大）→ 先关遮罩再 CopyFromScreen 抓屏 → 图落画布左上角 + 进剪贴板
    /// - 取消 = Esc / 右键 / 单击没拖开 / 窗口失焦 → 关窗即恢复原状，瞬时、无副作用、无等待
    /// 整个状态机闭环在遮罩窗口里，不存在"等外部程序结果"的悬空状态。
    ///
    /// 【两个截图入口的差异 = 一个布尔开关】
    /// - 快速截图：进遮罩前不藏 UI（用户在遮罩下看到屏幕原样，可框选桌面/PPT 等任何内容）；
    ///   选区确定的一瞬间（遮罩还盖着屏幕时）才同步藏 UI，保证最终截图里没有悬浮条。
    /// - 隐藏界面截图：进遮罩前先藏 UI 并等渲染完成，遮罩下只剩板书内容，
    ///   用于截取被悬浮条/面板挡住的画面。
    ///
    /// 模块自包含：所有入口 try/catch，异常只通知不抛出；UI 隐藏必配恢复（try/finally）。
    /// </summary>
    public partial class MainWindow
    {
        #region 截图按钮弹出菜单（悬浮条相机图标）

        /// <summary>相机图标：开/关截图功能弹出菜单（3 项，风格同清屏确认气泡）</summary>
        private void SymbolIconScreenshot_MenuToggle(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            ImageLayer_MenuSetVisible(BorderImageMenu.Visibility != Visibility.Visible);
        }

        /// <summary>菜单开关统一入口（true=显示，false=隐藏）</summary>
        private void ImageLayer_MenuSetVisible(bool visible)
        {
            BorderImageMenu.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        }

        // 「点菜单外自动关菜单」已并入 MW_PopupLayers.cs 的统一登记表（Window_PreviewMouseDown）。
        // 原实现是本文件一个独立方法 + MW_Init.cs 里 `PreviewMouseDown += ImageLayer_MenuCloseOnOutsideClick`
        // 订阅，与笔/选择/橡皮那三段 if 各写一遍、行为不一致；2026-09-11 合并。
        // 新增弹层请改那张表，不要再恢复独立订阅。

        /// <summary>判断元素是否是 target 的子孙（沿视觉树向上找）。
        /// 注：弹层"点外收起"用的同类判断已迁至 MW_PopupLayers.cs 的 IsVisualDescendantOf
        /// （多一层 ContentElement 兼容）；此处保留给粘贴气泡 / 墨迹层判断使用。</summary>
        private static bool IsDescendantOf(DependencyObject node, DependencyObject target)
        {
            while (node != null)
            {
                if (ReferenceEquals(node, target)) return true;
                node = System.Windows.Media.VisualTreeHelper.GetParent(node);
            }
            return false;
        }

        /// <summary>菜单项：快速截图（遮罩框选屏幕任意内容，悬浮条不会入镜）</summary>
        private void BtnMenuShotRegion_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false);
            Screenshot_Start(hideUiFirst: false);
        }

        /// <summary>菜单项：隐藏界面截图（先藏起本软件 UI 再框选，截被悬浮条/面板挡住的内容）</summary>
        private void BtnMenuShotHidden_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false);
            Screenshot_Start(hideUiFirst: true);
        }

        /// <summary>菜单项：上传本地图片（文件选择框 → 落画布左上角，多选时梯级错位）</summary>
        private void BtnMenuImportImage_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false); // 先关菜单再弹文件对话框，避免两个浮层叠着
            try
            {
                var dlg = new OpenFileDialog
                {
                    Title = "选择图片（将插入白板当前页）",
                    Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件|*.*",
                    Multiselect = true // 支持一次选多张，逐张插入
                };
                if (dlg.ShowDialog() != true) return;

                int count = 0;
                foreach (var file in dlg.FileNames)
                {
                    try
                    {
                        var source = new BitmapImage();
                        source.BeginInit();
                        source.CacheOption = BitmapCacheOption.OnLoad; // 立即读入内存，文件可立即被覆盖/删除
                        source.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                        source.UriSource = new Uri(file, UriKind.Absolute);
                        source.EndInit();

                        // cascade=true：连续插入的多张图右下错开摆放，避免完全叠死
                        if (ImageLayer_AddImage(source, cascade: true) != null) count++;
                    }
                    catch { /* 单张失败跳过，继续插后面的 */ }
                }
                if (count > 0)
                    ShowNotification(currentMode == 0
                        ? "已插入 " + count + " 张图片到批注层（第 0 页）"
                        : "已插入 " + count + " 张图片到白板第 " + CurrentWhiteboardIndex + " 页");
            }
            catch (Exception ex)
            {
                ShowNotification("打开图片失败：" + ex.Message);
            }
        }

        #endregion

        #region 剪贴板粘贴贯通（模态粘贴：快速截图 / Ctrl+V / 菜单粘贴 / 粘贴气泡）

        /// <summary>快速截图暂存的图片：非空 = 粘贴模式待放置（点白板弹"粘贴气泡"，一次性）。
        /// 翻页不清空——这正是"截图放到别的页"的核心诉求：截完图翻到目标页再点击白板。</summary>
        private BitmapSource _pendingPasteImage = null;

        /// <summary>待粘贴墨迹（ISF 来源，菜单"粘贴图片"时从剪贴板读出缓存）：
        /// 非 null = 粘贴落下的是墨迹（可编辑），优先于 _pendingPasteImage</summary>
        private StrokeCollection _pendingPasteStrokes = null;

        /// <summary>墨迹连续粘贴（Ctrl+V）级联步进：每贴一次右下错开 24px，5 档循环不叠死</summary>
        private int _pasteCascadeStep = -1;

        /// <summary>粘贴气泡要用的插入位置（= 用户点击白板的位置，图片中心对准该点）</summary>
        private Point _pendingPasteCenter;

        /// <summary>Window.PreviewMouseDown 记录的按下位置（区分"单击"与"拖动"用）</summary>
        private Point _pasteClickDownPos;

        /// <summary>粘贴模式激活中：画布挂起（EditingMode=None 禁书写/擦除/选择），
        /// 粘贴专用光标，点击画布=选位置。粘贴/取消/Esc/切工具后退出。</summary>
        private bool isPasteMode = false;

        /// <summary>进入粘贴模式前的编辑模式（退出时恢复——用户要求"退回上一个状态和光标"）</summary>
        private InkCanvasEditingMode _savedEditModeForPaste = InkCanvasEditingMode.Ink;

        /// <summary>粘贴专用光标（懒创建缓存：图片图标 + 右下角加号徽章，白色填充黑色描边）</summary>
        private Cursor _pasteCursor = null;

        /// <summary>粘贴模式的钩子只装一次（QueryCursor 覆盖 + EditingMode 监视）</summary>
        private bool _pasteModeHooksInstalled = false;

        /// <summary>进入粘贴模式（快速截图完成 / 菜单粘贴共用）：挂起画布 + 专用光标。
        /// 已在粘贴模式中（连续截图）不重复保存原模式。</summary>
        private void EnterPasteMode()
        {
            EnsurePasteModeHooks();
            if (isPasteMode) return; // 已激活（如连续两次快速截图）

            isPasteMode = true;
            _savedEditModeForPaste = inkCanvas.EditingMode; // 记住来路
            HidePastePrompt(); // 清掉可能残留的旧气泡
            // 画布挂起：None 模式下书写/擦除/选择全部失效——点画布只用来选粘贴位置，
            // 修复"点击选位置时会落墨点"的 bug（笔模式下点击=落笔）
            try { inkCanvas.EditingMode = InkCanvasEditingMode.None; } catch { }
        }

        /// <summary>退出粘贴模式：清状态 + 恢复原模式（restoreMode=false 时保留当前模式——
        /// 用户主动切了别的工具，以新工具为准，不能把人家刚点的笔模式顶回去）。</summary>
        private void ExitPasteMode(bool restoreMode = true)
        {
            if (!isPasteMode) return;
            isPasteMode = false;
            _pendingPasteImage = null;
            _pendingPasteStrokes = null;
            HidePastePrompt();

            if (restoreMode)
            {
                // None→原模式 的切换会触发 InkCanvas 内部光标机制重新上光标（笔形/橡皮光标自动恢复）
                try { inkCanvas.EditingMode = _savedEditModeForPaste; } catch { }
            }
        }

        /// <summary>
        /// 装一次粘贴模式钩子：
        /// 1. QueryCursor（handledEventsToo=true 压过 InkCanvas 内部光标逻辑）——
        ///    粘贴模式下画布上显示专用光标；退出后不再拦截，InkCanvas 自动接管（笔形光标等）。
        ///    【不用 Cursor 属性】直接改属性会破坏 InkCanvas 的动态光标机制（历史教训）。
        /// 2. EditingModeChanged——粘贴模式中被切到其他工具（点笔/橡皮/选择图标）→ 放弃粘贴状态。
        ///    注意：切白板页不碰 EditingMode，翻页粘贴的核心流程不受影响。
        /// </summary>
        private void EnsurePasteModeHooks()
        {
            if (_pasteModeHooksInstalled) return;
            _pasteModeHooksInstalled = true;
            try
            {
                inkCanvas.AddHandler(UIElement.QueryCursorEvent,
                    new QueryCursorEventHandler((s, e) =>
                    {
                        if (isPasteMode)
                        {
                            e.Cursor = PasteCursor;
                            e.Handled = true;
                        }
                    }), true); // handledEventsToo：必须压过 InkCanvas 的类级光标处理器
                inkCanvas.EditingModeChanged += (s, e) =>
                {
                    if (isPasteMode && inkCanvas.EditingMode != InkCanvasEditingMode.None)
                        ExitPasteMode(restoreMode: false); // 用户切了别的工具：保留新工具，只放弃粘贴状态
                };
            }
            catch { }
        }

        /// <summary>粘贴专用光标（懒创建）：32×32 位图 → .cur 流 → Cursor；失败兜底十字光标</summary>
        private Cursor PasteCursor
        {
            get
            {
                if (_pasteCursor == null)
                {
                    try { _pasteCursor = BuildPasteCursor(); } catch { }
                    if (_pasteCursor == null) _pasteCursor = Cursors.Cross;
                }
                return _pasteCursor;
            }
        }

        /// <summary>画粘贴光标：图片图标（圆角框+山+太阳）+ 右下角加号徽章——白底黑描边，
        /// 黑白背景上都看得清。热点在图片框中心（放置图片时图片中心对准该点）。</summary>
        private static Cursor BuildPasteCursor()
        {
            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                var outline = new Pen(Brushes.Black, 1.5) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round };
                // 图片框（5,5)-(23,23)：白底圆角矩形
                dc.DrawRoundedRectangle(Brushes.White, outline, new Rect(5, 5, 18, 18), 3, 3);
                // 山形（黑色填充）
                var geo = new StreamGeometry();
                using (var ctx = geo.Open())
                {
                    ctx.BeginFigure(new Point(8, 20), true, true);
                    ctx.LineTo(new Point(12.5, 11.5), true, false);
                    ctx.LineTo(new Point(15.5, 15.5), true, false);
                    ctx.LineTo(new Point(17.5, 13), true, false);
                    ctx.LineTo(new Point(20, 20), true, false);
                }
                dc.DrawGeometry(Brushes.Black, null, geo);
                // 太阳点
                dc.DrawEllipse(Brushes.Black, null, new Point(10, 9.5), 1.7, 1.7);
                // 右下角加号徽章（白底黑描边圆 + 黑加号）
                dc.DrawEllipse(Brushes.White, outline, new Point(23.5, 23.5), 6.5, 6.5);
                var plus = new Pen(Brushes.Black, 2.0) { StartLineCap = PenLineCap.Round, EndLineCap = PenLineCap.Round };
                dc.DrawLine(plus, new Point(23.5, 20.3), new Point(23.5, 26.7));
                dc.DrawLine(plus, new Point(20.3, 23.5), new Point(26.7, 23.5));
            }
            var rtb = new RenderTargetBitmap(32, 32, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            rtb.Render(dv);
            return CreateCursorFromBitmap(rtb, 14, 14); // 热点=图片框中心
        }

        /// <summary>位图 → WPF Cursor：手工拼 .cur 文件流（ICONDIR + ICONDIRENTRY[热点]
        /// + BITMAPINFOHEADER + 自下而上 BGRA 像素 + AND 掩码），走 Cursor(Stream) 解析。
        /// 32bpp 带 alpha，任何背景下都有平滑边缘。</summary>
        private static Cursor CreateCursorFromBitmap(BitmapSource source, int hotspotX, int hotspotY)
        {
            int w = source.PixelWidth, h = source.PixelHeight;
            // Pbgra32（预乘）→ Bgra32（直通 alpha，.cur DIB 要求）
            var bmp = new FormatConvertedBitmap(source, System.Windows.Media.PixelFormats.Bgra32, null, 0);
            var pixels = new byte[w * h * 4];
            bmp.CopyPixels(pixels, w * 4, 0);

            int maskRow = ((w + 31) / 32) * 4;      // AND 掩码每行字节数（1bpp 按 32 位对齐）
            int maskSize = maskRow * h;
            int xorSize = w * h * 4;                 // 彩色数据大小
            int dibSize = 40 + xorSize + maskSize;   // 头 + 彩色 + 掩码

            using (var ms = new MemoryStream())
            using (var bw = new BinaryWriter(ms))
            {
                // ICONDIR（type=2 表示光标）
                bw.Write((ushort)0); bw.Write((ushort)2); bw.Write((ushort)1);
                // ICONDIRENTRY（光标这两字段 = 热点坐标）
                bw.Write((byte)w); bw.Write((byte)h);
                bw.Write((byte)0); bw.Write((byte)0);
                bw.Write((ushort)hotspotX); bw.Write((ushort)hotspotY);
                bw.Write((uint)dibSize); bw.Write((uint)22); // 数据大小 + 偏移（头固定 22 字节）
                // BITMAPINFOHEADER（高度 = XOR+AND 双倍，DIB 行序自下而上）
                bw.Write((uint)40); bw.Write((int)w); bw.Write((int)(h * 2));
                bw.Write((ushort)1); bw.Write((ushort)32); bw.Write((uint)0);
                bw.Write((uint)(xorSize + maskSize));
                bw.Write((int)0); bw.Write((int)0); bw.Write((uint)0); bw.Write((uint)0);
                // 彩色数据：自下而上每像素 BGRA
                for (int y = h - 1; y >= 0; y--)
                    for (int x = 0; x < w; x++)
                    {
                        int i = (y * w + x) * 4;
                        bw.Write(pixels[i]);     // B
                        bw.Write(pixels[i + 1]); // G
                        bw.Write(pixels[i + 2]); // R
                        bw.Write(pixels[i + 3]); // A
                    }
                // AND 掩码全 0（透明交给 alpha 通道）
                for (int i = 0; i < maskSize; i++) bw.Write((byte)0);
                bw.Flush();
                ms.Position = 0;
                return new Cursor(ms);
            }
        }

        /// <summary>菜单项：粘贴图片——进入粘贴模式（与快速截图同款）：
        /// 画布挂起 + 专用光标，点白板选位置 → 弹气泡（粘贴到此处/取消），落点即粘贴位置。
        /// 剪贴板是纯墨迹（ISF）时落下的是墨迹（可编辑），否则是图片。</summary>
        private void BtnMenuPasteImage_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false);
            try
            {
                // 优先墨迹（纯墨迹复制 → 粘贴仍是墨迹），其次图片（截图/含图片的复制）
                var strokes = TryGetClipboardStrokes();
                if (strokes != null)
                {
                    _pendingPasteStrokes = strokes;
                    EnterPasteMode();
                    ShowNotification("点击白板选择墨迹粘贴位置（Esc 取消）");
                    return;
                }
                var pending = TryGetClipboardImage();
                if (pending == null)
                {
                    ShowNotification("剪贴板里没有可粘贴的图片或墨迹");
                    return;
                }
                _pendingPasteImage = pending;
                EnterPasteMode();
                ShowNotification("点击白板或桌面选择粘贴位置（Esc 取消）");
            }
            catch (Exception ex)
            {
                ShowNotification("读取剪贴板失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 从剪贴板插入白板当前页（Ctrl+V / 菜单粘贴 / 气泡确认共用）。
        /// 按剪贴板内容分流：优先墨迹（ISF——纯墨迹复制 → 粘贴仍是墨迹，可编辑），
        /// 其次图片（截图/含图片的复制 → 渲染好的 PNG）。
        /// center 有值 = 内容中心对准该点（粘贴气泡场景）；无值 = 级联错位（连续粘贴不完全叠死）。
        /// </summary>
        private void PasteClipboardImageToCanvas(Point? center)
        {
            try
            {
                // 优先墨迹（ISF）：纯墨迹复制 → 粘贴仍是墨迹（可编辑/可擦/可撤销）
                var strokes = TryGetClipboardStrokes();
                if (strokes != null)
                {
                    PasteStrokesToCanvas(strokes, center);
                    return;
                }
                var img = TryGetClipboardImage();
                if (img == null)
                {
                    ShowNotification("剪贴板里没有可粘贴的图片或墨迹");
                    return;
                }
                if (ImageLayer_AddImage(img, cascade: !center.HasValue, center: center) != null)
                    ShowNotification(currentMode == 0
                        ? "已粘贴到批注层（第 0 页）"
                        : "已粘贴到白板第 " + CurrentWhiteboardIndex + " 页");
            }
            catch (Exception ex)
            {
                ShowNotification("粘贴失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 把墨迹粘贴到画布当前视图（纯墨迹复制的落地动作：粘出来还是墨迹——
        /// 可编辑、可擦除、走撤销栈；落点页与图片同规则：白板模式=当前白板页，注释模式=批注层）。
        /// center 有值 = 墨迹包围盒中心对准该点（粘贴气泡场景，钳制在画布内）；
        /// 无值 = 级联错位：以原件位置为基准每贴一次右下错开 24px，5 档循环（Ctrl+V 连续粘贴不叠死）。
        /// </summary>
        private void PasteStrokesToCanvas(StrokeCollection strokes, Point? center)
        {
            try
            {
                if (strokes == null || strokes.Count == 0) return;
                var bounds = strokes.GetBounds();
                if (bounds.IsEmpty || bounds.Width < 0 || bounds.Height < 0) return;

                double dx, dy;
                if (center.HasValue)
                {
                    // 中心对准点击点，整体钳制在画布内（与图片粘贴同款规则）
                    double hostW = inkCanvas.ActualWidth, hostH = inkCanvas.ActualHeight;
                    if (hostW <= 0) hostW = SystemParameters.WorkArea.Width;
                    if (hostH <= 0) hostH = SystemParameters.WorkArea.Height;
                    double x = Math.Max(4, Math.Min(center.Value.X - bounds.Width / 2,
                        hostW - bounds.Width - 4));
                    double y = Math.Max(4, Math.Min(center.Value.Y - bounds.Height / 2,
                        hostH - bounds.Height - 4));
                    dx = x - bounds.X;
                    dy = y - bounds.Y;
                }
                else
                {
                    // 级联错位：第 1 贴偏 24px、第 2 贴 48px……5 档循环（不与原件完全重合）
                    _pasteCascadeStep = (_pasteCascadeStep + 1) % 5;
                    double off = 24 * (_pasteCascadeStep + 1);
                    dx = off;
                    dy = off;
                }

                // 平移笔迹到目标位置（纯平移不改笔画粗细——applyToStylusTip=false）
                if (dx != 0 || dy != 0)
                {
                    var m = new Matrix();
                    m.Translate(dx, dy);
                    foreach (Stroke s in strokes) s.Transform(m, false);
                }

                // 加入当前页墨迹集：StrokesChanged 自动进 TimeMachine（Ctrl+Z 可撤销）
                inkCanvas.Strokes.Add(strokes);
                ShowNotification(currentMode == 0
                    ? "已粘贴墨迹到批注层（第 0 页）"
                    : "已粘贴墨迹到白板第 " + CurrentWhiteboardIndex + " 页");
            }
            catch (Exception ex)
            {
                ShowNotification("粘贴墨迹失败：" + ex.Message);
            }
        }

        /// <summary>
        /// Window.PreviewMouseUp：粘贴模式下的画布单击 → 弹"粘贴气泡"。
        /// 判定条件（全满足才弹）：
        /// 1. 有待粘贴内容（图片或墨迹：快速截图/菜单粘贴刚完成，一次性）
        /// 2. 左键 + 按下→抬起位移 ≤ 6px（拖动不算单击）
        /// 3. 点击命中在画布上（inkCanvas 铺满主网格；点悬浮条/面板/菜单上不弹）
        /// 气泡显示后：点画布其他位置=移动气泡（换位置），点气泡外其他区域=关闭并退出粘贴模式。
        /// </summary>
        internal void Window_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            // ★ "点面板外收起面板"的那次单击不该画出点：先做单击/书写判定。
            //   必须放在最前面 —— 隧道抬起事件早于 InkCanvas 提交笔迹，
            //   在这里把提交类型恢复成 UserInput，才能让"书写"那一次照常进撤销栈。
            //   状态机与原因见 MW_PopupLayers.cs 的 DismissPressDown / DismissPressUp。
            DismissPressUp(e.GetPosition(Main_Grid));

            // 气泡已显示：点在气泡外 → 换位置（点画布）或关闭退出（点工具栏/面板等其他区域）
            if (BorderPastePrompt.Visibility == Visibility.Visible)
            {
                if (e.OriginalSource is DependencyObject d && !IsDescendantOf(d, BorderPastePrompt))
                {
                    if (IsDescendantOf(d, inkCanvas))
                        ShowPastePrompt(e.GetPosition(Main_Grid)); // 画布上换位置（继续粘贴模式）
                    else
                        ExitPasteMode(); // 其他区域：关气泡 + 退出粘贴模式（恢复原工具）
                }
                return;
            }

            // 图片或墨迹任一待粘贴（纯墨迹复制 → _pendingPasteStrokes）
            if (_pendingPasteImage == null && _pendingPasteStrokes == null) return;
            if (e.ChangedButton != MouseButton.Left) return;
            if (!(e.OriginalSource is DependencyObject dep)) return;

            // 单击判定：位移小（拖出去几厘米的不算选位置）
            var upPos = e.GetPosition(Main_Grid);
            if ((upPos - _pasteClickDownPos).Length > 6) return;

            // 点在画布上才弹（点悬浮条/面板/选区操作条上不弹——它们有自己的点击行为）
            if (!IsDescendantOf(dep, inkCanvas)) return;

            ShowPastePrompt(upPos);
        }

        /// <summary>显示/移动粘贴气泡（悬在点击点上方；顶部放不下改下方），记录粘贴中心位置</summary>
        private void ShowPastePrompt(Point clickPos)
        {
            _pendingPasteCenter = clickPos;
            _pendingPasteImage = null;    // 一次性：弹过就清（取消后想再粘贴走 Ctrl+V/菜单）
            _pendingPasteStrokes = null;  // 同上——真正粘贴时按剪贴板当前内容分流（墨迹/图片）

            BorderPastePrompt.Visibility = Visibility.Visible;
            BorderPastePrompt.UpdateLayout(); // 强制布局算出 ActualWidth 再定位

            double w = BorderPastePrompt.ActualWidth, h = BorderPastePrompt.ActualHeight;
            double x = clickPos.X - w / 2;
            double y = clickPos.Y - h - 12; // 悬在点击点上方（手指/光标不挡）
            if (y < 4) y = clickPos.Y + 12; // 顶到屏幕边了改弹下方
            x = Math.Max(4, Math.Min(x, Main_Grid.ActualWidth - w - 4));
            y = Math.Max(4, Math.Min(y, Main_Grid.ActualHeight - h - 4));
            BorderPastePrompt.Margin = new Thickness(x, y, 0, 0);
        }

        /// <summary>关闭粘贴气泡（不清剪贴板——图片还在，随时可再粘贴）</summary>
        private void HidePastePrompt()
        {
            BorderPastePrompt.Visibility = Visibility.Collapsed;
        }

        /// <summary>气泡按钮：粘贴（按剪贴板当前内容分流：墨迹中心对准点击点落下，或图片同规则插入，
        /// 随后退出粘贴模式）</summary>
        private void BtnPastePromptPaste_Click(object sender, RoutedEventArgs e)
        {
            HidePastePrompt();
            try
            {
                // 优先墨迹（ISF）：纯墨迹复制 → 粘贴仍是墨迹（可编辑），中心对准点击点
                var strokes = TryGetClipboardStrokes();
                if (strokes != null)
                {
                    PasteStrokesToCanvas(strokes, _pendingPasteCenter);
                }
                else
                {
                    var img = TryGetClipboardImage();
                    if (img != null && ImageLayer_AddImage(img, center: _pendingPasteCenter) != null)
                        ShowNotification(currentMode == 0
                            ? "已粘贴到批注层（第 0 页）"
                            : "已粘贴到白板第 " + CurrentWhiteboardIndex + " 页");
                }
            }
            catch (Exception ex)
            {
                ShowNotification("粘贴失败：" + ex.Message);
            }
            ExitPasteMode(); // 粘贴完成：恢复原工具和光标
        }

        /// <summary>气泡按钮：取消（图片仍留在剪贴板，退出粘贴模式恢复原工具）</summary>
        private void BtnPastePromptCancel_Click(object sender, RoutedEventArgs e)
        {
            ExitPasteMode();
        }

        #endregion

        #region 自绘遮罩框选截图（核心流程）

        /// <summary>重入锁：截图流程进行中忽略再次触发（防连点/快捷键连按）</summary>
        bool _isScreenshotBusy = false;

        /// <summary>截图期间被隐藏的 UI 元素表（元素 → 隐藏前的可见性），恢复时逐个还原</summary>
        System.Collections.Generic.Dictionary<UIElement, Visibility> _screenshotHiddenUi;

        /// <summary>截图期间是否隐藏了数学输入面板（独立 COM 窗口，恢复要单独 Show）</summary>
        bool _mathPanelHiddenForScreenshot = false;

        /// <summary>
        /// 截图入口（两个菜单项共用，只差"进遮罩前藏不藏 UI"）。
        /// hideUiFirst=false：快速截图——直接上遮罩，屏幕原样供框选；
        /// hideUiFirst=true：隐藏界面截图——先藏 UI、等渲染完成再上遮罩，遮罩下只剩板书。
        /// </summary>
        private void Screenshot_Start(bool hideUiFirst)
        {
            if (_isScreenshotBusy) return; // 上一次截图还没收尾，直接忽略
            _isScreenshotBusy = true;
            try
            {
                if (hideUiFirst)
                {
                    Screenshot_HideChrome(hideBoard: true);
                    // 等 UI 消失真正画完再上遮罩（两级 Background 回调 = 至少跑完一轮渲染），
                    // 否则遮罩下还能看到悬浮条的最后一帧
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() =>
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() =>
                            Screenshot_ShowMask(hideUiFirst)))));
                }
                else
                {
                    Screenshot_ShowMask(hideUiFirst);
                }
            }
            catch (Exception ex)
            {
                // 任何异常都必须恢复 UI 并解锁，不能出现"界面消失回不来"的死局
                Screenshot_RestoreChrome();
                _isScreenshotBusy = false;
                ShowNotification("截图失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 弹出遮罩窗口收集选区（模态）。返回后按"有选区/取消"两条路收尾。
        /// </summary>
        private void Screenshot_ShowMask(bool hideUiFirst)
        {
            var mask = new ScreenshotMaskWindow();
            if (!hideUiFirst)
            {
                // 快速截图：选区确定的一瞬间（遮罩还盖着屏幕时）才藏软件 UI——
                // 这样遮罩关闭后屏幕直接就是干净画面，不会有悬浮条闪现帧。
                // 注意 hideBoard=false：板书/白板保留入镜（快速截图截"眼前的一切"）
                mask.SelectionMade += () => Screenshot_HideChrome(hideBoard: false);
            }

            bool? ok = false;
            Rect rect = Rect.Empty;
            double scaleX = 1, scaleY = 1;
            try
            {
                ok = mask.ShowDialog(); // 模态：窗口关闭（完成或取消）才返回
                if (mask.SelectedRect.HasValue)
                {
                    rect = mask.SelectedRect.Value;
                    scaleX = mask.DpiScaleX;
                    scaleY = mask.DpiScaleY;
                }
            }
            catch { ok = false; }

            if (ok == true && !rect.IsEmpty)
            {
                // ---- 完成路径：等遮罩退场渲染完 → 抓屏 → 落图 → 恢复 ----
                // 逻辑坐标 × DPI 缩放 = 物理像素（CopyFromScreen 用物理像素；
                // 150% 缩放的教学一体机不做这步截图区域会偏移错位）
                int px = (int)Math.Round(rect.X * scaleX);
                int py = (int)Math.Round(rect.Y * scaleY);
                int pw = (int)Math.Round(rect.Width * scaleX);
                int ph = (int)Math.Round(rect.Height * scaleY);

                // 两级 Background 回调：确保"遮罩已关 + UI 已藏"渲染完毕，屏幕稳定后再抓
                Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() =>
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() =>
                    {
                        try
                        {
                            var bmp = Screenshot_CaptureRegion(px, py, pw, ph);
                            if (bmp == null)
                            {
                                ShowNotification("截图失败：无法捕获屏幕区域");
                                return;
                            }
                            // 两种截图都进剪贴板（可直接 Ctrl+V 到 PPT/微信）
                            try { Clipboard.SetImage(bmp); } catch { }
                            if (hideUiFirst)
                            {
                                // 隐藏界面截图：保持原行为——自动插入当前页（用户定稿）。
                                // 顺带结束可能存在的待粘贴状态（此路径不需要选位置）
                                ExitPasteMode();
                                if (ImageLayer_AddImage(bmp) != null)
                                    ShowNotification("截图已插入，并复制到剪贴板");
                            }
                            else
                            {
                                // 快速截图：只存剪贴板 + 进入粘贴模式，不自动插图——
                                // 核心诉求：第一页已有内容时，截的图要能放到第二页。
                                // 进入模态粘贴（画布挂起+专用光标），翻到目标页后点击白板选位置
                                // → 弹粘贴气泡；或 Ctrl+V / 截图菜单"粘贴图片"直接插入当前页
                                _pendingPasteImage = bmp;
                                EnterPasteMode();
                                ShowNotification("已复制到剪贴板——点击白板选位置粘贴（Esc 取消），或 Ctrl+V 插入当前页");
                            }
                        }
                        catch (Exception ex)
                        {
                            ShowNotification("截图失败：" + ex.Message);
                        }
                        finally
                        {
                            Screenshot_RestoreChrome();
                            _isScreenshotBusy = false;
                        }
                    }))));
            }
            else
            {
                // ---- 取消路径：瞬时恢复原状，无任何副作用 ----
                Screenshot_RestoreChrome();
                _isScreenshotBusy = false;
            }
        }

        /// <summary>
        /// 抓取屏幕指定区域（物理像素坐标）为 WPF 位图。
        /// 注意：必须在遮罩窗口关闭之后调用，否则会把遮罩自己截进去。
        /// </summary>
        private static BitmapSource Screenshot_CaptureRegion(int x, int y, int width, int height)
        {
            if (width < 1 || height < 1) return null;
            try
            {
                using (var bmp = new System.Drawing.Bitmap(width, height,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    using (var g = System.Drawing.Graphics.FromImage(bmp))
                    {
                        g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height),
                            System.Drawing.CopyPixelOperation.SourceCopy);
                    }
                    // GDI 位图 → WPF 位图源（复制像素，之后 GDI 资源可安全释放）
                    IntPtr hBitmap = bmp.GetHbitmap();
                    try
                    {
                        return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                            hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    }
                    finally
                    {
                        DeleteObject(hBitmap); // GetHbitmap 创建的非托管资源必须手动释放（防内存泄漏）
                    }
                }
            }
            catch { return null; }
        }

        /// <summary>
        /// 隐藏本软件所有会入镜的 UI。
        /// hideBoard=true（隐藏界面截图）：掀板去墨——把白板板面、墨迹/图片（inkCanvas）、
        /// 悬浮条、各弹层面板全部藏掉，回到"没开白板前"的裸桌面供框选。
        /// hideBoard=false（快速截图收尾时）：只藏悬浮条/面板——板书内容保留入镜
        /// （快速截图截的就是"眼前看到的一切"，选区确定瞬间只把软件 UI 撤走）。
        /// 细节：
        /// - GridBackgroundCoverHolder = 板面容器（白板模式 Visible / 批注模式本来 Collapsed），
        ///   藏它 = 掀板；批注模式（第 0 页透明白板）本来就无板面，记原值机制自动跳过。
        /// - inkCanvas = 墨迹 + 图片层；选区控制层父 Grid 绑定了 inkCanvas.Visibility，
        ///   藏它墨迹/图片/选区控件一层全隐（不动 Strokes 集合，零副作用，恢复一个属性搞定）。
        /// - 用 Visibility.Hidden（保持布局占位）而不是 Collapsed，恢复零风险；
        ///   记录隐藏前的可见性，恢复时逐个还原（本来就不显示的不动它）。
        /// </summary>
        private void Screenshot_HideChrome(bool hideBoard)
        {
            try
            {
                if (_screenshotHiddenUi != null) return; // 已在隐藏状态，不重复记录

                // 悬浮条（含笑脸收起态）+ 工具/图形面板 + 各弹出浮层（两种截图都要藏）
                var targets = new System.Collections.Generic.List<UIElement>
                {
                    ViewboxFloatingBar, BorderTools, BorderDrawShape,
                    BorderImageMenu, PenSettingsPanel, EraserSettingsPanel
                };
                if (hideBoard)
                {
                    targets.Add(GridBackgroundCoverHolder); // 板面（掀板；批注模式已 Collapsed 自动跳过）
                    targets.Add(inkCanvas);                 // 墨迹 + 图片（选区控制层绑定联动隐藏）
                }

                _screenshotHiddenUi = new System.Collections.Generic.Dictionary<UIElement, Visibility>();
                foreach (var t in targets)
                {
                    if (t == null) continue;
                    if (t.Visibility == Visibility.Hidden) continue; // 本来就藏的不记录，避免恢复时误显示
                    _screenshotHiddenUi[t] = t.Visibility;
                    t.Visibility = Visibility.Hidden;
                }

                // 数学输入面板是独立 COM 窗口（micaut），不属 WPF 视觉树，单独处理
                try
                {
                    bool visible = false;
                    _mathInputPanel?.IsVisible(out visible);
                    if (visible)
                    {
                        _mathPanelHiddenForScreenshot = true;
                        _mathInputPanel?.Hide();
                    }
                }
                catch { }
            }
            catch { }
        }

        /// <summary>恢复被隐藏的 UI（与 HideChrome 严格配对，可安全重复调用）</summary>
        private void Screenshot_RestoreChrome()
        {
            try
            {
                if (_screenshotHiddenUi != null)
                {
                    foreach (var kv in _screenshotHiddenUi)
                        kv.Key.Visibility = kv.Value; // 还原到隐藏前的原值
                    _screenshotHiddenUi = null;
                }
                if (_mathPanelHiddenForScreenshot)
                {
                    _mathPanelHiddenForScreenshot = false;
                    try { _mathInputPanel?.Show(); } catch { }
                }
            }
            catch { }
        }

        // Win32：释放 GDI 位图句柄（Screenshot_CaptureRegion 用）
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        #endregion
    }
}
