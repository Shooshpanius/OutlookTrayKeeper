// ============================================================================
//  OutlookTrayKeeper — [X] и Alt+F4 сворачивают классический Outlook в трей.
//
//  Что делает:
//    • отключает команду закрытия главного окна Outlook (кнопка [X] становится
//      серой — «случайно» закрыть Outlook больше не получится);
//    • перехватывает клик по [X] и нажатие Alt+F4 и превращает их в свертывание
//      Outlook в область уведомлений (включая штатную опцию Outlook
//      «Скрыть при свертывании» через параметр реестра MinToTray);
//    • настоящий выход остаётся доступным: «Файл → Выход» в Outlook или
//      «Завершить Outlook» в меню значка утилиты.
//
//  Версия 1.4 — выяснилось, что РЕАЛЬНЫЙ классический Outlook отвечает на
//  WM_NCHITTEST кодом HTCAPTION по всей полосе своих кнопок заголовка
//  (включая [X]), поэтому основная проверка «hittest == HTCLOSE» на нём не
//  срабатывала никогда — клик по [X] проходил мимо перехвата, а серая кнопка
//  ничего не делала. Теперь правая треть полосы кнопок (по данным DWM)
//  проверяется и в этом случае, а не только когда окно занято и не отвечает.
//
//  Версия 1.3 — клик по [X] не распознавался в двух случаях:
//    (1) перехват сравнивал координаты только с ПЕРЕДНИМ окном, а активация
//    происходит уже после нажатия кнопки, — клик по [X] неактивного окна
//    Outlook игнорировался; теперь берётся окно ПОД КУРСОРОМ
//    (WindowFromPoint + GetAncestor(GA_ROOT));
//    (2) «запасная проверка по прямоугольнику кнопки» из v1.2 не работала
//    никогда: прямоугольники кнопок есть только в структуре TITLEBARINFOEX
//    (сообщение WM_GETTITLEBARINFOEX), а не в TITLEBARINFO, где этот массив
//    отсутствует, — из-за лишнего поля cbSize не совпадал с настоящим.
//    Теперь запасная проверка — полоса кнопок заголовка из данных DWM
//    (DwmGetWindowAttribute/DWMWA_CAPTION_BUTTON_BOUNDS): она не требует
//    ответа от самого окна.
//
//  Версия 1.2 — клик по [X] не распознавался на экранах с масштабированием
//  (125% / 150% и т.п.): процесс без манифеста оставался DPI-unaware, Windows
//  масштабировала для него координаты, и точка клика из хука не совпадала
//  с координатами окна Outlook (Alt+F4 работал — координаты ему не нужны).
//  Теперь при старте включается DPI-awareness и ведётся журнал диагностики:
//  меню значка → «Открыть журнал диагностики» (файл %TEMP%\OutlookTrayKeeper.log).
//
//  Версия 1.1 — исправлено сравнение результата WM_NCHITTEST: код попадания
//  приходит в out-параметре SendMessageTimeout (в v1.0 сравнивалось
//  возвращаемое значение — из-за этого клик по [X] игнорировался).
//
//  Сборка: build.bat (штатный csc.exe из .NET Framework 4.x — есть в Windows
//  10/11 по умолчанию). Ничего скачивать и устанавливать не нужно.
//
//  Совместимость: классический Outlook 2016/2019/2021/2024/365 (outlook.exe).
//  «Новый Outlook» (веб-версия, olk.exe) не поддерживается.
//
//  Лицензия: MIT.
// ============================================================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("OutlookTrayKeeper")]
[assembly: AssemblyProduct("OutlookTrayKeeper")]
[assembly: AssemblyDescription("[X] and Alt+F4 minimize classic Outlook (outlook.exe) to the notification area instead of exiting")]
[assembly: AssemblyVersion("1.4.0.0")]

namespace OutlookTrayKeeper
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            // Включаем режим реальных координат экрана (DPI-awareness) — строго
            // ДО создания любых окон. Без этого при масштабировании 125%/150%
            // Windows «виртуализует» координаты нашего процесса, точка клика из
            // хука перестаёт совпадать с координатами Outlook, и клик по [X]
            // не распознаётся.
            Native.DpiResult = Native.BestEffortDpiAwareness();
            bool createdNew;
            Mutex mutex = new Mutex(true, "OutlookTrayKeeper_SingleInstance", out createdNew);
            if (!createdNew)
            {
                MessageBox.Show(
                    "OutlookTrayKeeper уже запущен — найдите его значок (конверт)\nв области уведомлений рядом с часами.",
                    "OutlookTrayKeeper", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                Application.Run(new KeeperContext());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Критическая ошибка OutlookTrayKeeper:\n\n" + ex.ToString(),
                    "OutlookTrayKeeper", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                GC.KeepAlive(mutex);
            }
        }
    }

    /// <summary>
    /// Контекст приложения: значок в трее, таймер наблюдения за Outlook,
    /// глобальные низкоуровневые хуки мыши и клавиатуры.
    /// </summary>
    internal sealed class KeeperContext : ApplicationContext
    {
        private const string CFG_KEY = @"Software\OutlookTrayKeeper";
        private const string CFG_FULL = @"HKEY_CURRENT_USER\Software\OutlookTrayKeeper";
        private const string RUN_KEY = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RUN_SUBKEY = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RUN_NAME = "OutlookTrayKeeper";
        private const string OUTLOOK_CLASS = "rctrl_renwnd32";   // класс главного окна классического Outlook

        private NotifyIcon _tray;
        private ContextMenuStrip _menu;
        private System.Windows.Forms.Timer _timer;
        private ToolStripMenuItem _miShow;
        private ToolStripMenuItem _miMinimize;
        private ToolStripMenuItem _miExitOutlook;
        private ToolStripMenuItem _miAllowClose;
        private ToolStripMenuItem _miAutorun;
        private ToolStripMenuItem _miLog;
        private ToolStripMenuItem _miAbout;
        private ToolStripMenuItem _miQuit;

        private IntPtr _mouseHook = IntPtr.Zero;
        private IntPtr _kbHook = IntPtr.Zero;
        private Native.HookProc _mouseProc;      // ссылку держим, чтобы GC не собрал делегат
        private Native.HookProc _kbProc;

        private HashSet<uint> _outlookPids = new HashSet<uint>();
        private List<IntPtr> _windows = new List<IntPtr>();   // главные окна Outlook (текущие)
        private List<IntPtr> _grayed = new List<IntPtr>();    // окна, у которых [X] переведён в серый
        private IntPtr _mainHwnd = IntPtr.Zero;

        private bool _allowCloseMode;            // «Разрешить закрытие» (перехват выключен)
        private bool _suppressNextLButtonUp;      // проглотить отпускание кнопки после блока клика по [X]
        private bool _cleaned;

        private string _logPath;                  // %TEMP%\OutlookTrayKeeper.log
        private readonly object _logSync = new object();
        private int _lastWinCount = -1;
        private const string VERSION = "1.4";

        public KeeperContext()
        {
            InitSettings();
            LogInit();
            InstallHooks();
            BuildMenu();

            _tray = new NotifyIcon();
            _tray.Icon = CreateIcon();
            _tray.ContextMenuStrip = _menu;
            _tray.Visible = true;
            _tray.DoubleClick += delegate { ShowOutlook(); };

            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 1000;
            _timer.Tick += delegate { RefreshOutlookState(); };

            RefreshOutlookState();
            _timer.Start();
            ShowFirstRunBalloon();
        }

        // ------------------------------------------------------------------
        //  Настройки: автозапуск + параметр Outlook «Скрывать при свертывании»
        // ------------------------------------------------------------------

        private void InitSettings()
        {
            // Автозапуск: при первом запуске включаем (по желанию пользователя),
            // дальше ориентируемся на сохранённое предпочтение.
            try
            {
                object pref = Registry.GetValue(CFG_FULL, "AutorunPref", null);
                if (pref == null)
                {
                    SetAutorun(true);
                    Registry.SetValue(CFG_FULL, "AutorunPref", 1, RegistryValueKind.DWord);
                }
                else if ((int)pref == 1)
                {
                    SetAutorun(true);   // обновляем путь (exe мог переехать после пересборки)
                }
            }
            catch { }

            // MinToTray=1 — та же галка, что «Скрыть при свертывании» в меню
            // значка Outlook. Свернутое окно прячется в трей, а не на панель задач.
            try
            {
                string[] vers = new string[] { "16.0", "15.0" };   // 2016–365 и 2013
                foreach (string ver in vers)
                {
                    string prefPath = @"Software\Microsoft\Office\" + ver + @"\Outlook\Preferences";
                    RegistryKey k;
                    if (ver == "16.0") k = Registry.CurrentUser.CreateSubKey(prefPath);
                    else k = Registry.CurrentUser.OpenSubKey(prefPath, true);
                    if (k == null) continue;
                    object v = k.GetValue("MinToTray");
                    if (v == null || (int)v != 1) k.SetValue("MinToTray", 1, RegistryValueKind.DWord);
                    k.Close();
                }
            }
            catch { }
        }

        private static bool IsAutorunEnabled()
        {
            try { return Registry.GetValue(RUN_KEY, RUN_NAME, null) != null; }
            catch { return false; }
        }

        private static void SetAutorun(bool on)
        {
            if (on)
            {
                Registry.SetValue(RUN_KEY, RUN_NAME, "\"" + Application.ExecutablePath + "\"",
                    RegistryValueKind.String);
            }
            else
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RUN_SUBKEY, true))
                {
                    if (k != null) k.DeleteValue(RUN_NAME, false);
                }
            }
        }

        // ------------------------------------------------------------------
        //  Глобальные низкоуровневые хуки (без внедрения DLL в Outlook)
        // ------------------------------------------------------------------

        private void InstallHooks()
        {
            IntPtr hMod = Native.GetModuleHandle(null);
            _mouseProc = new Native.HookProc(MouseHookProc);
            _kbProc = new Native.HookProc(KeyboardHookProc);
            _mouseHook = Native.SetWindowsHookEx(Native.WH_MOUSE_LL, _mouseProc, hMod, 0);
            _kbHook = Native.SetWindowsHookEx(Native.WH_KEYBOARD_LL, _kbProc, hMod, 0);
            Log("хуки: мышь=" + (_mouseHook != IntPtr.Zero ? "OK" : "ОШИБКА")
                + ", клавиатура=" + (_kbHook != IntPtr.Zero ? "OK" : "ОШИБКА"));
            if (_mouseHook == IntPtr.Zero || _kbHook == IntPtr.Zero)
            {
                int err = Marshal.GetLastWin32Error();
                MessageBox.Show(
                    "Не удалось установить перехват ввода (код " + err + ").\n\n" +
                    "Утилита продолжит работу, но клик по [X] может закрывать Outlook.\n" +
                    "Попробуйте перезапустить утилиту (при необходимости — от администратора).",
                    "OutlookTrayKeeper", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ------------------------------------------------------------------
        //  Меню значка в трее
        // ------------------------------------------------------------------

        private void BuildMenu()
        {
            _menu = new ContextMenuStrip();

            _miShow = new ToolStripMenuItem("Показать Outlook");
            _miShow.Click += delegate { ShowOutlook(); };

            _miMinimize = new ToolStripMenuItem("Свернуть Outlook в трей");
            _miMinimize.Click += delegate { MinimizeOutlook(); };

            _miExitOutlook = new ToolStripMenuItem("Завершить Outlook…");
            _miExitOutlook.Click += delegate { ExitOutlook(); };

            _miAllowClose = new ToolStripMenuItem("Разрешить закрытие (снять перехват)");
            _miAllowClose.Click += delegate { ToggleAllowClose(); };

            _miAutorun = new ToolStripMenuItem("Автозапуск при входе в Windows");
            _miAutorun.Click += ToggleAutorun;

            _miLog = new ToolStripMenuItem("Открыть журнал диагностики");
            _miLog.Click += delegate { OpenLog(); };

            _miAbout = new ToolStripMenuItem("О программе…");
            _miAbout.Click += delegate { ShowAbout(); };

            _miQuit = new ToolStripMenuItem("Выход из OutlookTrayKeeper…");
            _miQuit.Click += delegate { QuitKeeper(); };

            _menu.Items.Add(_miShow);
            _menu.Items.Add(_miMinimize);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(_miExitOutlook);
            _menu.Items.Add(_miAllowClose);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(_miAutorun);
            _menu.Items.Add(_miLog);
            _menu.Items.Add(_miAbout);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(_miQuit);

            _menu.Opening += delegate(object s, CancelEventArgs e) { UpdateMenuState(); };
        }

        private void UpdateMenuState()
        {
            bool running = _mainHwnd != IntPtr.Zero && Native.IsWindow(_mainHwnd);
            _miShow.Text = running ? "Показать Outlook" : "Запустить Outlook";
            _miMinimize.Enabled = running;
            _miExitOutlook.Enabled = running;
            _miAllowClose.Checked = _allowCloseMode;
            _miAutorun.Checked = IsAutorunEnabled();
            if (_tray != null)
            {
                string tip = running
                    ? "OutlookTrayKeeper — Outlook запущен; [X] = свернуть в трей"
                    : "OutlookTrayKeeper — ждёт запуска Outlook";
                if (_tray.Text != tip) _tray.Text = tip;
            }
        }

        // ------------------------------------------------------------------
        //  Действия меню
        // ------------------------------------------------------------------

        private void ShowOutlook()
        {
            if (_mainHwnd == IntPtr.Zero || !Native.IsWindow(_mainHwnd))
            {
                try { Process.Start("outlook.exe"); }
                catch (Exception ex)
                {
                    MessageBox.Show("Не удалось запустить Outlook:\n" + ex.Message,
                        "OutlookTrayKeeper", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return;
            }
            if (!Native.IsWindowVisible(_mainHwnd)) Native.ShowWindow(_mainHwnd, Native.SW_SHOW);
            if (Native.IsIconic(_mainHwnd)) Native.ShowWindow(_mainHwnd, Native.SW_RESTORE);
            Native.PostMessage(_mainHwnd, Native.WM_SYSCOMMAND, Native.SC_RESTORE, IntPtr.Zero);
            Native.SetForegroundWindow(_mainHwnd);
        }

        private void MinimizeOutlook()
        {
            if (_mainHwnd != IntPtr.Zero && Native.IsWindow(_mainHwnd))
            {
                Native.PostMessage(_mainHwnd, Native.WM_SYSCOMMAND, Native.SC_MINIMIZE, IntPtr.Zero);
            }
        }

        private void ExitOutlook()
        {
            if (_mainHwnd == IntPtr.Zero || !Native.IsWindow(_mainHwnd)) return;
            DialogResult r = MessageBox.Show("Полностью завершить работу Outlook?",
                "OutlookTrayKeeper", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            if (r != DialogResult.Yes) return;
            SetCloseAllowed();   // вернуть [X] в активный вид
            Native.PostMessage(_mainHwnd, Native.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        private void ToggleAllowClose()
        {
            _allowCloseMode = !_allowCloseMode;
            if (_allowCloseMode)
            {
                SetCloseAllowed();
                if (_tray != null)
                {
                    _tray.ShowBalloonTip(4000, "OutlookTrayKeeper",
                        "Перехват отключён: [X] и Alt+F4 снова закрывают Outlook,\n" +
                        "пока вы не снимете галочку в меню значка.", ToolTipIcon.Warning);
                }
            }
            UpdateMenuState();
        }

        private void ToggleAutorun(object sender, EventArgs e)
        {
            bool on = !_miAutorun.Checked;
            try
            {
                SetAutorun(on);
                Registry.SetValue(CFG_FULL, "AutorunPref", on ? 1 : 0, RegistryValueKind.DWord);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось изменить автозагрузку:\n" + ex.Message,
                    "OutlookTrayKeeper", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            UpdateMenuState();
        }

        private void ShowAbout()
        {
            // Небольшой диалог вместо MessageBox: в MessageBox гиперссылку
            // сделать кликабельной нельзя, а в LinkLabel — можно.
            using (Form f = new Form())
            {
                f.Text = "О программе — OutlookTrayKeeper";
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterScreen;
                f.MinimizeBox = false;
                f.MaximizeBox = false;
                f.ShowInTaskbar = false;
                f.Font = SystemFonts.MessageBoxFont;
                if (_tray != null) f.Icon = _tray.Icon;

                Label top = new Label();
                top.Text =
                    "OutlookTrayKeeper " + VERSION + "\n\n" +
                    "Нажатие [X] или Alt+F4 в классическом Outlook (outlook.exe)\n" +
                    "сворачивает окно в область уведомлений вместо закрытия.\n\n" +
                    "Вернуть окно: щелчок по значку Outlook в трее\n" +
                    "(или двойной щелчок по значку утилиты).\n\n" +
                    "Настоящий выход из Outlook:\n" +
                    "  • меню Outlook «Файл» → «Выход»;\n" +
                    "  • или меню этого значка → «Завершить Outlook…».\n\n" +
                    "Исходный код и обновления:";
                LinkLabel link = new LinkLabel();
                link.Text = "github.com/Shooshpanius/OutlookTrayKeeper";
                link.LinkClicked += delegate
                {
                    try { Process.Start("https://github.com/Shooshpanius/OutlookTrayKeeper"); }
                    catch { }
                };
                Label bottom = new Label();
                bottom.Text =
                    "Утилита не изменяет сам Outlook: все её настройки обратимы\n" +
                    "и полностью удаляются вместе с программой.\n\n" +
                    "Если [X] не сворачивает окно:\n" +
                    "меню значка → «Открыть журнал диагностики» — в журнале видно,\n" +
                    "как перехватчик распознал последний клик (приложите его к письму).";
                Button ok = new Button();
                ok.Text = "OK";
                ok.DialogResult = DialogResult.OK;
                ok.AutoSize = true;

                // Вёрстка по MeasureString: у LinkLabel с AutoSize бывают
                // неверные расчёты высоты (ссылка наезжала на текст ниже),
                // поэтому размеры всех блоков считаем сами. Высота — максимум
                // из двух рендеров текста (GDI+ и GDI) плюс запас.
                int width = 420;
                top.AutoSize = false;
                link.AutoSize = false;
                bottom.AutoSize = false;
                link.LinkArea = new LinkArea(0, link.Text.Length);
                using (Graphics g = f.CreateGraphics())
                {
                    top.Size = TextBlockSize(g, top.Text, f.Font, width);
                    link.Size = TextBlockSize(g, link.Text, f.Font, width - 28);   // x=42: правый край как у остальных
                    bottom.Size = TextBlockSize(g, bottom.Text, f.Font, width);
                }

                int y = 14;
                top.Location = new Point(14, y);
                y += top.Height + 8;
                link.Location = new Point(42, y);
                y += link.Height + 8;
                bottom.Location = new Point(14, y);
                y += bottom.Height + 10;
                ok.Location = new Point(width + 28 - 90, y);
                f.ClientSize = new Size(width + 28, y + ok.PreferredSize.Height + 14);
                f.AcceptButton = ok;
                f.CancelButton = ok;

                f.Controls.Add(top);
                f.Controls.Add(link);
                f.Controls.Add(bottom);
                f.Controls.Add(ok);
                f.ShowDialog();
            }
        }

        /// <summary>
        /// Размер текстового блока диалога: максимум из измерений GDI+
        /// (Graphics.MeasureString) и GDI (TextRenderer — им рисует Label)
        /// плюс небольшой запас, чтобы текст не обрезался.
        /// </summary>
        private static Size TextBlockSize(Graphics g, string text, Font font, int width)
        {
            SizeF sgdiPlus = g.MeasureString(text, font, width);
            Size sgdi = TextRenderer.MeasureText(text, font,
                new Size(width, int.MaxValue), TextFormatFlags.WordBreak);
            return new Size(width, Math.Max((int)Math.Ceiling(sgdiPlus.Height), sgdi.Height) + 2);
        }

        private void QuitKeeper()
        {
            DialogResult r = MessageBox.Show(
                "Выйти из OutlookTrayKeeper?\n" +
                "Кнопка [X] Outlook снова начнёт закрывать программу\n" +
                "(утилита вернёт кнопку в обычный вид).",
                "OutlookTrayKeeper", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            if (r != DialogResult.Yes) return;
            Cleanup();
            Application.ExitThread();
        }

        private void Cleanup()
        {
            if (_cleaned) return;
            _cleaned = true;
            try
            {
                if (_mouseHook != IntPtr.Zero) { Native.UnhookWindowsHookEx(_mouseHook); _mouseHook = IntPtr.Zero; }
                if (_kbHook != IntPtr.Zero) { Native.UnhookWindowsHookEx(_kbHook); _kbHook = IntPtr.Zero; }
                SetCloseAllowed();   // вернуть [X] всех известных окон Outlook в активный вид
            }
            catch { }
            if (_timer != null) { _timer.Stop(); _timer.Dispose(); _timer = null; }
            if (_tray != null) { _tray.Visible = false; _tray.Dispose(); _tray = null; }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) Cleanup();
            base.Dispose(disposing);
        }

        // ------------------------------------------------------------------
        //  Перехват: клик по [X] сворачивает Outlook в трей
        // ------------------------------------------------------------------

        private IntPtr MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode == Native.HC_ACTION && !_allowCloseMode)
                {
                    uint msg = (uint)wParam.ToInt64();
                    if (msg == Native.WM_LBUTTONDOWN)
                    {
                        Native.MSLLHOOKSTRUCT ms =
                            (Native.MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.MSLLHOOKSTRUCT));
                        // v1.3: проверяем окно ПОД КУРСОРОМ, а не переднее окно.
                        // Активация происходит уже ПОСЛЕ нажатия кнопки: если
                        // Outlook неактивен и пользователь щёлкает его [X],
                        // в момент LBUTTONDOWN передним всё ещё является чужое
                        // окно — v1.2 такой клик не проверяла вовсе.
                        IntPtr target = Native.WindowFromPoint(ms.pt);
                        if (target != IntPtr.Zero) target = Native.GetAncestor(target, Native.GA_ROOT);
                        if (target != IntPtr.Zero && IsOutlookMain(target))
                        {
                            // Куда попал клик? Две независимые проверки:
                            //  1. хит-тест самого окна: WM_NCHITTEST -> HTCLOSE
                            //     (код попадания приходит в out-параметре
                            //     SendMessageTimeout; возвращаемое значение —
                            //     лишь признак успеха отправки);
                            //  2. правая часть полосы кнопок заголовка из
                            //     данных DWM. Нужна в двух случаях:
                            //      • окно занято и не ответило на хит-тест;
                            //      • РЕАЛЬНЫЙ классический Outlook отвечает
                            //        HTCAPTION на всю полосу своих кнопок
                            //        (установлено по журналу v1.3: клики точно
                            //        по [X] и по «свернуть» давали hittest=2),
                            //        поэтому на нём HTCLOSE не приходит никогда.
                            // Кнопки «свернуть»/«развернуть» эта проверка не
                            // задевает: на обычных окнах хит-тест честно
                            // отвечает HTMINBUTTON/HTMAXBUTTON, и до неё дело
                            // не доходит.
                            int ht = Native.WindowHitTest(target, ms.pt.X, ms.pt.Y);
                            bool htClose = (ht == Native.HTCLOSE);
                            bool dwmClose = false;
                            if (!htClose && (ht == Native.HTCAPTION || ht == Native.HT_NO_ANSWER))
                            {
                                dwmClose = Native.DwmCloseRectContains(target, ms.pt.X, ms.pt.Y);
                            }
                            if (htClose || dwmClose)
                            {
                                // Это клик по [X]: сворачиваем Outlook в трей и глотаем клик,
                                // чтобы команда закрытия вообще не дошла до окна.
                                Log("клик по [X]: pt=(" + ms.pt.X + "," + ms.pt.Y + ") hittest=" + ht
                                    + (dwmClose ? (ht == Native.HT_NO_ANSWER
                                        ? " (окно не ответило, запасная проверка DWM)"
                                        : " (кнопки заголовка отвечают HTCAPTION, полоса кнопок DWM)") : "")
                                    + " -> свернуть в трей");
                                Native.PostMessage(target, Native.WM_SYSCOMMAND, Native.SC_MINIMIZE, IntPtr.Zero);
                                _suppressNextLButtonUp = true;
                                return (IntPtr)1;
                            }
                            if (ht == Native.HT_NO_ANSWER)
                            {
                                Log("клик: Outlook не ответил на хит-тест (занят), pt=("
                                    + ms.pt.X + "," + ms.pt.Y + ")");
                            }
                            else if (ht == Native.HTCAPTION)
                            {
                                Log("клик по заголовку (не [X]): pt=(" + ms.pt.X + "," + ms.pt.Y
                                    + ") hittest=" + ht + " -> пропущен");
                            }
                        }
                    }
                    else if (msg == Native.WM_LBUTTONUP && _suppressNextLButtonUp)
                    {
                        _suppressNextLButtonUp = false;
                        return (IntPtr)1;
                    }
                }
            }
            catch { }
            return Native.CallNextHookEx(_mouseHook, nCode, wParam, lParam);
        }

        // ------------------------------------------------------------------
        //  Перехват: Alt+F4 тоже сворачивает Outlook в трей
        // ------------------------------------------------------------------

        private IntPtr KeyboardHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode == Native.HC_ACTION && !_allowCloseMode)
                {
                    uint msg = (uint)wParam.ToInt64();
                    if (msg == Native.WM_SYSKEYDOWN)
                    {
                        Native.KBDLLHOOKSTRUCT kb =
                            (Native.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Native.KBDLLHOOKSTRUCT));
                        if (kb.vkCode == Native.VK_F4 && (kb.flags & Native.LLKHF_ALTDOWN) != 0
                            && !KeyIsDown(Native.VK_CONTROL) && !KeyIsDown(Native.VK_SHIFT)
                            && !KeyIsDown(Native.VK_LWIN) && !KeyIsDown(Native.VK_RWIN))
                        {
                            IntPtr fg = Native.GetForegroundWindow();
                            if (IsOutlookMain(fg))
                            {
                                Log("Alt+F4 в окне Outlook -> свернуть в трей");
                                Native.PostMessage(fg, Native.WM_SYSCOMMAND, Native.SC_MINIMIZE, IntPtr.Zero);
                                return (IntPtr)1;
                            }
                        }
                    }
                }
            }
            catch { }
            return Native.CallNextHookEx(_kbHook, nCode, wParam, lParam);
        }

        private static bool KeyIsDown(uint vk)
        {
            return (Native.GetAsyncKeyState((int)vk) & 0x8000) != 0;
        }

        /// <summary>Главное ли это окно Outlook (класс + pid + наличие кнопок свернуть/развернуть).</summary>
        private bool IsOutlookMain(IntPtr h)
        {
            if (h == IntPtr.Zero || _outlookPids.Count == 0) return false;
            if (!_windows.Contains(h))
            {
                uint pid;
                Native.GetWindowThreadProcessId(h, out pid);
                if (!_outlookPids.Contains(pid)) return false;
                StringBuilder sb = new StringBuilder(64);
                Native.GetClassName(h, sb, 64);
                if (sb.ToString() != OUTLOOK_CLASS) return false;
                if (!Native.IsMainStyleWindow(h)) return false;
            }
            return true;
        }

        // ------------------------------------------------------------------
        //  Наблюдение за Outlook (раз в секунду)
        // ------------------------------------------------------------------

        private void RefreshOutlookState()
        {
            // PID всех запущенных процессов outlook.exe
            _outlookPids.Clear();
            Process[] ps = Process.GetProcessesByName("OUTLOOK");
            foreach (Process p in ps)
            {
                try { _outlookPids.Add((uint)p.Id); }
                finally { p.Dispose(); }
            }

            // Их главные окна
            _windows = Native.FindOutlookWindows(_outlookPids);
            _mainHwnd = IntPtr.Zero;
            foreach (IntPtr h in _windows)
            {
                if (Native.IsWindowVisible(h)) { _mainHwnd = h; break; }
            }
            if (_mainHwnd == IntPtr.Zero && _windows.Count > 0) _mainHwnd = _windows[0];

            // Чистим «мертвые» окна из списка серых [X]
            for (int i = _grayed.Count - 1; i >= 0; i--)
            {
                if (!Native.IsWindow(_grayed[i])) _grayed.RemoveAt(i);
            }

            // Гасим [X] во всех главных окнах Outlook (самовосстановление, если
            // Outlook пересоздал окно или вернул кнопку)
            if (!_allowCloseMode)
            {
                foreach (IntPtr h in _windows)
                {
                    IntPtr hSys = Native.GetSystemMenu(h, false);
                    if (hSys == IntPtr.Zero) continue;
                    int st = Native.GetMenuState(hSys, Native.SC_CLOSE, (int)Native.MF_BYCOMMAND);
                    if ((st & 3) == 0)   // кнопка еще активна — делаем серой
                    {
                        Native.EnableMenuItem(hSys, Native.SC_CLOSE,
                            Native.MF_BYCOMMAND | Native.MF_GRAYED | Native.MF_DISABLED);
                        Native.SetWindowPos(h, IntPtr.Zero, 0, 0, 0, 0,
                            Native.SWP_NOSIZE | Native.SWP_NOMOVE | Native.SWP_NOZORDER
                            | Native.SWP_NOACTIVATE | Native.SWP_FRAMECHANGED);
                    }
                    if (!_grayed.Contains(h)) _grayed.Add(h);
                }
            }

            if (_windows.Count != _lastWinCount)
            {
                _lastWinCount = _windows.Count;
                Log("состояние Outlook: процессов=" + _outlookPids.Count
                    + ", главных окон=" + _windows.Count);
            }

            UpdateMenuState();
        }

        /// <summary>Вернуть кнопке [X] обычный вид во всех известных окнах Outlook.</summary>
        private void SetCloseAllowed()
        {
            foreach (IntPtr h in _grayed)
            {
                if (!Native.IsWindow(h)) continue;
                IntPtr hSys = Native.GetSystemMenu(h, false);
                if (hSys != IntPtr.Zero)
                {
                    Native.EnableMenuItem(hSys, Native.SC_CLOSE, Native.MF_BYCOMMAND | Native.MF_ENABLED);
                    Native.SetWindowPos(h, IntPtr.Zero, 0, 0, 0, 0,
                        Native.SWP_NOSIZE | Native.SWP_NOMOVE | Native.SWP_NOZORDER
                        | Native.SWP_NOACTIVATE | Native.SWP_FRAMECHANGED);
                }
            }
            _grayed.Clear();
        }

        // ------------------------------------------------------------------
        //  Журнал диагностики (%TEMP%\OutlookTrayKeeper.log)
        // ------------------------------------------------------------------

        private void LogInit()
        {
            _logPath = Path.Combine(Path.GetTempPath(), "OutlookTrayKeeper.log");
            try
            {
                FileInfo fi = new FileInfo(_logPath);
                if (fi.Exists && fi.Length > 262144) fi.Delete();   // не растим бесконечно
            }
            catch { }
            Log("=== запуск OutlookTrayKeeper " + VERSION + " | " + Environment.OSVersion.VersionString
                + " | 64-бит: " + (Environment.Is64BitProcess ? "да" : "нет")
                + " | DPI-режим: " + Native.DpiResult
                + " | журнал: " + _logPath);
        }

        private void Log(string msg)
        {
            try
            {
                lock (_logSync)
                {
                    if (_logPath != null)
                        File.AppendAllText(_logPath,
                            DateTime.Now.ToString("HH:mm:ss.fff") + "  " + msg + Environment.NewLine,
                            System.Text.Encoding.UTF8);
                }
            }
            catch { }
        }

        private void OpenLog()
        {
            try
            {
                if (!File.Exists(_logPath)) Log("журнал пуст — события еще не фиксировались");
                Process.Start("notepad.exe", "\"" + _logPath + "\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось открыть журнал:\n" + _logPath + "\n\n" + ex.Message,
                    "OutlookTrayKeeper", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ------------------------------------------------------------------
        //  Первый запуск: подсказка
        // ------------------------------------------------------------------

        private void ShowFirstRunBalloon()
        {
            try
            {
                object shown = Registry.GetValue(CFG_FULL, "IntroShown", null);
                if (shown == null && _tray != null)
                {
                    Registry.SetValue(CFG_FULL, "IntroShown", 1, RegistryValueKind.DWord);
                    _tray.ShowBalloonTip(9000, "OutlookTrayKeeper запущен",
                        "Кнопка [X] и Alt+F4 теперь сворачивают Outlook в область уведомлений.\n" +
                        "Вернуть окно — щелчок по значку Outlook в трее.\n" +
                        "Настоящий выход — меню этого значка → «Завершить Outlook…».",
                        ToolTipIcon.Info);
                }
            }
            catch { }
        }

        // ------------------------------------------------------------------
        //  Значок: конверт (рисуется программно, без внешних файлов)
        // ------------------------------------------------------------------

        private static Icon CreateIcon()
        {
            using (Bitmap bmp = new Bitmap(32, 32))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    using (SolidBrush bg = new SolidBrush(Color.FromArgb(0, 99, 177)))
                    {
                        g.FillRectangle(bg, 0, 0, 32, 32);
                    }
                    // белый конверт
                    g.FillRectangle(Brushes.White, 5, 9, 22, 15);
                    using (Pen p = new Pen(Color.FromArgb(0, 99, 177), 2.5f))
                    {
                        p.StartCap = LineCap.Round;
                        p.EndCap = LineCap.Round;
                        g.DrawLine(p, 5f, 10f, 16f, 18.5f);
                        g.DrawLine(p, 27f, 10f, 16f, 18.5f);
                    }
                }
                return Icon.FromHandle(bmp.GetHicon());
            }
        }
    }

    // ======================================================================
    //  Win32 API
    // ======================================================================

    internal static class Native
    {
        // --- сообщения и команды окна ---
        public const uint WM_SYSCOMMAND = 0x0112;
        public const uint WM_CLOSE = 0x0010;
        public const uint WM_NCHITTEST = 0x0084;
        public const int HTCLOSE = 20;
        public const int HTCAPTION = 2;
        public const int HTCLIENT = 1;
        public const int HT_NO_ANSWER = -999;   // окно не ответило на SendMessageTimeout
        public const uint SC_CLOSE = 0xF060;
        public static readonly IntPtr SC_MINIMIZE = (IntPtr)0xF020;
        public static readonly IntPtr SC_RESTORE = (IntPtr)0xF120;
        public const int SW_SHOW = 5;
        public const int SW_RESTORE = 9;

        // --- системное меню ---
        public const uint MF_BYCOMMAND = 0x0000;
        public const uint MF_ENABLED = 0x0000;
        public const uint MF_GRAYED = 0x0001;
        public const uint MF_DISABLED = 0x0002;

        // --- стили окна ---
        public const int GWL_STYLE = -16;
        public const int WS_SYSMENU = 0x00080000;
        public const int WS_MINIMIZEBOX = 0x00020000;
        public const int WS_MAXIMIZEBOX = 0x00010000;

        // --- низкоуровневые хуки ---
        public const int WH_MOUSE_LL = 14;
        public const int WH_KEYBOARD_LL = 13;
        public const int HC_ACTION = 0;
        public const uint WM_LBUTTONDOWN = 0x0201;
        public const uint WM_LBUTTONUP = 0x0202;
        public const uint WM_SYSKEYDOWN = 0x0104;
        public const uint VK_F4 = 0x73;
        public const uint VK_SHIFT = 0x10;
        public const uint VK_CONTROL = 0x11;
        public const uint VK_MENU = 0x12;
        public const uint VK_LWIN = 0x5B;
        public const uint VK_RWIN = 0x5C;
        public const uint LLKHF_ALTDOWN = 0x20;
        public const uint SMTO_ABORTIFHUNG = 0x0002;

        // --- SetWindowPos ---
        public const uint SWP_NOSIZE = 0x0001;
        public const uint SWP_NOMOVE = 0x0002;
        public const uint SWP_NOZORDER = 0x0004;
        public const uint SWP_NOACTIVATE = 0x0010;
        public const uint SWP_FRAMECHANGED = 0x0020;

        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("user32.dll")]
        public static extern int EnableMenuItem(IntPtr hMenu, uint uIDEnableItem, uint uEnable);

        [DllImport("user32.dll")]
        public static extern int GetMenuState(IntPtr hMenu, uint uId, int uFlags);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam,
            IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string lpFileName);

        /// <summary>Упаковать экранные координаты в lParam для WM_NCHITTEST.</summary>
        public static IntPtr PackPoint(int x, int y)
        {
            return (IntPtr)(int)(((y & 0xFFFF) << 16) | (x & 0xFFFF));
        }

        /// <summary>Похоже ли окно на главное окно приложения (есть кнопки свернуть/развернуть).</summary>
        public static bool IsMainStyleWindow(IntPtr h)
        {
            int style = GetWindowLong(h, GWL_STYLE);
            int required = WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX;
            return (style & required) == required;
        }

        /// <summary>Найти все главные окна Outlook (класс rctrl_renwnd32) по списку PID.</summary>
        public static List<IntPtr> FindOutlookWindows(HashSet<uint> pids)
        {
            List<IntPtr> result = new List<IntPtr>();
            if (pids == null || pids.Count == 0) return result;
            EnumWindowsProc proc = delegate(IntPtr h, IntPtr l)
            {
                uint pid;
                GetWindowThreadProcessId(h, out pid);
                if (!pids.Contains(pid)) return true;
                StringBuilder sb = new StringBuilder(64);
                GetClassName(h, sb, 64);
                if (sb.ToString() != "rctrl_renwnd32") return true;
                if (!IsMainStyleWindow(h)) return true;
                result.Add(h);
                return true;
            };
            EnumWindows(proc, IntPtr.Zero);
            GC.KeepAlive(proc);
            return result;
        }

        // --- v1.3: окно под курсором и полоса кнопок заголовка (DWM) ---
        // Прежний запасной способ (v1.2) — прямоугольники кнопок через
        // GetTitleBarInfo — был нерабочим: в TITLEBARINFO массив rgrect
        // отсутствует (прямоугольники заполняет только TITLEBARINFOEX по
        // сообщению WM_GETTITLEBARINFOEX, и уже на Windows 11 его данные
        // не соответствуют реальному расположению кнопок), а из-за лишнего
        // поля объявленная структура не совпадала с настоящей по размеру.

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public const uint GA_ROOT = 2;

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT pt);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        public const int DWMWA_CAPTION_BUTTON_BOUNDS = 5;

        [DllImport("dwmapi.dll")]
        public static extern int DwmGetWindowAttribute(IntPtr hWnd, int attr, out RECT rect, int cb);

        /// <summary>
        /// Проверка клика по [X] без опоры на ответ окна: полоса кнопок
        /// заголовка из данных DWM, [X] — её правая треть (ширина одной
        /// кнопки). Полоса в координатах, относительных окну; DWM отдаёт
        /// её даже, когда окно занято, не отвечает на сообщения или
        /// подменяет хит-тест заголовка (реальный Outlook отвечает
        /// HTCAPTION на всю полосу кнопок).
        /// </summary>
        public static bool DwmCloseRectContains(IntPtr h, int x, int y)
        {
            RECT strip;
            if (DwmGetWindowAttribute(h, DWMWA_CAPTION_BUTTON_BOUNDS, out strip,
                Marshal.SizeOf(typeof(RECT))) != 0) return false;
            RECT wr;
            if (!GetWindowRect(h, out wr)) return false;
            int width = strip.Right - strip.Left;
            int height = strip.Bottom - strip.Top;
            if (width <= 0 || height <= 0) return false;
            int right = wr.Left + strip.Right;
            int left = right - width / 3;   // ширина одной кнопки ~ трети полосы
            return x >= left && x < right
                && y >= wr.Top + strip.Top && y < wr.Top + strip.Bottom;
        }

        /// <summary>
        /// Хит-тест окна: код HT* или HT_NO_ANSWER, если окно не ответило вовремя.
        /// </summary>
        public static int WindowHitTest(IntPtr h, int x, int y)
        {
            IntPtr res = IntPtr.Zero;
            IntPtr ok = SendMessageTimeout(h, WM_NCHITTEST, IntPtr.Zero,
                PackPoint(x, y), SMTO_ABORTIFHUNG, 100, out res);
            if (ok == IntPtr.Zero) return HT_NO_ANSWER;
            return res.ToInt32();
        }

        public static string DpiResult = "не включался";

        /// <summary>
        /// Включить восприятие реального разрешения экрана (DPI-awareness).
        /// Вызывать строго ДО создания любых окон. Без этого Windows
        /// «виртуализует» координаты процесса, и при масштабировании
        /// 125%/150% точка клика из хука не совпадает с координатами
        /// окон других процессов (клике по [X] перестаёт распознаваться).
        /// </summary>
        public static string BestEffortDpiAwareness()
        {
            try
            {
                // Windows 10 1703+: per-monitor v2 — лучший вариант.
                if (SetProcessDpiAwarenessContext(new IntPtr(-4))) return "PerMonitorV2";
            }
            catch { }
            try
            {
                if (SetProcessDpiAwareness(2) == 0) return "PerMonitor";   // Windows 8.1+
            }
            catch { }
            try
            {
                if (SetProcessDPIAware()) return "System";                 // Vista+
            }
            catch { }
            return "не удалось (координаты могут масштабироваться)";
        }

        [DllImport("user32.dll")]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr value);

        [DllImport("shcore.dll")]
        private static extern int SetProcessDpiAwareness(int value);

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();
    }
}
