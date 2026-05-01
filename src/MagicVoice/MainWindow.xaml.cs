using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.UI;
using Microsoft.UI.Input;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Controls.Primitives;
using MagicVoice.Core;
using MagicVoice.Infrastructure;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using Windows.System;
using Windows.UI.Core;
using WinRT.Interop;

namespace MagicVoice;

public sealed partial class MainWindow : Window
{
    private const int DefaultWindowWidth = 1500;
    private const int DefaultWindowHeight = 960;
    private const int MinimumWindowWidth = 1180;
    /// <summary>Größe des Status-Infoleisten-Icons: Standard-InfoBar nutzt für „Informational“ einen kleinen ProgressRing in einer Kachel – wir ersetzen durch flächige MDL2-Glyphen.</summary>
    private const double StatusInfoBarIconSize = 22.0;
    private const int MinimumWindowHeight = 780;
    private const string OverviewPage = "overview";
    private const string AudioLanguagePage = "audioLanguage";
    private const string ProviderPage = "providers";
    private const string HotkeyPage = "hotkeys";
    private const string InsertPage = "insert";
    private const string GeneralPage = "general";
    private const string DiagnosticsPage = "diagnostics";
    private const string AboutPage = "about";
    private const string GlobalStandardLabel = "(globaler Standard)";

    private readonly ISettingsService _settingsService;
    private readonly ISecretProtector _secretProtector;
    private readonly ITrayStatusService _statusService;
    private readonly IHotkeyService _hotkeyService;
    private readonly IAutostartService _autostartService;
    private readonly IAudioDeviceService _audioDeviceService;
    private readonly ILlmService _llmService;
    private readonly FileLogger _fileLogger;
    private readonly IProcessingFailureLog _processingFailureLog;
    private readonly NavigationView _navigation = new()
    {
        PaneDisplayMode = NavigationViewPaneDisplayMode.Left,
        IsBackButtonVisible = NavigationViewBackButtonVisible.Collapsed,
        IsSettingsVisible = false
    };
    private readonly InfoBar _statusInfo = new() { IsOpen = true, Title = "Einrichtung erforderlich", Message = "Bitte prüfe Anbieter, API-Schlüssel und Tastenkürzel der Assistenten.", Severity = InfoBarSeverity.Warning };
    private readonly ComboBox _sttProvider = Combo("Anbieter");
    private readonly ComboBox _sttModel = new()
    {
        Header = "Modell",
        HorizontalAlignment = HorizontalAlignment.Stretch,
        MinWidth = 260,
        IsEditable = true,
        PlaceholderText = "Vorschlag aus Liste oder eigene Modell-ID (z. B. gpt-4o-mini-transcribe)"
    };
    private readonly TextBox _sttApiKey = new()
    {
        Header = "API-Schlüssel für die Transkription",
        PlaceholderText = "OpenAI-API-Schlüssel (sk-…)",
        HorizontalAlignment = HorizontalAlignment.Stretch,
        IsSpellCheckEnabled = false
    };
    private readonly ComboBox _audioInputDevice = Combo("Audioquelle");
    private readonly ComboBox _inputLanguage = Combo("Eingabesprache");
    private readonly ComboBox _outputLanguage = Combo("Ausgabesprache");
    private readonly ComboBox _llmProvider = Combo("Anbieter");
    private readonly ComboBox _llmModel = new()
    {
        Header = "Modell",
        HorizontalAlignment = HorizontalAlignment.Stretch,
        MinWidth = 260,
        IsEditable = true,
        PlaceholderText = "Vorschlag aus Liste oder eigene Modell-ID (z. B. gpt-4o)"
    };
    private readonly TextBox _llmApiKey = new()
    {
        Header = "API-Schlüssel für die KI-Verarbeitung",
        PlaceholderText = "OpenAI-API-Schlüssel (sk-…)",
        HorizontalAlignment = HorizontalAlignment.Stretch,
        IsSpellCheckEnabled = false
    };
    private readonly ComboBox _insertMethod = Combo("Einfügemethode");
    private readonly TextBox _insertionTestTarget = new() { TextWrapping = TextWrapping.Wrap, PlaceholderText = "Testeinfügung schreibt nur in dieses Feld.", MinHeight = 96, AcceptsReturn = true };
    private readonly CheckBox _restoreClipboard = new() { Content = "Zwischenablage nach dem Einfügen wiederherstellen" };
    private readonly CheckBox _launchMinimized = new() { Content = "beim Start in den Infobereich minimieren" };
    private readonly CheckBox _minimizeToTray = new() { Content = "beim Minimieren in den Infobereich verbergen" };
    private readonly CheckBox _playRecordingSounds = new() { Content = "Hinweistöne bei Aufnahme-Start und -Ende abspielen" };
    private readonly CheckBox _autostart = new() { Content = "mit Windows starten" };
    private readonly TextBox _maxSeconds = TextField("Maximale Aufnahmedauer", "60");
    private readonly TextBox _timeoutSeconds = TextField("Zeitlimit für die Verarbeitung (Sekunden)", "45");
    private readonly TextBox _diagnostics = new() { AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, IsReadOnly = true, MinHeight = 260, FontFamily = new FontFamily("Cascadia Mono") };
    private readonly Dictionary<string, TextBlock> _hotkeyValues = [];
    private readonly Dictionary<string, TextBox> _prompts = [];
    private readonly Dictionary<string, TextBox> _names = [];
    private readonly Dictionary<string, CheckBox> _systemPromptOverrideCheckBoxes = [];
    private readonly Dictionary<string, TextBox> _systemPrompts = [];
    private readonly Dictionary<string, ComboBox> _assistantInputLanguageOverrides = [];
    private readonly Dictionary<string, ComboBox> _assistantOutputLanguageOverrides = [];
    private readonly Dictionary<string, ToolTip> _policyTooltips = [];
    private readonly Dictionary<string, Slider> _intensitySliders = [];
    private readonly Dictionary<string, TextBlock> _intensityValues = [];
    private readonly Dictionary<string, Button> _captureButtons = [];
    private readonly Dictionary<string, ComboBox> _writingStyles = [];
    private readonly Dictionary<string, ComboBox> _paragraphDensities = [];
    private readonly Dictionary<string, ComboBox> _emojiExpressions = [];
    private readonly Dictionary<string, ComboBox> _assistantTypes = [];
    private StackPanel? _hotkeyPagePanel;
    private ScrollViewer? _hotkeyPageScroll;
    private readonly Dictionary<string, string> _audioDeviceIdsByDisplayName = [];
    private readonly Dictionary<string, UIElement> _pages = [];
    private static readonly IReadOnlyDictionary<InsertMethod, string> InsertMethodLabels = new Dictionary<InsertMethod, string>
    {
        [InsertMethod.Clipboard] = "über die Zwischenablage einfügen",
        [InsertMethod.SendInput] = "direkt tippen (zeichenweise)"
    };
    private static readonly IReadOnlyDictionary<WritingStyle, string> WritingStyleLabels = new Dictionary<WritingStyle, string>
    {
        [WritingStyle.Casual] = "locker",
        [WritingStyle.Neutral] = "neutral",
        [WritingStyle.Professional] = "professionell",
        [WritingStyle.Academic] = "wissenschaftlich"
    };
    private static readonly IReadOnlyDictionary<ParagraphDensity, string> ParagraphDensityLabels = new Dictionary<ParagraphDensity, string>
    {
        [ParagraphDensity.Compact] = "kompakt (wenig Absätze)",
        [ParagraphDensity.Balanced] = "ausgewogen",
        [ParagraphDensity.Spacious] = "luftig (mehr Absätze)"
    };
    private static readonly IReadOnlyDictionary<EmojiExpression, string> EmojiExpressionLabels = new Dictionary<EmojiExpression, string>
    {
        [EmojiExpression.None] = "keine",
        [EmojiExpression.Sparse] = "sparsam",
        [EmojiExpression.Balanced] = "mittel",
        [EmojiExpression.Lively] = "lebhaft (Social)",
        [EmojiExpression.Heavy] = "sehr viel (Social Media)"
    };
    private AppSettings _settings = new();
    private string? _capturingAssistantId;
    private bool _windowConfigured;
    private bool _isExiting;
    private bool _initialNavigationDone;
    private bool _autoPersistWired;
    private SubclassProc? _minimizeSubclassProc;

    private readonly IAppProfile _profile;

    public MainWindow(
        IAppProfile profile,
        ISettingsService settingsService,
        ISecretProtector secretProtector,
        ITrayStatusService statusService,
        IHotkeyService hotkeyService,
        IAutostartService autostartService,
        IAudioDeviceService audioDeviceService,
        ILlmService llmService,
        FileLogger fileLogger,
        IProcessingFailureLog processingFailureLog)
    {
        Root = new Grid();
        Content = Root;
        _profile = profile;
        _settingsService = settingsService;
        _secretProtector = secretProtector;
        _statusService = statusService;
        _hotkeyService = hotkeyService;
        _autostartService = autostartService;
        _audioDeviceService = audioDeviceService;
        _llmService = llmService;
        _fileLogger = fileLogger;
        _processingFailureLog = processingFailureLog;

        Title = $"{_profile.AppName} {GetAppVersion()}";
        PopulateLists();
        BuildShell();
        _statusService.StatusChanged += (_, args) => DispatcherQueue.TryEnqueue(() => UpdateInfoBar(args.Status, args.Message));
    }

    public async Task InitializeAfterActivationAsync()
    {
        ConfigureWindow();
        await LoadSettingsAsync();
    }

    public void HideToTray()
    {
        _ = SaveWindowBoundsAsync();
        ShowWindow(WindowNative.GetWindowHandle(this), SwHide);
    }

    public void ShowFromTray()
    {
        // Falls das Fenster beim Start off-screen platziert wurde (siehe App.OnLaunched), zurück
        // an die gespeicherte oder eine sinnvolle Default-Position holen, bevor wir es zeigen.
        ApplyWindowBounds();
        var hwnd = WindowNative.GetWindowHandle(this);
        // Falls zuvor per SW_HIDE versteckt: erst wieder sichtbar machen, dann aktivieren.
        ShowWindow(hwnd, SwRestore);
        Activate();
    }

    public void PrepareForExit() => _isExiting = true;

    private async Task ShutdownFromCloseAsync()
    {
        try
        {
            await PersistAsync();
        }
        catch
        {
            // Beim Schließen darf ein Speicherfehler die App nicht am Beenden hindern.
        }

        if (Microsoft.UI.Xaml.Application.Current is App app)
        {
            await app.ShutdownAsync();
        }

        Microsoft.UI.Xaml.Application.Current.Exit();
    }

    public async Task SaveWindowBoundsAsync()
    {
        var appWindow = CurrentAppWindow();
        _settings.WindowBounds.Width = Math.Max(MinimumWindowWidth, appWindow.Size.Width);
        _settings.WindowBounds.Height = Math.Max(MinimumWindowHeight, appWindow.Size.Height);
        _settings.WindowBounds.X = appWindow.Position.X;
        _settings.WindowBounds.Y = appWindow.Position.Y;
        await _settingsService.SaveAsync(_settings);
    }

    public void RefreshDiagnostics()
    {
        var readiness = _settingsService.Validate(_settings);
        var exePath = Environment.ProcessPath;
        _diagnostics.Text =
            $"App-Version: {GetAppVersion()}{Environment.NewLine}" +
            $"App-Version (Assembly): {GetAssemblyInformationalVersion()}{Environment.NewLine}" +
            $"App-Version (Exe/Product): {GetExeProductVersion()}{Environment.NewLine}" +
            $"Exe: {exePath ?? "unbekannt"}{Environment.NewLine}" +
            $".NET-Laufzeit: {Environment.Version}{Environment.NewLine}" +
            $"App-Datenpfad: {_settingsService.DataDirectory}{Environment.NewLine}" +
            $"Protokolle: {_settingsService.LogDirectory}{Environment.NewLine}" +
            $"Transkriptionsanbieter: {_settings.SttProvider}{Environment.NewLine}" +
            $"Transkriptionsmodell: {_settings.SttModel}{Environment.NewLine}" +
            $"Audioquelle: {AudioDeviceName(_settings.AudioInputDeviceId)}{Environment.NewLine}" +
            $"Eingabesprache: {Defaults.LanguageName(_settings.InputLanguage)}{Environment.NewLine}" +
            $"Ausgabesprache: {Defaults.LanguageName(_settings.OutputLanguage)}{Environment.NewLine}" +
            $"KI-Anbieter: {_settings.LlmProvider}{Environment.NewLine}" +
            $"KI-Modell: {_settings.LlmModel}{Environment.NewLine}" +
            $"Tastenkürzel-Status: {(readiness.Issues.Any(issue => issue.Field.StartsWith("hotkey", StringComparison.OrdinalIgnoreCase)) ? "prüfen" : "gültig")}{Environment.NewLine}" +
            $"Einrichtungsstatus: {(readiness.IsReady ? "Bereit" : "Einrichtung erforderlich")}{Environment.NewLine}" +
            $"Status: {FriendlyStatus(_statusService.CurrentStatus)} - {SanitizeStatus(_statusService.Message)}{Environment.NewLine}" +
            $"Hinweis: Einfügen in erhöhte Zielanwendungen kann scheitern, wenn Magic-Voice nicht ebenfalls erhöht läuft.{Environment.NewLine}" +
            string.Join(Environment.NewLine, readiness.Issues.Select(issue => $"- {issue.Message}")) +
            LastProcessingFailureSection();
    }

    private string LastProcessingFailureSection()
    {
        var entry = _processingFailureLog.LastEntry;
        if (string.IsNullOrWhiteSpace(entry))
        {
            return string.Empty;
        }

        return $"{Environment.NewLine}{Environment.NewLine}Letzte Verarbeitung (Fehler):{Environment.NewLine}{entry}";
    }

    private void ConfigureWindow()
    {
        if (_windowConfigured)
        {
            return;
        }

        _windowConfigured = true;
        var appWindow = CurrentAppWindow();
        appWindow.Resize(new SizeInt32(DefaultWindowWidth, DefaultWindowHeight));
        appWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "App.ico"));
        // AppWindow.Closing: Schließen über das X beendet immer. Cancel synchron setzen wäre nötig,
        // wenn wir abbrechen wollten – tun wir nicht mehr.
        appWindow.Closing += (sender, args) =>
        {
            if (_isExiting)
            {
                return;
            }

            args.Cancel = true;
            _ = ShutdownFromCloseAsync();
        };

        InstallMinimizeHook();

        if (appWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.IsResizable = true;
            presenter.IsMaximizable = true;
            presenter.PreferredMinimumWidth = MinimumWindowWidth;
            presenter.PreferredMinimumHeight = MinimumWindowHeight;
        }
    }

    /// <summary>
    /// Subclassing der WinUI-Window-Procedure, um SC_MINIMIZE abzufangen und – falls in den Einstellungen
    /// gewünscht – statt einer regulären Minimierung das Fenster ins Tray zu verbergen.
    /// </summary>
    private void InstallMinimizeHook()
    {
        var hwnd = WindowNative.GetWindowHandle(this);
        // Delegate als Field halten, damit der GC ihn nicht einsammelt, solange das Fenster lebt.
        _minimizeSubclassProc = MinimizeSubclassProc;
        SetWindowSubclass(hwnd, _minimizeSubclassProc, new UIntPtr(1), UIntPtr.Zero);
    }

    private IntPtr MinimizeSubclassProc(IntPtr hWnd, uint uMsg, UIntPtr wParam, IntPtr lParam, UIntPtr uIdSubclass, UIntPtr dwRefData)
    {
        const uint WM_SYSCOMMAND = 0x0112;
        const uint SC_MINIMIZE = 0xF020;
        const uint SC_MASK = 0xFFF0;

        if (uMsg == WM_SYSCOMMAND && (wParam.ToUInt32() & SC_MASK) == SC_MINIMIZE && _settings.MinimizeToTray)
        {
            HideToTray();
            return IntPtr.Zero;
        }

        return DefSubclassProc(hWnd, uMsg, wParam, lParam);
    }

    private AppWindow CurrentAppWindow()
    {
        var hwnd = WindowNative.GetWindowHandle(this);
        return AppWindow.GetFromWindowId(Win32Interop.GetWindowIdFromWindow(hwnd));
    }

    private void ApplyWindowBounds()
    {
        var appWindow = CurrentAppWindow();
        var width = Math.Max(DefaultWindowWidth, (int)_settings.WindowBounds.Width);
        var height = Math.Max(DefaultWindowHeight, (int)_settings.WindowBounds.Height);
        appWindow.Resize(new SizeInt32(width, height));
        if (_settings.WindowBounds.IsSet)
        {
            var desiredX = (int)_settings.WindowBounds.X!.Value;
            var desiredY = (int)_settings.WindowBounds.Y!.Value;
            var desiredRect = new RectInt32(desiredX, desiredY, width, height);

            // Koordinaten können nach Monitorwechsel/DPI/WorkArea-Änderung off-screen werden.
            // Wir prüfen gegen die WorkArea des nächstgelegenen Monitors und erzwingen, dass
            // mindestens ein sinnvoller Teil sichtbar ist (mind. Titelzeile/Grip-Bereich).
            var workArea = MonitorWorkAreaFor(desiredRect);
            if (!IsRectSufficientlyVisible(desiredRect, workArea))
            {
                var centered = CenteredRectWithin(workArea, width, height);
                appWindow.Move(new PointInt32(centered.X, centered.Y));
            }
            else
            {
                var clamped = ClampRectIntoWorkArea(desiredRect, workArea);
                appWindow.Move(new PointInt32(clamped.X, clamped.Y));
            }
        }
        else
        {
            var workArea = DisplayArea.GetFromWindowId(appWindow.Id, DisplayAreaFallback.Primary).WorkArea;
            var x = workArea.X + Math.Max(0, (workArea.Width - width) / 2);
            var y = workArea.Y + Math.Max(0, (workArea.Height - height) / 2);
            appWindow.Move(new PointInt32(x, y));
        }
    }

    private static RectInt32 CenteredRectWithin(RectInt32 workArea, int width, int height)
    {
        var w = Math.Min(width, workArea.Width);
        var h = Math.Min(height, workArea.Height);
        var x = workArea.X + Math.Max(0, (workArea.Width - w) / 2);
        var y = workArea.Y + Math.Max(0, (workArea.Height - h) / 2);
        return new RectInt32(x, y, w, h);
    }

    private static RectInt32 ClampRectIntoWorkArea(RectInt32 rect, RectInt32 workArea)
    {
        // Wir wollen mindestens ein Stück (Titlebar/Grip) sichtbar halten.
        const int minVisibleX = 80;
        const int minVisibleY = 40;

        var maxX = workArea.X + Math.Max(0, workArea.Width - minVisibleX);
        var maxY = workArea.Y + Math.Max(0, workArea.Height - minVisibleY);
        var minX = workArea.X - Math.Max(0, rect.Width - minVisibleX);
        var minY = workArea.Y;

        var x = Math.Clamp(rect.X, minX, maxX);
        var y = Math.Clamp(rect.Y, minY, maxY);
        return new RectInt32(x, y, rect.Width, rect.Height);
    }

    private static bool IsRectSufficientlyVisible(RectInt32 rect, RectInt32 workArea)
    {
        var intersect = Intersect(rect, workArea);
        // Mindestens ein brauchbarer sichtbarer Bereich, sonst neu zentrieren.
        return intersect.Width >= 120 && intersect.Height >= 80;
    }

    private static RectInt32 Intersect(RectInt32 a, RectInt32 b)
    {
        var left = Math.Max(a.X, b.X);
        var top = Math.Max(a.Y, b.Y);
        var right = Math.Min(a.X + a.Width, b.X + b.Width);
        var bottom = Math.Min(a.Y + a.Height, b.Y + b.Height);
        var width = Math.Max(0, right - left);
        var height = Math.Max(0, bottom - top);
        return new RectInt32(left, top, width, height);
    }

    private static RectInt32 MonitorWorkAreaFor(RectInt32 rect)
    {
        var r = new Rect
        {
            Left = rect.X,
            Top = rect.Y,
            Right = rect.X + rect.Width,
            Bottom = rect.Y + rect.Height
        };
        var monitor = MonitorFromRect(ref r, MonitorDefaultToNearest);
        var info = new MonitorInfo { cbSize = Marshal.SizeOf<MonitorInfo>() };
        if (monitor != IntPtr.Zero && GetMonitorInfo(monitor, ref info))
        {
            var wa = info.rcWork;
            return new RectInt32(wa.Left, wa.Top, wa.Right - wa.Left, wa.Bottom - wa.Top);
        }

        // Fallback: Primary work area via DisplayArea.
        var primary = DisplayArea.GetFromPoint(new PointInt32(0, 0), DisplayAreaFallback.Primary).WorkArea;
        return new RectInt32(primary.X, primary.Y, primary.Width, primary.Height);
    }

    private const int SwHide = 0;
    private const int SwRestore = 9;
    private const uint MonitorDefaultToNearest = 2;

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MonitorInfo
    {
        public int cbSize;
        public Rect rcMonitor;
        public Rect rcWork;
        public uint dwFlags;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromRect(ref Rect lprc, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MonitorInfo lpmi);

    private void PopulateLists()
    {
        AddItems(_sttProvider, Defaults.KnownProviders);
        AddItems(_llmProvider, Defaults.KnownProviders);
        AddItems(_sttModel, Defaults.OpenAiSttModels);
        AddItems(_llmModel, Defaults.OpenAiLlmModels);
        AddItems(_inputLanguage, Defaults.InputLanguages.Select(language => language.Name));
        AddItems(_outputLanguage, Defaults.OutputLanguages.Select(language => language.Name));
        AddItems(_insertMethod, InsertMethodLabels.Values);
        RefreshAudioDeviceList();
        _sttProvider.IsEnabled = false;
        _llmProvider.IsEnabled = false;
    }

    private void BuildShell()
    {
        Root.Background = ResourceBrush("ApplicationPageBackgroundThemeBrush");
        Root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        Root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        Root.Padding = new Thickness(16);
        Root.RowSpacing = 12;
        Root.AddHandler(UIElement.KeyDownEvent, new KeyEventHandler(OnRootKeyDown), true);

        Grid.SetRow(_statusInfo, 0);
        Root.Children.Add(_statusInfo);

        _navigation.Header = null;
        _navigation.OpenPaneLength = 188;
        _navigation.CompactPaneLength = 44;
        _navigation.IsPaneToggleButtonVisible = false;
        StripNavigationViewContentChrome(_navigation);
        _navigation.MenuItems.Add(NavItem("Übersicht", OverviewPage, "\uE80F"));
        _navigation.MenuItems.Add(NavItem("Anbieter", ProviderPage, "\uE774"));
        _navigation.MenuItems.Add(NavItem("Assistenten", HotkeyPage, "\uE765"));
        _navigation.MenuItems.Add(NavItem("Audio & Sprache", AudioLanguagePage, "\uE8D6"));
        _navigation.MenuItems.Add(NavItem("Allgemein", GeneralPage, "\uE713"));
        _navigation.MenuItems.Add(NavItem("Einfügen", InsertPage, "\uE77F"));
        _navigation.MenuItems.Add(NavItem("Diagnose", DiagnosticsPage, "\uE9D9"));
        _navigation.MenuItems.Add(NavItem("Über", AboutPage, "\uE946"));
        _navigation.SelectionChanged += OnNavigationSelectionChanged;
        Grid.SetRow(_navigation, 1);
        Root.Children.Add(_navigation);

        _navigation.SelectedItem = _navigation.MenuItems[0];
        Navigate(OverviewPage);
    }

    private void Navigate(string tag)
    {
        _capturingAssistantId = null;
        ResetCaptureButtons();
        if (_initialNavigationDone)
        {
            // UI-Werte zuerst in _settings übernehmen, damit z.B. die Diagnose-Seite
            // die aktuellen Einstellungen anzeigt (auch ohne den I/O-Persist abzuwarten).
            SaveSettingsFromFields();
        }
        _settings.LastSelectedSettingsSection = tag;
        _navigation.Content = null;
        try
        {
            if (tag == DiagnosticsPage)
            {
                RefreshDiagnostics();
            }

            if (!_pages.TryGetValue(tag, out var page))
            {
                page = tag switch
                {
                    AudioLanguagePage => BuildAudioLanguagePage(),
                    ProviderPage => BuildProviderPage(),
                    HotkeyPage => BuildHotkeyPage(),
                    InsertPage => BuildInsertPage(),
                    GeneralPage => BuildGeneralPage(),
                    DiagnosticsPage => BuildDiagnosticsPage(),
                    AboutPage => BuildAboutPage(),
                    _ => BuildOverviewPage()
                };
                _pages[tag] = page;
            }

            _navigation.Content = page;
        }
        catch (Exception ex)
        {
            MagicVoiceStartupCrashLogger.Write(ex);
            _navigation.Content = BuildNavigationErrorPage(tag, ex);
            _statusService.SetStatus(TrayStatus.Error, $"Seite „{TitleFor(tag)}“ konnte nicht geöffnet werden.");
        }
    }

    private void ClearPageCache()
    {
        _navigation.Content = null;
        _pages.Clear();
    }

    private UIElement BuildOverviewPage()
    {
        var panel = PageStack();
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Bereit zum Diktieren"),
                Body("Halte ein Tastenkürzel gedrückt, sprich deinen Text und lasse los. Der Assistent transkribiert, verarbeitet und fügt den finalen Text ein."),
                ActionRow(
                    Button("Verbindung prüfen", async () => await TestLlmConnectionAsync()))
            }
        }));
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Testfeld"),
                Body("Hier kannst du das Einfügen ausprobieren: Klicke ins Feld, halte ein Assistenten-Tastenkürzel gedrückt und sprich. Der finale Text wird in dieses Feld eingefügt."),
                _insertionTestTarget
            }
        }));
        return Page(panel);
    }

    private UIElement BuildNavigationErrorPage(string tag, Exception exception)
    {
        var panel = PageStack();
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Seite konnte nicht geöffnet werden"),
                Body($"Beim Öffnen von „{TitleFor(tag)}“ ist ein Fehler aufgetreten."),
                Body(exception.Message),
                ActionRow(Button("Diagnose öffnen", () => Navigate(DiagnosticsPage)))
            }
        }));
        return Page(panel);
    }

    private UIElement BuildProviderPage()
    {
        var panel = PageStack();
        panel.Children.Add(Card(Form("Transkription", "Aktuell wird OpenAI unterstützt.",
            _sttProvider,
            _sttModel,
            _sttApiKey)));
        panel.Children.Add(Card(Form("KI-Verarbeitung", "Aktuell wird OpenAI unterstützt.",
            _llmProvider,
            _llmModel,
            _llmApiKey)));
        return Page(panel);
    }

    private UIElement BuildAudioLanguagePage()
    {
        var panel = PageStack();
        panel.Children.Add(Card(Form("Audioeingabe", "Wähle das Mikrofon, das beim Sprechen bei gedrückter Taste verwendet wird.",
            _audioInputDevice,
            ActionRow(Button("Geräte aktualisieren", RefreshAudioDevices)))));
        panel.Children.Add(Card(Form("Sprache", "Damit kann der Assistent zugleich übersetzen: Eingabe erkennen lassen und Ausgabe gezielt wählen.",
            _inputLanguage,
            _outputLanguage)));
        return Page(panel);
    }

    private UIElement BuildHotkeyPage()
    {
        var panel = PageStack();
        panel.Children.Add(Body("Du kannst beliebig viele Assistenten anlegen. Der Typ legt nur fest, ob ein Transkript, nur eine Anweisung oder die Zwischenablage als Quelle dient; Stil, Intensität und Absätze steuerst du in der jeweiligen Karte."));

        var listPanel = new StackPanel { Spacing = 16 };
        _hotkeyPagePanel = listPanel;
        RenderAssistantCards(listPanel);
        panel.Children.Add(listPanel);

        var typePicker = new ComboBox
        {
            Header = "Typ des neuen Assistenten",
            MinWidth = 240,
            HorizontalAlignment = HorizontalAlignment.Stretch
        };
        foreach (var mode in _profile.Modes)
        {
            typePicker.Items.Add(mode.Name);
        }
        if (typePicker.Items.Count > 0)
        {
            typePicker.SelectedIndex = 0;
        }

        var addButton = Button("+ Assistent hinzufügen", () =>
        {
            if (typePicker.SelectedIndex < 0 || typePicker.SelectedIndex >= _profile.Modes.Count)
            {
                return;
            }

            var template = _profile.Modes[typePicker.SelectedIndex];
            var newAssistant = new AssistantInstance
            {
                Id = Guid.NewGuid().ToString("N"),
                Type = template.Mode,
                Name = template.Name,
                Hotkey = string.Empty,
                Prompt = template.DefaultPrompt,
                Intensity = Defaults.DefaultModeIntensity,
                WritingStyle = WritingStyle.Neutral
            };
            SaveSettingsFromFields();
            _settings.Assistants.Add(newAssistant);
            RebuildHotkeyPage();
            ScrollToAssistant(newAssistant.Id);
        });
        addButton.VerticalAlignment = VerticalAlignment.Bottom;

        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Neuen Assistenten anlegen"),
                new Grid
                {
                    ColumnDefinitions =
                    {
                        new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                        new ColumnDefinition { Width = GridLength.Auto }
                    },
                    ColumnSpacing = 12,
                    Children =
                    {
                        typePicker,
                        WithColumn(addButton, 1)
                    }
                }
            }
        }));

        var page = Page(panel);
        _hotkeyPageScroll = page as ScrollViewer;
        return page;
    }

    private void ScrollToAssistant(string assistantId)
    {
        if (_hotkeyPagePanel is null)
        {
            return;
        }

        if (!_names.TryGetValue(assistantId, out var anchor))
        {
            return;
        }

        DispatcherQueue.TryEnqueue(() => anchor.StartBringIntoView(new BringIntoViewOptions
        {
            AnimationDesired = true,
            VerticalAlignmentRatio = 0.0
        }));
    }

    private void RebuildHotkeyPage()
    {
        if (_hotkeyPagePanel is null)
        {
            return;
        }

        RenderAssistantCards(_hotkeyPagePanel);
    }

    private void MoveAssistantRelative(string assistantId, int delta)
    {
        var list = _settings.Assistants;
        var index = list.FindIndex(a => string.Equals(a.Id, assistantId, StringComparison.Ordinal));
        if (index < 0)
        {
            return;
        }

        var newIndex = index + delta;
        if (newIndex < 0 || newIndex >= list.Count)
        {
            return;
        }

        SaveSettingsFromFields();
        (list[index], list[newIndex]) = (list[newIndex], list[index]);
        RebuildHotkeyPage();
        ScrollToAssistant(assistantId);
        _ = AutoPersistAsync();
    }

    private void RenderAssistantCards(StackPanel listPanel)
    {
        listPanel.Children.Clear();
        _hotkeyValues.Clear();
        _prompts.Clear();
        _names.Clear();
        _systemPromptOverrideCheckBoxes.Clear();
        _systemPrompts.Clear();
        _assistantInputLanguageOverrides.Clear();
        _assistantOutputLanguageOverrides.Clear();
        _policyTooltips.Clear();
        _intensitySliders.Clear();
        _intensityValues.Clear();
        _captureButtons.Clear();
        _writingStyles.Clear();
        _paragraphDensities.Clear();
        _emojiExpressions.Clear();
        _assistantTypes.Clear();

        if (_settings.Assistants.Count == 0)
        {
            listPanel.Children.Add(Body("Noch keine Assistenten vorhanden. Lege oben einen an."));
            return;
        }

        foreach (var assistant in _settings.Assistants.ToList())
        {
            listPanel.Children.Add(BuildAssistantCard(assistant));
        }
    }

    private UIElement BuildAssistantCard(AssistantInstance assistant)
    {
        var template = _profile.Modes.FirstOrDefault(m => m.Mode == assistant.Type);
        var typeLabel = template?.Name ?? assistant.Type.ToString();
        var assistantListIndex = _settings.Assistants.FindIndex(a =>
            string.Equals(a.Id, assistant.Id, StringComparison.Ordinal));

        var nameBox = new TextBox
        {
            Header = "Anzeigename",
            Text = assistant.Name,
            HorizontalAlignment = HorizontalAlignment.Stretch
        };
        nameBox.LostFocus += (_, _) => _ = AutoPersistAsync();
        _names[assistant.Id] = nameBox;

        var typeCombo = new ComboBox
        {
            Header = "Assistenten-Typ",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            MinWidth = 260
        };
        foreach (var mode in _profile.Modes)
        {
            typeCombo.Items.Add(mode.Name);
        }
        typeCombo.SelectedItem = template?.Name ?? _profile.Modes[0].Name;
        _assistantTypes[assistant.Id] = typeCombo;
        var assistantIdForType = assistant.Id;
        typeCombo.SelectionChanged += (_, _) =>
        {
            if (typeCombo.SelectedItem is not string selectedName)
            {
                return;
            }

            var modeDef = _profile.Modes.FirstOrDefault(m =>
                string.Equals(m.Name, selectedName, StringComparison.OrdinalIgnoreCase));
            if (modeDef is null)
            {
                return;
            }

            var stored = _settings.Assistants.FirstOrDefault(a => string.Equals(a.Id, assistantIdForType, StringComparison.Ordinal));
            if (stored is null || stored.Type == modeDef.Mode)
            {
                return;
            }

            SaveSettingsFromFields();
            RebuildHotkeyPage();
            ScrollToAssistant(assistantIdForType);
            _ = AutoPersistAsync();
        };

        var hotkeyValue = new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(assistant.Hotkey) ? "(kein Tastenkürzel)" : assistant.Hotkey,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        };
        _hotkeyValues[assistant.Id] = hotkeyValue;

        var captureButton = Button("Tastenkürzel aufnehmen", () => StartHotkeyCapture(assistant.Id));
        _captureButtons[assistant.Id] = captureButton;

        var initialIntensity = Math.Clamp(assistant.Intensity, Defaults.MinModeIntensity, Defaults.MaxModeIntensity);
        var intensityValue = new TextBlock
        {
            Text = FormatIntensityValue(assistant.Type, initialIntensity),
            VerticalAlignment = VerticalAlignment.Center,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            MinWidth = 170,
            TextWrapping = TextWrapping.Wrap,
            TextAlignment = TextAlignment.Right
        };
        var intensity = new Slider
        {
            Minimum = Defaults.MinModeIntensity,
            Maximum = Defaults.MaxModeIntensity,
            StepFrequency = 1,
            Value = initialIntensity,
            SmallChange = 1,
            LargeChange = 1,
            TickFrequency = 1,
            TickPlacement = Microsoft.UI.Xaml.Controls.Primitives.TickPlacement.BottomRight,
            SnapsTo = Microsoft.UI.Xaml.Controls.Primitives.SliderSnapsTo.StepValues,
            HorizontalAlignment = HorizontalAlignment.Stretch
        };
        var capturedType = assistant.Type;
        intensity.ValueChanged += (_, args) =>
        {
            intensityValue.Text = FormatIntensityValue(capturedType, (int)Math.Round(args.NewValue));
            _ = AutoPersistAsync();
        };
        _intensitySliders[assistant.Id] = intensity;
        _intensityValues[assistant.Id] = intensityValue;

        var prompt = new TextBox
        {
            Header = "Anweisung",
            Text = assistant.Prompt,
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            MinHeight = 120
        };
        prompt.LostFocus += (_, _) => _ = AutoPersistAsync();
        _prompts[assistant.Id] = prompt;

        // Sprachen (Overrides)
        var inputOverride = new ComboBox
        {
            Header = "Eingabesprache (Assistent)",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            MinWidth = 260
        };
        inputOverride.Items.Add(GlobalStandardLabel);
        foreach (var language in Defaults.InputLanguages.Select(l => l.Name))
        {
            inputOverride.Items.Add(language);
        }
        inputOverride.SelectedItem = string.IsNullOrWhiteSpace(assistant.InputLanguageOverride)
            ? GlobalStandardLabel
            : Defaults.LanguageName(assistant.InputLanguageOverride);
        inputOverride.SelectionChanged += (_, _) => _ = AutoPersistAsync();
        _assistantInputLanguageOverrides[assistant.Id] = inputOverride;

        var outputOverride = new ComboBox
        {
            Header = "Ausgabesprache (Assistent)",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            MinWidth = 260
        };
        outputOverride.Items.Add(GlobalStandardLabel);
        foreach (var language in Defaults.OutputLanguages.Select(l => l.Name))
        {
            outputOverride.Items.Add(language);
        }
        outputOverride.SelectedItem = string.IsNullOrWhiteSpace(assistant.OutputLanguageOverride)
            ? GlobalStandardLabel
            : Defaults.LanguageName(assistant.OutputLanguageOverride);
        outputOverride.SelectionChanged += (_, _) => _ = AutoPersistAsync();
        _assistantOutputLanguageOverrides[assistant.Id] = outputOverride;

        // System-Prompt (Override)
        var systemPromptOverrideCheck = new CheckBox
        {
            Content = "benutzerdefinierter System-Prompt",
            IsChecked = !string.IsNullOrWhiteSpace(assistant.SystemPromptOverride),
            HorizontalAlignment = HorizontalAlignment.Stretch
        };
        _systemPromptOverrideCheckBoxes[assistant.Id] = systemPromptOverrideCheck;

        var systemPrompt = new TextBox
        {
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            MinHeight = 96,
            Text = string.IsNullOrWhiteSpace(assistant.SystemPromptOverride) ? _profile.SystemPrompt : assistant.SystemPromptOverride
        };
        systemPrompt.IsEnabled = systemPromptOverrideCheck.IsChecked == true;
        _systemPrompts[assistant.Id] = systemPrompt;

        void OnSystemPromptOverrideCheckChanged()
        {
            var useCustom = systemPromptOverrideCheck.IsChecked == true;
            systemPrompt.IsEnabled = useCustom;
            // Copy-on-enable: wenn leer, Standard übernehmen, damit das Feld nicht leer startet.
            if (useCustom && string.IsNullOrWhiteSpace(systemPrompt.Text))
            {
                systemPrompt.Text = _profile.SystemPrompt;
            }
            _ = AutoPersistAsync();
        }

        systemPromptOverrideCheck.Checked += (_, _) => OnSystemPromptOverrideCheckChanged();
        systemPromptOverrideCheck.Unchecked += (_, _) => OnSystemPromptOverrideCheckChanged();
        systemPrompt.LostFocus += (_, _) => _ = AutoPersistAsync();

        var assistantIdForDelete = assistant.Id;
        var deleteButton = CompactDeleteAssistantButton(async () =>
        {
            if (_settings.Assistants.Count <= 1)
            {
                return;
            }

            var displayName = string.IsNullOrWhiteSpace(nameBox.Text) ? assistant.Name : nameBox.Text.Trim();
            if (!await ConfirmDeleteAssistantAsync(displayName))
            {
                return;
            }

            SaveSettingsFromFields();
            _settings.Assistants.RemoveAll(a => a.Id == assistantIdForDelete);
            RebuildHotkeyPage();
            _ = AutoPersistAsync();
        });
        deleteButton.IsEnabled = _settings.Assistants.Count > 1;

        var moveUpButton = CompactArrowReorderButton(up: true, "Nach oben verschieben", () => MoveAssistantRelative(assistant.Id, -1));
        moveUpButton.IsEnabled = assistantListIndex > 0;
        var moveDownButton = CompactArrowReorderButton(up: false, "Nach unten verschieben", () => MoveAssistantRelative(assistant.Id, 1));
        moveDownButton.IsEnabled = assistantListIndex >= 0 && assistantListIndex < _settings.Assistants.Count - 1;

        var cardHeaderActions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Top,
            Children = { moveUpButton, moveDownButton, deleteButton }
        };

        var card = new StackPanel
        {
            Spacing = 10,
            Children =
            {
                new Grid
                {
                    ColumnDefinitions =
                    {
                        new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                        new ColumnDefinition { Width = GridLength.Auto }
                    },
                    ColumnSpacing = 12,
                    Children =
                    {
                        new StackPanel
                        {
                            Spacing = 4,
                            Children =
                            {
                                Header($"{typeLabel} – Assistent"),
                                Body(template?.Description ?? string.Empty)
                            }
                        },
                        WithColumn(cardHeaderActions, 1)
                    }
                },
                nameBox,
                typeCombo,
                new Grid
                {
                    ColumnDefinitions =
                    {
                        new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                        new ColumnDefinition { Width = GridLength.Auto }
                    },
                    ColumnSpacing = 12,
                    Children =
                    {
                        hotkeyValue,
                        WithColumn(captureButton, 1)
                    }
                },
                new Grid
                {
                    ColumnDefinitions =
                    {
                        new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                        new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }
                    },
                    ColumnSpacing = 12,
                    Children =
                    {
                        inputOverride,
                        WithColumn(outputOverride, 1)
                    }
                },
                systemPromptOverrideCheck,
                systemPrompt,
                new Grid
                {
                    ColumnDefinitions =
                    {
                        new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) },
                        new ColumnDefinition { Width = GridLength.Auto }
                    },
                    ColumnSpacing = 12,
                    Children =
                    {
                        BuildIntensityWithPolicyTooltip(assistant, intensity),
                        WithColumn(intensityValue, 1)
                    }
                }
            }
        };

        if (ProfileSupportsWritingStyle())
        {
            var styleCombo = new ComboBox
            {
                Header = "Schreibstil",
                HorizontalAlignment = HorizontalAlignment.Stretch,
                MinWidth = 260
            };
            foreach (var label in WritingStyleLabels.Values)
            {
                styleCombo.Items.Add(label);
            }
            styleCombo.SelectedItem = WritingStyleLabels[assistant.WritingStyle];
            _writingStyles[assistant.Id] = styleCombo;
            styleCombo.SelectionChanged += (_, _) =>
            {
                if (_policyTooltips.TryGetValue(assistant.Id, out var t))
                {
                    t.Content = BuildPolicyPreviewText(assistant.Id);
                }
                _ = AutoPersistAsync();
            };
            card.Children.Add(styleCombo);
        }

        var paragraphCombo = new ComboBox
        {
            Header = "Absätze und Leerzeilen",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            MinWidth = 260
        };
        foreach (var label in ParagraphDensityLabels.Values)
        {
            paragraphCombo.Items.Add(label);
        }
        paragraphCombo.SelectedItem = ParagraphDensityLabels[assistant.ParagraphDensity];
        _paragraphDensities[assistant.Id] = paragraphCombo;
        paragraphCombo.SelectionChanged += (_, _) =>
        {
            if (_policyTooltips.TryGetValue(assistant.Id, out var t))
            {
                t.Content = BuildPolicyPreviewText(assistant.Id);
            }
            _ = AutoPersistAsync();
        };
        card.Children.Add(paragraphCombo);

        if (ProfileSupportsEmojiExpression())
        {
            var emojiCombo = new ComboBox
            {
                Header = "Emojis und Social-Ton",
                HorizontalAlignment = HorizontalAlignment.Stretch,
                MinWidth = 260
            };
            foreach (var label in EmojiExpressionLabels.Values)
            {
                emojiCombo.Items.Add(label);
            }

            emojiCombo.SelectedItem = EmojiExpressionLabels[assistant.EmojiExpression];
            _emojiExpressions[assistant.Id] = emojiCombo;
            emojiCombo.SelectionChanged += (_, _) =>
            {
                if (_policyTooltips.TryGetValue(assistant.Id, out var t))
                {
                    t.Content = BuildPolicyPreviewText(assistant.Id);
                }
                _ = AutoPersistAsync();
            };
            card.Children.Add(emojiCombo);
        }

        card.Children.Add(prompt);
        return Card(card);
    }

    private UIElement BuildIntensityWithPolicyTooltip(AssistantInstance assistant, Slider intensitySlider)
    {
        var label = new TextBlock
        {
            Text = IntensityLabel(assistant.Type),
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        };
        // SymbolIcon skaliert in StackPanels sonst oft falsch (riesiges Fragezeichen) — feste Kachel.
        var helpGlyph = new SymbolIcon(Symbol.Help);
        var info = new Viewbox
        {
            Width = 16,
            Height = 16,
            Margin = new Thickness(6, 0, 0, 0),
            Stretch = Stretch.Uniform,
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Center,
            Child = helpGlyph
        };
        var tooltip = new ToolTip { Content = BuildPolicyPreviewText(assistant.Id) };
        _policyTooltips[assistant.Id] = tooltip;
        ToolTipService.SetToolTip(info, tooltip);

        void UpdateTooltip()
        {
            if (_policyTooltips.TryGetValue(assistant.Id, out var t))
            {
                t.Content = BuildPolicyPreviewText(assistant.Id);
            }
        }

        intensitySlider.ValueChanged += (_, _) => UpdateTooltip();
        if (_assistantInputLanguageOverrides.TryGetValue(assistant.Id, out var inCombo)) inCombo.SelectionChanged += (_, _) => UpdateTooltip();
        if (_assistantOutputLanguageOverrides.TryGetValue(assistant.Id, out var outCombo)) outCombo.SelectionChanged += (_, _) => UpdateTooltip();

        var headerRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Left,
            Children = { label, info }
        };

        return new StackPanel
        {
            Spacing = 6,
            Children =
            {
                headerRow,
                intensitySlider
            }
        };
    }

    private string BuildPolicyPreviewText(string assistantId)
    {
        var assistant = _settings.Assistants.FirstOrDefault(a => string.Equals(a.Id, assistantId, StringComparison.Ordinal));
        if (assistant is null)
        {
            return string.Empty;
        }

        var effective = new AssistantInstance
        {
            Id = assistant.Id,
            Type = assistant.Type,
            Name = assistant.Name,
            Hotkey = assistant.Hotkey,
            Prompt = assistant.Prompt,
            Intensity = assistant.Intensity,
            WritingStyle = assistant.WritingStyle,
            ParagraphDensity = assistant.ParagraphDensity,
            EmojiExpression = assistant.EmojiExpression,
            SystemPromptOverride = assistant.SystemPromptOverride,
            InputLanguageOverride = assistant.InputLanguageOverride,
            OutputLanguageOverride = assistant.OutputLanguageOverride
        };

        if (_intensitySliders.TryGetValue(assistantId, out var intensity))
        {
            effective.Intensity = Math.Clamp((int)Math.Round(intensity.Value), Defaults.MinModeIntensity, Defaults.MaxModeIntensity);
        }
        if (_writingStyles.TryGetValue(assistantId, out var style) && style.SelectedItem is string label)
        {
            var match = WritingStyleLabels.FirstOrDefault(pair => pair.Value == label);
            if (!string.IsNullOrEmpty(match.Value))
            {
                effective.WritingStyle = match.Key;
            }
        }
        if (_assistantInputLanguageOverrides.TryGetValue(assistantId, out var inCombo))
        {
            effective.InputLanguageOverride = SelectedAssistantLanguageOverride(inCombo, isInput: true);
        }
        if (_assistantOutputLanguageOverrides.TryGetValue(assistantId, out var outCombo))
        {
            effective.OutputLanguageOverride = SelectedAssistantLanguageOverride(outCombo, isInput: false);
        }
        if (_paragraphDensities.TryGetValue(assistantId, out var paragraphCombo) && paragraphCombo.SelectedItem is string pLabel)
        {
            var match = ParagraphDensityLabels.FirstOrDefault(pair => pair.Value == pLabel);
            if (!string.IsNullOrEmpty(match.Value))
            {
                effective.ParagraphDensity = match.Key;
            }
        }

        if (_emojiExpressions.TryGetValue(assistantId, out var emojiCombo) && emojiCombo.SelectedItem is string eLabel)
        {
            var match = EmojiExpressionLabels.FirstOrDefault(pair => pair.Value == eLabel);
            if (!string.IsNullOrEmpty(match.Value))
            {
                effective.EmojiExpression = match.Key;
            }
        }

        var effectiveInput = PromptComposition.EffectiveInputLanguage(_settings, effective);
        var effectiveOutput = PromptComposition.EffectiveOutputLanguage(_settings, effective);
        return PromptComposition.BuildPolicyBlock(_profile, effective, effectiveInput, effectiveOutput);
    }

    private static string? SelectedAssistantLanguageOverride(ComboBox combo, bool isInput)
    {
        var selected = combo.SelectedItem?.ToString();
        if (string.IsNullOrWhiteSpace(selected) || selected.Equals(GlobalStandardLabel, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var list = isInput ? Defaults.InputLanguages : Defaults.OutputLanguages;
        return list.FirstOrDefault(language => string.Equals(language.Name, selected, StringComparison.OrdinalIgnoreCase))?.Code;
    }

    private bool ProfileSupportsWritingStyle() =>
        !string.IsNullOrWhiteSpace(_profile.WritingStyleInstruction(WritingStyle.Neutral));

    private bool ProfileSupportsEmojiExpression() =>
        !string.IsNullOrWhiteSpace(_profile.EmojiExpressionInstruction(EmojiExpression.Balanced));

    private UIElement BuildInsertPage()
    {
        var panel = PageStack();
        panel.Children.Add(Card(Form("Einfügen", "Empfohlen: Einfügen über die Zwischenablage – der Text landet auf einmal im Ziel, wie bei Strg+V, und bleibt auch bei längeren Antworten schnell. Blockiert ein anderes Programm die Zwischenablage, versucht Magic Voice automatisch einmal „direkt tippen“. „Direkt tippen“ simuliert einzelne Tastendrücke und wirkt bei viel Text wie langsames Tippen; wähle es manuell, wenn die Zielanwendung Einfügen aus der Zwischenablage nicht zuverlässig übernimmt (ohne Fehlermeldung lässt sich das nicht automatisch erkennen).",
            _insertMethod,
            _restoreClipboard)));
        return Page(panel);
    }

    private UIElement BuildGeneralPage()
    {
        var panel = PageStack();
        panel.Children.Add(Card(Form("Startverhalten und Grenzen", null,
            _autostart,
            _launchMinimized,
            _minimizeToTray,
            _playRecordingSounds,
            _maxSeconds,
            _timeoutSeconds)));
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Zurücksetzen"),
                Body("Setzt alle Einstellungen auf die Standardwerte zurück. API-Schlüssel bleiben dabei erhalten."),
                ActionRow(Button("Standard wiederherstellen", ResetDefaults))
            }
        }));
        return Page(panel);
    }

    private UIElement BuildDiagnosticsPage()
    {
        RefreshDiagnostics();
        var panel = PageStack();
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 12,
            Children =
            {
                Header("Diagnose"),
                ActionRow(
                    Button("Aktualisieren", RefreshDiagnostics),
                    Button("Logdatei öffnen", OpenLogFile),
                    Button("Kopieren", CopyDiagnostics)),
                _diagnostics
            }
        }));
        return Page(panel);
    }

    private UIElement BuildAboutPage()
    {
        var panel = PageStack();
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 12,
            Children =
            {
                Header(_profile.AppName),
                Body("Assistent für gesprochene Eingabe bei gedrückter Taste (Push-to-Talk): Transkription, KI-gestützte Textverarbeitung und Einfügen in aktive Anwendungen."),
                KeyValue("Version", GetAppVersion()),
                KeyValue("Autor", _profile.AuthorName),
                KeyValue("Copyright", _profile.CopyrightText),
                KeyValue("Lizenz", _profile.LicenseName)
            }
        }));
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Open-Source"),
                Body("Magic-Voice wird als Open-Source-Software veröffentlicht. Du darfst die App verwenden, kopieren, verändern, weitergeben und eigene Varianten daraus erstellen."),
                Body("Die Veröffentlichung erfolgt unter der MIT-Lizenz. Der Urheberrechtshinweis von Ronny Schulz und der Lizenztext müssen bei Weitergabe erhalten bleiben."),
                Body("Die Software wird ohne Gewährleistung bereitgestellt.")
            }
        }));
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Funktionsumfang"),
                Body("Die App nimmt Sprache nur auf, solange ein Tastenkürzel gedrückt gehalten wird, transkribiert sie über den konfigurierten Cloud-Anbieter und verarbeitet den Text nach deiner Anweisung, dem Schreibstil, der Intensität und den Absatz-Optionen in der Assistenten-Karte."),
                Body("Eingabe- und Ausgabesprache können getrennt eingestellt werden, wodurch die Verarbeitung auch als Übersetzung genutzt werden kann."),
                Body("Der Assistenten-Typ wählt nur die Quelle (Transkript, nur Anweisung oder Zwischenablage); Feinsteuerung erfolgt über die weiteren Felder pro Tastenkürzel.")
            }
        }));
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Datenschutz und Verarbeitung"),
                Body("Audiodaten, Transkripte und finale Texte werden standardmäßig nicht in Logs gespeichert. Logs enthalten technische Statusinformationen und Fehlerhinweise."),
                Body("API-Schlüssel werden lokal für den aktuellen Windows-Benutzer per DPAPI verschlüsselt gespeichert."),
                Body("Für Transkription und KI-Verarbeitung werden Audio bzw. Text an den konfigurierten Anbieter übertragen. Welche Daten dort verarbeitet werden, richtet sich nach dessen Bedingungen.")
            }
        }));
        panel.Children.Add(Card(new StackPanel
        {
            Spacing = 10,
            Children =
            {
                Header("Technik"),
                KeyValue("UI", "WinUI 3 / Windows App SDK"),
                KeyValue("Runtime", $".NET {Environment.Version}"),
                KeyValue("Architektur", "Dienstorientiert mit Abhängigkeitsinjektion"),
                ActionRow(
                    Button("Lizenz öffnen", OpenLicense),
                    Button("Drittanbieterhinweise öffnen", OpenThirdPartyNotices),
                    Button("Diagnose öffnen", () => Navigate(DiagnosticsPage)))
            }
        }));
        return Page(panel);
    }

    private async Task LoadSettingsAsync()
    {
        _settings = await _settingsService.LoadAsync();
        SelectCombo(_sttProvider, _settings.SttProvider);
        SelectEditableComboValue(_sttModel, _settings.SttModel, Defaults.OpenAiSttModels[0]);
        SelectLanguage(_inputLanguage, _settings.InputLanguage);
        SelectLanguage(_outputLanguage, _settings.OutputLanguage);
        SelectAudioDevice(_settings.AudioInputDeviceId);
        SelectCombo(_llmProvider, _settings.LlmProvider);
        SelectEditableComboValue(_llmModel, _settings.LlmModel, Defaults.DefaultLlmModel);
        SelectCombo(_insertMethod, InsertMethodLabel(_settings.InsertMethod));
        _restoreClipboard.IsChecked = _settings.RestoreClipboard;
        _launchMinimized.IsChecked = _settings.LaunchMinimizedToTray;
        _minimizeToTray.IsChecked = _settings.MinimizeToTray;
        _playRecordingSounds.IsChecked = _settings.PlayRecordingSounds;
        _autostart.IsChecked = _autostartService.IsEnabled();
        _maxSeconds.Text = _settings.RecordingMaxSeconds.ToString();
        _timeoutSeconds.Text = _settings.ProcessingTimeoutSeconds.ToString();
        ApplyApiKeysFromSettings();
        RefreshDiagnostics();
        _settings.LastSelectedSettingsSection = NormalizeSectionTag(_settings.LastSelectedSettingsSection);
        SelectNavigationItemByTag(_settings.LastSelectedSettingsSection);
        Navigate(_settings.LastSelectedSettingsSection);
        _initialNavigationDone = true;
        WireAutoPersist();
    }

    /// <summary>
    /// Hängt LostFocus-/SelectionChanged-/Toggle-Handler an die relevanten
    /// Eingabefelder, damit Änderungen sofort gespeichert werden und der
    /// Tray-Status (Warnung "Einrichtung erforderlich") ohne Neustart oder
    /// Navigationswechsel wegfällt, sobald alle Pflichtfelder gefüllt sind.
    /// </summary>
    private void WireAutoPersist()
    {
        if (_autoPersistWired)
        {
            return;
        }
        _autoPersistWired = true;

        void OnTextLost(object _, RoutedEventArgs __) => _ = AutoPersistAsync();
        void OnSelectionChanged(object _, SelectionChangedEventArgs __) => _ = AutoPersistAsync();
        void OnToggle(object _, RoutedEventArgs __) => _ = AutoPersistAsync();

        foreach (var tb in new[] { _llmApiKey, _sttApiKey, _maxSeconds, _timeoutSeconds })
        {
            tb.LostFocus += OnTextLost;
        }

        foreach (var cb in new[] { _sttProvider, _sttModel, _llmProvider, _llmModel,
                                   _audioInputDevice, _inputLanguage, _outputLanguage, _insertMethod })
        {
            cb.SelectionChanged += OnSelectionChanged;
        }

        foreach (var chk in new[] { _restoreClipboard, _launchMinimized, _minimizeToTray, _playRecordingSounds, _autostart })
        {
            chk.Checked += OnToggle;
            chk.Unchecked += OnToggle;
        }
    }

    private async Task AutoPersistAsync()
    {
        if (!_initialNavigationDone || _isExiting)
        {
            return;
        }
        try
        {
            await PersistAsync();
        }
        catch (Exception ex)
        {
            _statusService.SetStatus(TrayStatus.Error, $"Einstellungen konnten nicht gespeichert werden: {ex.Message}");
        }
    }

    private void ApplyApiKeysFromSettings()
    {
        try
        {
            _llmApiKey.Text = _settings.HasEncryptedLlmApiKey
                ? _secretProtector.Unprotect(_settings.LlmApiKeyEncrypted!)
                : string.Empty;
        }
        catch
        {
            _llmApiKey.Text = string.Empty;
        }

        try
        {
            _sttApiKey.Text = _settings.HasEncryptedSttApiKey
                ? _secretProtector.Unprotect(_settings.SttApiKeyEncrypted!)
                : string.Empty;
        }
        catch
        {
            _sttApiKey.Text = string.Empty;
        }
    }

    private async Task PersistAsync(string? successMessage = null)
    {
        SaveSettingsFromFields();
        _settings.LlmApiKeyEncrypted = string.IsNullOrWhiteSpace(_llmApiKey.Text)
            ? null
            : _secretProtector.Protect(_llmApiKey.Text.Trim());
        _settings.SttApiKeyEncrypted = string.IsNullOrWhiteSpace(_sttApiKey.Text)
            ? null
            : _secretProtector.Protect(_sttApiKey.Text.Trim());

        await _settingsService.SaveAsync(_settings);
        _autostartService.SetEnabled(_autostart.IsChecked == true);
        var issues = await _hotkeyService.RegisterAsync(_settings);
        var readiness = _settingsService.Validate(_settings);
        if (issues.Count > 0)
        {
            _statusService.SetStatus(TrayStatus.ConfigurationRequired, "Dieses Tastenkürzel konnte nicht registriert werden. Bitte wähle eine andere Kombination.");
        }
        else if (!readiness.IsReady)
        {
            _statusService.SetStatus(TrayStatus.ConfigurationRequired, "Einrichtung erforderlich. Bitte prüfe die markierten Einstellungen.");
        }
        else if (_statusService.CurrentStatus is not TrayStatus.Recording and not TrayStatus.Processing)
        {
            _statusService.SetStatus(
                TrayStatus.Idle,
                !string.IsNullOrEmpty(successMessage)
                    ? successMessage
                    : "Bereit. Halte ein Tastenkürzel gedrückt, um zu diktieren.");
        }

        await _fileLogger.WriteAsync($"Einstellungen gespeichert. Status={_statusService.CurrentStatus}");
        if (!_isExiting)
        {
            ApplyApiKeysFromSettings();
            RefreshDiagnostics();
        }
    }

    private void SaveSettingsFromFields()
    {
        _settings.SttProvider = Selected(_sttProvider, Defaults.OpenAiProviderName);
        _settings.SttModel = ReadEditableComboValue(_sttModel, Defaults.OpenAiSttModels[0]);
        _settings.InputLanguage = SelectedLanguage(_inputLanguage, "de");
        _settings.OutputLanguage = SelectedLanguage(_outputLanguage, Defaults.SameAsInputLanguageCode);
        _settings.AudioInputDeviceId = SelectedAudioDeviceId();
        _settings.LlmProvider = Selected(_llmProvider, Defaults.OpenAiProviderName);
        _settings.LlmModel = ReadEditableComboValue(_llmModel, Defaults.DefaultLlmModel);
        _settings.InsertMethod = SelectedInsertMethod();
        _settings.RestoreClipboard = _restoreClipboard.IsChecked == true;
        _settings.LaunchMinimizedToTray = _launchMinimized.IsChecked == true;
        _settings.MinimizeToTray = _minimizeToTray.IsChecked == true;
        _settings.PlayRecordingSounds = _playRecordingSounds.IsChecked == true;
        _settings.RecordingMaxSeconds = ParseBoundedInt(_maxSeconds.Text, 60, 1, 300);
        _settings.ProcessingTimeoutSeconds = ParseBoundedInt(_timeoutSeconds.Text, 45, 5, 180);

        foreach (var assistant in _settings.Assistants)
        {
            if (_names.TryGetValue(assistant.Id, out var nameBox) && !string.IsNullOrWhiteSpace(nameBox.Text))
            {
                assistant.Name = nameBox.Text.Trim();
            }

            if (_assistantTypes.TryGetValue(assistant.Id, out var typePicker) && typePicker.SelectedItem is string typeName)
            {
                var modeDef = _profile.Modes.FirstOrDefault(m =>
                    string.Equals(m.Name, typeName, StringComparison.OrdinalIgnoreCase));
                if (modeDef is not null)
                {
                    assistant.Type = modeDef.Mode;
                }
            }

            if (_prompts.TryGetValue(assistant.Id, out var prompt))
            {
                assistant.Prompt = prompt.Text;
            }

            if (_systemPromptOverrideCheckBoxes.TryGetValue(assistant.Id, out var systemPromptOverrideCheck)
                && _systemPrompts.TryGetValue(assistant.Id, out var systemPrompt))
            {
                assistant.SystemPromptOverride = systemPromptOverrideCheck.IsChecked == true
                    ? (string.IsNullOrWhiteSpace(systemPrompt.Text) ? null : systemPrompt.Text)
                    : null;
            }

            if (_assistantInputLanguageOverrides.TryGetValue(assistant.Id, out var inputOverride))
            {
                assistant.InputLanguageOverride = SelectedAssistantLanguageOverride(inputOverride, isInput: true);
            }
            if (_assistantOutputLanguageOverrides.TryGetValue(assistant.Id, out var outputOverride))
            {
                assistant.OutputLanguageOverride = SelectedAssistantLanguageOverride(outputOverride, isInput: false);
            }

            if (_intensitySliders.TryGetValue(assistant.Id, out var intensity))
            {
                assistant.Intensity = Math.Clamp((int)Math.Round(intensity.Value), Defaults.MinModeIntensity, Defaults.MaxModeIntensity);
            }

            if (_writingStyles.TryGetValue(assistant.Id, out var style) && style.SelectedItem is string label)
            {
                var match = WritingStyleLabels.FirstOrDefault(pair => pair.Value == label);
                if (!string.IsNullOrEmpty(match.Value))
                {
                    assistant.WritingStyle = match.Key;
                }
            }

            if (_paragraphDensities.TryGetValue(assistant.Id, out var paragraphCombo) && paragraphCombo.SelectedItem is string pLabel)
            {
                var pMatch = ParagraphDensityLabels.FirstOrDefault(pair => pair.Value == pLabel);
                if (!string.IsNullOrEmpty(pMatch.Value))
                {
                    assistant.ParagraphDensity = pMatch.Key;
                }
            }

            if (_emojiExpressions.TryGetValue(assistant.Id, out var emojiCombo) && emojiCombo.SelectedItem is string eLabel)
            {
                var eMatch = EmojiExpressionLabels.FirstOrDefault(pair => pair.Value == eLabel);
                if (!string.IsNullOrEmpty(eMatch.Value))
                {
                    assistant.EmojiExpression = eMatch.Key;
                }
            }
        }
    }

    private string FormatIntensityValue(AssistantMode type, int intensity) =>
        $"{intensity}/{Defaults.MaxModeIntensity} – {_profile.IntensityStepName(type, intensity)}";

    private void ResetDefaults()
    {
        SelectCombo(_sttProvider, Defaults.OpenAiProviderName);
        SelectEditableComboValue(_sttModel, Defaults.OpenAiSttModels[0], Defaults.OpenAiSttModels[0]);
        SelectLanguage(_inputLanguage, "de");
        SelectLanguage(_outputLanguage, Defaults.SameAsInputLanguageCode);
        SelectAudioDevice(Defaults.DefaultAudioInputDeviceId);
        SelectCombo(_llmProvider, Defaults.OpenAiProviderName);
        SelectEditableComboValue(_llmModel, Defaults.DefaultLlmModel, Defaults.DefaultLlmModel);
        SelectCombo(_insertMethod, InsertMethodLabels[InsertMethod.Clipboard]);
        _launchMinimized.IsChecked = true;
        _minimizeToTray.IsChecked = true;
        _playRecordingSounds.IsChecked = true;

        _settings.Assistants = _profile.CreateDefaultAssistants();
        RebuildHotkeyPage();

        _ = PersistAsync("Standardwerte wurden eingesetzt und gespeichert.");
    }

    private void StartHotkeyCapture(string assistantId)
    {
        _capturingAssistantId = assistantId;
        ResetCaptureButtons();
        if (_captureButtons.TryGetValue(assistantId, out var button))
        {
            button.Content = "Tastenkombination drücken …";
        }
        _statusService.SetStatus(TrayStatus.ConfigurationRequired, "Drücke jetzt die gewünschte Tastenkombination. Escape bricht ab.");
    }

    private void OnRootKeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (_capturingAssistantId is not { } assistantId)
        {
            return;
        }

        if (e.Key == VirtualKey.Escape)
        {
            _capturingAssistantId = null;
            ResetCaptureButtons();
            _statusService.SetStatus(TrayStatus.ConfigurationRequired, "Aufnahme des Tastenkürzels abgebrochen.");
            e.Handled = true;
            return;
        }

        if (IsModifierKey(e.Key))
        {
            return;
        }

        var gesture = BuildGesture(e.Key);
        if (!HotkeyParser.TryParse(gesture, out _, out var error))
        {
            _statusService.SetStatus(TrayStatus.ConfigurationRequired, error);
            e.Handled = true;
            return;
        }

        var assistant = _settings.Assistants.FirstOrDefault(a => a.Id == assistantId);
        if (assistant is null)
        {
            _capturingAssistantId = null;
            ResetCaptureButtons();
            return;
        }

        assistant.Hotkey = gesture;
        if (_hotkeyValues.TryGetValue(assistantId, out var hotkeyValue))
        {
            hotkeyValue.Text = gesture;
        }
        _capturingAssistantId = null;
        ResetCaptureButtons();
        e.Handled = true;
        _ = PersistAsync("Tastenkürzel übernommen und aktiviert.");
    }

    private string BuildGesture(VirtualKey key)
    {
        var parts = new List<string>();
        if (IsPressed(0x11)) parts.Add("Ctrl");
        if (IsPressed(0x12)) parts.Add("Alt");
        if (IsPressed(0x10)) parts.Add("Shift");
        if (IsPressed(0x5B) || IsPressed(0x5C)) parts.Add("Win");
        parts.Add(KeyName(key));
        return string.Join("+", parts);
    }

    private void ResetCaptureButtons()
    {
        foreach (var button in _captureButtons.Values)
        {
            button.Content = "Tastenkürzel aufnehmen";
        }
    }

    private async Task TestLlmConnectionAsync()
    {
        await PersistAsync();
        var readiness = _settingsService.Validate(_settings);
        var issues = readiness.Issues
            .Where(issue => issue.Field.StartsWith("llm", StringComparison.OrdinalIgnoreCase) || issue.Field.StartsWith("stt", StringComparison.OrdinalIgnoreCase))
            .Where(issue => !(issue.Field == "llmApiKey" && !string.IsNullOrWhiteSpace(_llmApiKey.Text)))
            .Where(issue => !(issue.Field == "sttApiKey" && !string.IsNullOrWhiteSpace(_sttApiKey.Text)))
            .ToList();
        if (issues.Count > 0)
        {
            _statusService.SetStatus(TrayStatus.ConfigurationRequired, string.Join(" ", issues.Select(issue => issue.Message)));
            RefreshDiagnostics();
            return;
        }

        var keyPlain = !string.IsNullOrWhiteSpace(_llmApiKey.Text)
            ? _llmApiKey.Text.Trim()
            : _settings.HasEncryptedLlmApiKey
                ? _secretProtector.Unprotect(_settings.LlmApiKeyEncrypted!)
                : null;
        if (string.IsNullOrEmpty(keyPlain))
        {
            _statusService.SetStatus(TrayStatus.ConfigurationRequired, "Bitte einen API-Schlüssel für die KI-Verarbeitung eintragen oder speichern.");
            RefreshDiagnostics();
            return;
        }

        try
        {
            Defaults.TryGetLlmEndpoint(_settings.LlmProvider, out var endpoint);
            await _llmService.ProcessAsync(
                new LlmRequest("Verbindungstest", _profile.Modes[0].Mode, _settings.LlmProvider, endpoint, _settings.LlmModel, keyPlain, _profile.SystemPrompt, "Antworte exakt mit OK."),
                CancellationToken.None);
            _statusService.SetStatus(TrayStatus.Idle, "KI-Verbindung erfolgreich getestet.");
        }
        catch
        {
            _statusService.SetStatus(TrayStatus.Error, "KI-Verbindung fehlgeschlagen. Bitte prüfe API-Schlüssel und Modell.");
        }

        RefreshDiagnostics();
    }

    private void OpenLogFile()
    {
        try
        {
            Directory.CreateDirectory(_settingsService.LogDirectory);
            if (!File.Exists(_fileLogger.CurrentLogPath))
            {
                File.WriteAllText(_fileLogger.CurrentLogPath, string.Empty);
            }

            Process.Start(new ProcessStartInfo(_fileLogger.CurrentLogPath) { UseShellExecute = true });
        }
        catch
        {
            _statusService.SetStatus(TrayStatus.Error, "Logdatei konnte nicht geöffnet werden.");
        }
    }

    private void OpenThirdPartyNotices()
    {
        try
        {
            var path = Path.Combine(AppContext.BaseDirectory, "THIRD_PARTY_NOTICES.md");
            if (!File.Exists(path))
            {
                _statusService.SetStatus(TrayStatus.Error, "Drittanbieterhinweise wurden im Installationsverzeichnis nicht gefunden.");
                return;
            }

            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch
        {
            _statusService.SetStatus(TrayStatus.Error, "Drittanbieterhinweise konnten nicht geöffnet werden.");
        }
    }

    private void OpenLicense()
    {
        try
        {
            var path = Path.Combine(AppContext.BaseDirectory, "LICENSE");
            if (!File.Exists(path))
            {
                _statusService.SetStatus(TrayStatus.Error, "Lizenzdatei wurde im Installationsverzeichnis nicht gefunden.");
                return;
            }

            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch
        {
            _statusService.SetStatus(TrayStatus.Error, "Lizenzdatei konnte nicht geöffnet werden.");
        }
    }

    private void CopyDiagnostics()
    {
        try
        {
            var package = new DataPackage();
            package.SetText(_diagnostics.Text);
            Clipboard.SetContent(package);
            Clipboard.Flush();
            _statusService.SetStatus(TrayStatus.Idle, "Diagnose wurde in die Zwischenablage kopiert.");
        }
        catch
        {
            _statusService.SetStatus(TrayStatus.Error, "Diagnose konnte nicht kopiert werden.");
        }
    }

    private void UpdateInfoBar(TrayStatus status, string message)
    {
        _statusInfo.Title = status switch
        {
            TrayStatus.Idle => "Bereit",
            TrayStatus.Paused => "Inaktiv",
            TrayStatus.Recording => "Aufnahme läuft",
            TrayStatus.Processing => "Verarbeitung",
            TrayStatus.Success => "Erfolgreich",
            TrayStatus.Error => "Fehler",
            _ => "Einrichtung erforderlich"
        };
        _statusInfo.Message = message;
        _statusInfo.Severity = status switch
        {
            TrayStatus.Idle or TrayStatus.Success => InfoBarSeverity.Success,
            TrayStatus.Error => InfoBarSeverity.Error,
            TrayStatus.Processing or TrayStatus.Recording => InfoBarSeverity.Informational,
            _ => InfoBarSeverity.Warning
        };
        // Kein kleines ProgressRing-Symbol im Kachelfeld: Segoe MDL2 wie bei den Navigations-Icons (ungeführt, größer).
        _statusInfo.IconSource = status switch
        {
            TrayStatus.Recording => new FontIconSource { Glyph = "\uE8B7", FontSize = 20 },
            TrayStatus.Processing => new FontIconSource { Glyph = "\uE8EF", FontSize = 20 },
            _ => null
        };
        RefreshDiagnostics();
    }

    private static NavigationViewItem NavItem(string text, string tag, string glyph) => new()
    {
        Content = text,
        Tag = tag,
        Icon = new FontIcon { Glyph = glyph }
    };

    private UIElement Page(UIElement content)
    {
        Detach(content);
        return new ScrollViewer
        {
            Content = content,
            Background = new SolidColorBrush(Colors.Transparent),
            Padding = new Thickness(20, 16, 20, 28),
            HorizontalContentAlignment = HorizontalAlignment.Stretch,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        };
    }

    /// <summary>
    /// Entfernt den Standard-Inhaltsrahmen der NavigationView (eigener Layer, Kartenstrich, abgeschnittene Ecken).
    /// </summary>
    private static void StripNavigationViewContentChrome(NavigationView navigationView)
    {
        navigationView.Resources["NavigationViewContentGridBorderBrush"] = new SolidColorBrush(Colors.Transparent);
        navigationView.Resources["NavigationViewContentGridBorderThickness"] = new Thickness(0);
        navigationView.Resources["NavigationViewContentBackground"] =
            ResourceBrush("ApplicationPageBackgroundThemeBrush", new SolidColorBrush(Colors.Transparent));
        navigationView.Resources["NavigationViewContentGridCornerRadius"] = new CornerRadius(0);
    }

    private static StackPanel PageStack() => new() { Spacing = 22, HorizontalAlignment = HorizontalAlignment.Stretch };

    private static Border Card(UIElement child) => new()
    {
        Child = Detach(child),
        Padding = new Thickness(22),
        CornerRadius = new CornerRadius(12),
        Background = ResourceBrush("CardBackgroundFillColorDefaultBrush", new SolidColorBrush(Windows.UI.Color.FromArgb(18, 128, 128, 128))),
        BorderBrush = ResourceBrush("CardStrokeColorDefaultBrush", new SolidColorBrush(Windows.UI.Color.FromArgb(72, 128, 128, 128))),
        BorderThickness = new Thickness(1),
        HorizontalAlignment = HorizontalAlignment.Stretch
    };

    private static StackPanel Form(string title, string? description, params UIElement[] children)
    {
        var panel = new StackPanel { Spacing = 12, HorizontalAlignment = HorizontalAlignment.Stretch };
        panel.Children.Add(Header(title));
        if (!string.IsNullOrWhiteSpace(description))
        {
            panel.Children.Add(Body(description));
        }

        foreach (var child in children)
        {
            panel.Children.Add(Detach(child));
        }

        return panel;
    }

    private static StackPanel ActionRow(params UIElement[] children)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        foreach (var child in children)
        {
            panel.Children.Add(Detach(child));
        }

        return panel;
    }

    private static UIElement Detach(UIElement element)
    {
        if (element is not FrameworkElement frameworkElement || frameworkElement.Parent is not { } parent)
        {
            return element;
        }

        switch (parent)
        {
            case Panel panel:
                panel.Children.Remove(element);
                break;
            case Border border when ReferenceEquals(border.Child, element):
                border.Child = null;
                break;
            case ContentControl contentControl when ReferenceEquals(contentControl.Content, element):
                contentControl.Content = null;
                break;
            case ScrollViewer scrollViewer when ReferenceEquals(scrollViewer.Content, element):
                scrollViewer.Content = null;
                break;
        }

        return element;
    }

    private static TextBlock Header(string text) => new() { Text = text, FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap };

    private static TextBlock Body(string text) => new() { Text = text, TextWrapping = TextWrapping.Wrap, Opacity = 0.84 };

    private static Grid KeyValue(string key, string value)
    {
        var grid = new Grid
        {
            ColumnDefinitions =
            {
                new ColumnDefinition { Width = new GridLength(170) },
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }
            },
            ColumnSpacing = 12
        };
        grid.Children.Add(new TextBlock
        {
            Text = key,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap
        });
        grid.Children.Add(WithColumn(new TextBlock
        {
            Text = value,
            TextWrapping = TextWrapping.Wrap,
            Opacity = 0.86
        }, 1));
        return grid;
    }

    private static Button Button(string text, Action action)
    {
        var button = new Button { Content = text };
        button.Click += (_, _) => action();
        return button;
    }

    private static Button Button(string text, Func<Task> action)
    {
        var button = new Button { Content = text };
        button.Click += async (_, _) => await action();
        return button;
    }

    private static Button CompactArrowReorderButton(bool up, string toolTip, Action action)
    {
        var button = new Button
        {
            Content = new TextBlock
            {
                Text = up ? "▲" : "▼",
                FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            },
            Padding = new Thickness(8, 4, 8, 4),
            MinWidth = 36,
            MinHeight = 32
        };
        ToolTipService.SetToolTip(button, toolTip);
        button.Click += (_, _) => action();
        return button;
    }

    /// <summary>Kompakter Löschen-Button (MDL2 „Delete“), gleiche Maße wie <see cref="CompactArrowReorderButton"/>.</summary>
    private static Button CompactDeleteAssistantButton(Func<Task> action)
    {
        var button = new Button
        {
            Content = new FontIcon
            {
                Glyph = "\uE74D",
                FontSize = 14,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            },
            Padding = new Thickness(8, 4, 8, 4),
            MinWidth = 36,
            MinHeight = 32
        };
        ToolTipService.SetToolTip(button, "Assistent löschen");
        button.Click += async (_, _) => await action();
        return button;
    }

    private async Task<bool> ConfirmDeleteAssistantAsync(string displayName)
    {
        if (Root.XamlRoot is null)
        {
            return false;
        }

        var dialog = new ContentDialog
        {
            Title = "Assistent löschen?",
            Content = $"„{displayName}“ unwiderruflich aus der Liste entfernen?",
            PrimaryButtonText = "Löschen",
            CloseButtonText = "Abbrechen",
            DefaultButton = ContentDialogButton.Close,
            XamlRoot = Root.XamlRoot
        };

        var result = await dialog.ShowAsync();
        return result == ContentDialogResult.Primary;
    }

    private static UIElement WithColumn(FrameworkElement element, int column)
    {
        Grid.SetColumn(element, column);
        return element;
    }

    private static ComboBox Combo(string header) => new() { Header = header, HorizontalAlignment = HorizontalAlignment.Stretch, MinWidth = 260 };

    private static TextBox TextField(string header, string placeholder) => new() { Header = header, PlaceholderText = placeholder, HorizontalAlignment = HorizontalAlignment.Stretch };

    private void RefreshAudioDevices()
    {
        RefreshAudioDeviceList();
        SelectAudioDevice(_settings.AudioInputDeviceId);
        _statusService.SetStatus(TrayStatus.ConfigurationRequired, "Audiogeräte wurden aktualisiert. Speichere die Einstellungen, um die Auswahl zu übernehmen.");
    }

    private void RefreshAudioDeviceList()
    {
        var selectedId = SelectedAudioDeviceId();
        _audioInputDevice.Items.Clear();
        _audioDeviceIdsByDisplayName.Clear();
        foreach (var device in _audioDeviceService.GetInputDevices())
        {
            var displayName = DeviceDisplayName(device);
            _audioDeviceIdsByDisplayName[displayName] = device.Id;
            _audioInputDevice.Items.Add(displayName);
        }

        SelectAudioDevice(string.IsNullOrWhiteSpace(selectedId) ? Defaults.DefaultAudioInputDeviceId : selectedId);
    }

    private void SelectAudioDevice(string deviceId)
    {
        var match = _audioDeviceIdsByDisplayName
            .FirstOrDefault(pair => pair.Value.Equals(deviceId, StringComparison.OrdinalIgnoreCase));
        _audioInputDevice.SelectedItem = string.IsNullOrWhiteSpace(match.Key) ? _audioInputDevice.Items.FirstOrDefault() : match.Key;
    }

    private string SelectedAudioDeviceId() =>
        _audioInputDevice.SelectedItem is string displayName && _audioDeviceIdsByDisplayName.TryGetValue(displayName, out var deviceId)
            ? deviceId
            : Defaults.DefaultAudioInputDeviceId;

    private string AudioDeviceName(string deviceId) =>
        _audioDeviceService.GetInputDevices()
            .FirstOrDefault(device => device.Id.Equals(deviceId, StringComparison.OrdinalIgnoreCase))?.Name
        ?? "Systemstandard";

    private static string DeviceDisplayName(AudioInputDevice device) =>
        device.IsDefault ? "Systemstandard" : $"{device.Name} ({device.Id})";

    private static void SelectLanguage(ComboBox combo, string languageCode)
    {
        combo.SelectedItem = Defaults.LanguageName(languageCode);
        if (combo.SelectedItem is null || !combo.Items.Contains(combo.SelectedItem))
        {
            combo.SelectedItem = combo.Items.FirstOrDefault();
        }
    }

    private static string SelectedLanguage(ComboBox combo, string fallback) =>
        Defaults.InputLanguages.Concat(Defaults.OutputLanguages)
            .FirstOrDefault(language => string.Equals(language.Name, combo.SelectedItem?.ToString(), StringComparison.OrdinalIgnoreCase))?.Code
        ?? fallback;

    private static void AddItems<T>(ComboBox combo, IEnumerable<T> items)
    {
        foreach (var item in items)
        {
            combo.Items.Add(item);
        }
    }

    private static void SelectCombo(ComboBox combo, string value)
    {
        combo.SelectedItem = combo.Items.Cast<object>().FirstOrDefault(item => string.Equals(item.ToString(), value, StringComparison.OrdinalIgnoreCase)) ?? combo.Items.FirstOrDefault();
    }

    private void SelectEditableComboValue(ComboBox combo, string value, string fallback)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            value = fallback;
        }

        var match = combo.Items.Cast<object>().FirstOrDefault(item =>
            string.Equals(item.ToString(), value, StringComparison.OrdinalIgnoreCase));
        combo.SelectedItem = match;
        combo.Text = value;
        _ = DispatcherQueue.TryEnqueue(() =>
        {
            combo.Text = value;
            if (match is not null)
            {
                combo.SelectedItem = match;
            }
        });
    }

    private static string ReadEditableComboValue(ComboBox combo, string fallback)
    {
        if (combo.IsEditable)
        {
            var text = combo.Text?.Trim();
            if (!string.IsNullOrEmpty(text))
            {
                return text;
            }
        }

        return combo.SelectedItem?.ToString()?.Trim() ?? fallback;
    }

    private static string Selected(ComboBox combo, string fallback) => combo.SelectedItem?.ToString() ?? fallback;

    private InsertMethod SelectedInsertMethod() =>
        _insertMethod.SelectedItem?.ToString() == InsertMethodLabels[InsertMethod.Clipboard]
            ? InsertMethod.Clipboard
            : InsertMethod.SendInput;

    private static string InsertMethodLabel(InsertMethod method) =>
        InsertMethodLabels.TryGetValue(method, out var label) ? label : InsertMethodLabels[InsertMethod.Clipboard];

    private static Style? ResourceStyle(string key) =>
        Microsoft.UI.Xaml.Application.Current.Resources.TryGetValue(key, out var value) && value is Style style
            ? style
            : null;

    private static Brush ResourceBrush(string key, Brush? fallback = null) =>
        Microsoft.UI.Xaml.Application.Current.Resources.TryGetValue(key, out var value) && value is Brush brush
            ? brush
            : fallback ?? new SolidColorBrush(Colors.Transparent);

    private static string TitleFor(string tag) => tag switch
    {
        OverviewPage => "Übersicht",
        ProviderPage => "Anbieter",
        HotkeyPage => "Assistenten",
        InsertPage => "Einfügen",
        GeneralPage => "Allgemein",
        DiagnosticsPage => "Diagnose",
        AudioLanguagePage => "Audio & Sprache",
        AboutPage => "Über",
        _ => "Übersicht"
    };

    private async void OnNavigationSelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (args.SelectedItem is NavigationViewItem item && item.Tag is string tag)
        {
            var shouldPersist = _initialNavigationDone
                && !string.Equals(_settings.LastSelectedSettingsSection, tag, StringComparison.Ordinal);
            Navigate(tag);
            if (shouldPersist)
            {
                try
                {
                    await PersistAsync();
                }
                catch (Exception ex)
                {
                    _statusService.SetStatus(TrayStatus.Error, $"Einstellungen konnten nicht gespeichert werden: {ex.Message}");
                }
            }
        }
    }

    private void SelectNavigationItemByTag(string tag)
    {
        foreach (var item in _navigation.MenuItems)
        {
            if (item is NavigationViewItem nav && nav.Tag is string t && t == tag)
            {
                _navigation.SelectedItem = item;
                return;
            }
        }

        if (_navigation.MenuItems.Count > 0)
        {
            _navigation.SelectedItem = _navigation.MenuItems[0];
        }
    }

    private static string NormalizeSectionTag(string? tag)
    {
        var t = tag switch
        {
            null or "" => OverviewPage,
            "Übersicht" => OverviewPage,
            _ => tag
        };

        return t is OverviewPage or AudioLanguagePage or ProviderPage or HotkeyPage or InsertPage or GeneralPage or DiagnosticsPage or AboutPage
            ? t
            : OverviewPage;
    }

    private static string IntensityLabel(AssistantMode mode) => mode switch
    {
        AssistantMode.Transform => "Umarbeitungstiefe (Transkript)",
        AssistantMode.Generate => "Freiheit der Generierung",
        AssistantMode.AnswerClipboard => "Ausführlichkeit der Antwort",
        _ => "Intensität"
    };

    private static int ParseBoundedInt(string text, int fallback, int min, int max) =>
        int.TryParse(text, out var value) ? Math.Clamp(value, min, max) : fallback;

    private static string SanitizeStatus(string message) =>
        message.Contains(Environment.NewLine, StringComparison.Ordinal) ? message.Split(Environment.NewLine)[0] : message;

    private static string FriendlyStatus(TrayStatus status) => status switch
    {
        TrayStatus.Idle => "Bereit",
        TrayStatus.Paused => "Inaktiv",
        TrayStatus.Recording => "Aufnahme",
        TrayStatus.Processing => "Verarbeitung",
        TrayStatus.Success => "Erfolgreich",
        TrayStatus.Error => "Fehler",
        TrayStatus.ConfigurationRequired => "Einrichtung erforderlich",
        _ => status.ToString()
    };

    private static string GetAppVersion() =>
        GetExeProductVersion()
        ?? GetAssemblyInformationalVersion()
        ?? typeof(MainWindow).Assembly.GetName().Version?.ToString()
        ?? "unbekannt";

    private static string? GetAssemblyInformationalVersion() =>
        typeof(MainWindow).Assembly
            .GetCustomAttributes<AssemblyInformationalVersionAttribute>()
            .FirstOrDefault()
            ?.InformationalVersion;

    private static string? GetExeProductVersion()
    {
        var path = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            var v = FileVersionInfo.GetVersionInfo(path).ProductVersion;
            return string.IsNullOrWhiteSpace(v) ? null : v;
        }
        catch
        {
            return null;
        }
    }

    private static bool IsModifierKey(VirtualKey key) =>
        key is VirtualKey.Control or VirtualKey.LeftControl or VirtualKey.RightControl
            or VirtualKey.Menu or VirtualKey.LeftMenu or VirtualKey.RightMenu
            or VirtualKey.Shift or VirtualKey.LeftShift or VirtualKey.RightShift
            or VirtualKey.LeftWindows or VirtualKey.RightWindows;

    private static string KeyName(VirtualKey key) => key switch
    {
        >= VirtualKey.Number0 and <= VirtualKey.Number9 => ((int)key - (int)VirtualKey.Number0).ToString(),
        >= VirtualKey.A and <= VirtualKey.Z => key.ToString(),
        >= VirtualKey.F1 and <= VirtualKey.F24 => key.ToString(),
        VirtualKey.Space => "Space",
        VirtualKey.Enter => "Enter",
        VirtualKey.Tab => "Tab",
        VirtualKey.Escape => "Esc",
        _ => key.ToString()
    };

    private static bool IsPressed(int key) => (GetKeyState(key) & 0x8000) != 0;

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int nVirtKey);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private delegate IntPtr SubclassProc(IntPtr hWnd, uint uMsg, UIntPtr wParam, IntPtr lParam, UIntPtr uIdSubclass, UIntPtr dwRefData);

    [DllImport("Comctl32.dll", SetLastError = true)]
    private static extern bool SetWindowSubclass(IntPtr hWnd, SubclassProc pfnSubclass, UIntPtr uIdSubclass, UIntPtr dwRefData);

    [DllImport("Comctl32.dll")]
    private static extern IntPtr DefSubclassProc(IntPtr hWnd, uint uMsg, UIntPtr wParam, IntPtr lParam);
}
