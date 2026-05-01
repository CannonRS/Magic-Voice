using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;
using MagicVoice.Core;
using MagicVoice.Infrastructure;

namespace MagicVoice;

public sealed partial class App : Microsoft.UI.Xaml.Application
{
    private ServiceProvider? _services;
    private MainWindow? _window;
    private TrayIconController? _tray;
    private IHotkeyService? _hotkeys;
    private SpeechPipeline? _pipeline;
    private IAudioRecorder? _recorder;
    private ITrayStatusService? _status;
    private IFeedbackSoundService? _sounds;
    private IAppProfile? _profile;
    private IClipboardSourceCapture? _clipboardCapture;
    private string? _capturedSourceText;
    private Mutex? _singleInstanceMutex;

    public App()
    {
        InitializeComponent();
        UnhandledException += (_, args) =>
        {
            MagicVoiceStartupCrashLogger.Write(args.Exception);
            args.Handled = true;
        };
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            if (args.ExceptionObject is Exception ex)
            {
                MagicVoiceStartupCrashLogger.Write(ex);
            }
        };
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            MagicVoiceStartupCrashLogger.Write(args.Exception);
            args.SetObserved();
        };
    }

    protected override async void OnLaunched(LaunchActivatedEventArgs args)
    {
        try
        {
            MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: start");
            _services = ConfigureServices();
            MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: services configured");
            _profile = _services.GetRequiredService<IAppProfile>();
            _singleInstanceMutex = new Mutex(initiallyOwned: true, _profile.MutexName, out var firstInstance);
            if (!firstInstance)
            {
                MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: another instance is already running, exiting");
                Exit();
                return;
            }
            _window = ActivatorUtilities.CreateInstance<MainWindow>(_services);
            MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: window created");
            _hotkeys = _services.GetRequiredService<IHotkeyService>();
            _pipeline = _services.GetRequiredService<SpeechPipeline>();
            _recorder = _services.GetRequiredService<IAudioRecorder>();
            _status = _services.GetRequiredService<ITrayStatusService>();
            _sounds = _services.GetRequiredService<IFeedbackSoundService>();
            _clipboardCapture = _services.GetRequiredService<IClipboardSourceCapture>();

            _hotkeys.HotkeyDown += OnHotkeyDown;
            _hotkeys.HotkeyUp += OnHotkeyUp;

            var settingsService = _services.GetRequiredService<ISettingsService>();
            var settings = await settingsService.LoadAsync();
            var readiness = settingsService.Validate(settings);
            var startHidden = settings.LaunchMinimizedToTray && readiness.IsReady;

            if (startHidden)
            {
                // Off-Screen positionieren und aus Switcher/Taskbar ausblenden, damit Activate
                // keinen sichtbaren Frame zeigt. HideToTray danach versteckt das Fenster vollständig.
                _window.AppWindow.IsShownInSwitchers = false;
                _window.AppWindow.Move(new Windows.Graphics.PointInt32(-32000, -32000));
            }

            _window.Activate();
            if (startHidden)
            {
                _window.HideToTray();
                _window.AppWindow.IsShownInSwitchers = true;
                MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: started hidden in tray");
            }
            MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: window activated");
            await _window.InitializeAfterActivationAsync();
            MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: window initialized");

            await _hotkeys.RegisterAsync(settings);
            MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: hotkeys registered");
            _tray = ActivatorUtilities.CreateInstance<TrayIconController>(_services, _window);
            MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: tray created");

            if (!readiness.IsReady)
            {
                _status.SetStatus(TrayStatus.ConfigurationRequired, "Einrichtung erforderlich. Bitte prüfe die markierten Einstellungen.");
                _window.Activate();
                MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: configuration required");
            }
            else
            {
                _status.SetStatus(TrayStatus.Idle, "Bereit. Halte ein Tastenkürzel gedrückt, um zu diktieren.");
            }

            MagicVoiceStartupCrashLogger.WriteMessage("OnLaunched: complete");
        }
        catch (Exception ex)
        {
            MagicVoiceStartupCrashLogger.Write(ex);
            Exit();
        }
    }

    private static ServiceProvider ConfigureServices()
    {
        var services = new ServiceCollection();
        services.AddLogging(builder => builder.AddDebug());
        services.AddSingleton<IAppProfile, MagicVoiceProfile>();
        services.AddSingleton<ISettingsService, SettingsService>();
        services.AddSingleton<ISecretProtector, DpapiSecretProtector>();
        services.AddSingleton<ITrayStatusService, InMemoryTrayStatusService>();
        services.AddSingleton<IProcessingFailureLog, InMemoryProcessingFailureLog>();
        services.AddSingleton<IHotkeyService, LowLevelKeyboardHotkeyService>();
        services.AddSingleton<IAudioDeviceService, NAudioDeviceService>();
        services.AddSingleton<IAudioRecorder, NAudioRecorder>();
        services.AddSingleton<IInputInjector, ClipboardInputInjector>();
        services.AddSingleton<IAutostartService, WindowsAutostartService>();
        services.AddSingleton<IFeedbackSoundService, NAudioFeedbackSoundService>();
        services.AddSingleton<IClipboardSourceCapture, WindowsClipboardSourceCapture>();
        services.AddSingleton<FileLogger>();
        services.AddHttpClient<ISttService, OpenAiCompatibleSttService>();
        services.AddHttpClient<ILlmService, OpenAiCompatibleLlmService>();
        services.AddSingleton<SpeechPipeline>();
        return services.BuildServiceProvider();
    }

    private async void OnHotkeyDown(object? sender, HotkeyPressedEventArgs e)
    {
        if (_status?.CurrentStatus is TrayStatus.Paused or TrayStatus.Processing or TrayStatus.Recording)
        {
            return;
        }

        var assistant = await ResolveAssistantAsync(e.AssistantId);
        if (assistant is null)
        {
            return;
        }

        var typeDefinition = _profile?.Modes.FirstOrDefault(m => m.Mode == assistant.Type);
        if (typeDefinition is { RequiresClipboardSource: true })
        {
            var clipboard = _clipboardCapture?.TryGetText();
            if (string.IsNullOrWhiteSpace(clipboard))
            {
                var label = string.IsNullOrWhiteSpace(assistant.Name) ? typeDefinition.Name : assistant.Name;
                _status?.SetStatus(TrayStatus.Error, $"{label} benötigt Text in der Zwischenablage. Es wurde nichts aufgenommen.");
                _capturedSourceText = null;
                return;
            }

            _capturedSourceText = clipboard;
        }
        else
        {
            _capturedSourceText = null;
        }

        try
        {
            _status?.SetStatus(TrayStatus.Recording, "Aufnahme läuft … Loslassen, um zu verarbeiten.");
            await _recorder!.StartAsync();
            _sounds?.PlayRecordingStart();
        }
        catch (InvalidOperationException)
        {
            _status?.SetStatus(TrayStatus.Error, "Mikrofonzugriff fehlgeschlagen. Bitte prüfe Mikrofon und Windows-Datenschutzeinstellungen.");
            _capturedSourceText = null;
        }
    }

    private async void OnHotkeyUp(object? sender, HotkeyPressedEventArgs e)
    {
        if (_pipeline is null)
        {
            return;
        }

        var sourceText = _capturedSourceText;
        _capturedSourceText = null;

        _sounds?.PlayRecordingStop();
        await _pipeline.RunAsync(e.AssistantId, sourceText);
    }

    private async Task<AssistantInstance?> ResolveAssistantAsync(string assistantId)
    {
        if (_services is null)
        {
            return null;
        }

        var settings = await _services.GetRequiredService<ISettingsService>().LoadAsync();
        return settings.Assistants.FirstOrDefault(a => string.Equals(a.Id, assistantId, StringComparison.Ordinal));
    }

    public async Task ShutdownAsync()
    {
        _window?.PrepareForExit();
        if (_window is not null)
        {
            await _window.SaveWindowBoundsAsync();
        }

        if (_hotkeys is not null)
        {
            await _hotkeys.DisposeAsync();
        }

        _tray?.Dispose();
        if (_services is not null)
        {
            await _services.DisposeAsync();
        }

        _singleInstanceMutex?.ReleaseMutex();
        _singleInstanceMutex?.Dispose();
    }
}

public sealed class TrayIconController : IDisposable
{
    private const int TrayId = 1;
    private const int TrayCallbackMessage = 0x8000 + 42;
    private const int WmCommand = 0x0111;
    private const int WmDestroy = 0x0002;
    private const int WmLButtonDblClk = 0x0203;
    private const int WmRButtonUp = 0x0205;
    private const uint NimAdd = 0x00000000;
    private const uint NimModify = 0x00000001;
    private const uint NimDelete = 0x00000002;
    private const uint NifMessage = 0x00000001;
    private const uint NifIcon = 0x00000002;
    private const uint NifTip = 0x00000004;
    private const uint NotifyIconVersion4 = 4;
    private const uint ImageIcon = 1;
    private const uint LrLoadFromFile = 0x00000010;
    private const uint LrDefaultSize = 0x00000040;
    private const int IdiApplication = 32512;
    private const int MfString = 0x00000000;
    private const int MfSeparator = 0x00000800;
    private const int MfGrayed = 0x00000001;
    private const int MfChecked = 0x00000008;
    private const int MfDefault = 0x00001000;
    private const uint TpmRightButton = 0x0002;
    private const uint TpmReturnCmd = 0x0100;
    private const int CmdOpen = 1000;
    private const int CmdActive = 1001;
    private const int CmdExit = 1004;

    private readonly IAppProfile _profile;
    private readonly MainWindow _window;
    private readonly ITrayStatusService _status;
    private readonly IHotkeyService _hotkeys;
    private readonly ISettingsService _settingsService;
    private readonly WndProc _wndProc;
    private readonly IntPtr _hwnd;
    private IntPtr _icon;
    private bool _ownsIcon;
    private bool _disposed;

    public TrayIconController(IAppProfile profile, MainWindow window, ITrayStatusService status, IHotkeyService hotkeys, ISettingsService settingsService)
    {
        _profile = profile;
        _window = window;
        _status = status;
        _hotkeys = hotkeys;
        _settingsService = settingsService;
        _wndProc = WindowProc;
        _hwnd = CreateMessageWindow(_wndProc);
        if (_hwnd == IntPtr.Zero)
        {
            throw new InvalidOperationException("Tray-Fenster konnte nicht erstellt werden.");
        }

        AddOrUpdateIcon(TrayStatus.ConfigurationRequired, $"{_profile.AppName} - Einrichtung erforderlich", NimAdd);
        _status.StatusChanged += OnStatusChanged;
    }

    private void OnStatusChanged(object? sender, TrayStatusChangedEventArgs args) => Update(args.Status, args.Message);

    private async Task ToggleActiveAsync()
    {
        if (_status.CurrentStatus == TrayStatus.Recording)
        {
            return;
        }

        if (_status.CurrentStatus == TrayStatus.Paused)
        {
            _hotkeys.Resume();
            var readiness = _settingsService.Validate(await _settingsService.LoadAsync());
            _status.SetStatus(
                readiness.IsReady ? TrayStatus.Idle : TrayStatus.ConfigurationRequired,
                readiness.IsReady
                    ? "Bereit. Halte ein Tastenkürzel gedrückt, um zu diktieren."
                    : "Einrichtung erforderlich. Bitte prüfe die markierten Einstellungen.");
            return;
        }

        _hotkeys.Pause();
        if (_status.CurrentStatus != TrayStatus.Processing)
        {
            _status.SetStatus(TrayStatus.Paused, "Inaktiv. Aktiviere die App im Tray-Menü, um Tastenkürzel zu nutzen.");
        }
    }

    private void Update(TrayStatus status, string message)
    {
        AddOrUpdateIcon(status, TooltipFor(status, message), NimModify);
    }

    private void AddOrUpdateIcon(TrayStatus status, string tooltip, uint message)
    {
        var newIcon = LoadIconFor(status, out var ownsNewIcon);
        var data = new NotifyIconData
        {
            cbSize = Marshal.SizeOf<NotifyIconData>(),
            hWnd = _hwnd,
            uID = TrayId,
            uFlags = NifMessage | NifIcon | NifTip,
            uCallbackMessage = TrayCallbackMessage,
            hIcon = newIcon,
            szTip = TrimTooltip(tooltip),
            uVersionOrTimeout = NotifyIconVersion4
        };

        if (!Shell_NotifyIcon(message, ref data))
        {
            DestroyIconIfOwned(newIcon, ownsNewIcon);
            return;
        }

        DestroyIconIfOwned(_icon, _ownsIcon);
        _icon = newIcon;
        _ownsIcon = ownsNewIcon;
    }

    private IntPtr WindowProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam)
    {
        if (msg == TrayCallbackMessage)
        {
            var mouseMessage = lParam.ToInt32();
            if (mouseMessage == WmLButtonDblClk)
            {
                _window.ShowFromTray();
                return IntPtr.Zero;
            }

            if (mouseMessage == WmRButtonUp)
            {
                ShowContextMenu();
                return IntPtr.Zero;
            }
        }

        if (msg == WmCommand)
        {
            HandleCommand(wParam.ToInt32() & 0xFFFF);
            return IntPtr.Zero;
        }

        if (msg == WmDestroy)
        {
            return IntPtr.Zero;
        }

        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    private void ShowContextMenu()
    {
        var menu = CreatePopupMenu();
        AppendMenu(menu, MfString | MfDefault, CmdOpen, "Öffnen");
        AppendMenu(menu, MfSeparator, 0, string.Empty);
        AppendMenu(menu, MfString | (_hotkeys.IsPaused ? 0 : MfChecked) | (_status.CurrentStatus == TrayStatus.Recording ? MfGrayed : 0), CmdActive, "aktiv");
        AppendMenu(menu, MfSeparator, 0, string.Empty);
        AppendMenu(menu, MfString, CmdExit, "Beenden");
        GetCursorPos(out var point);
        SetForegroundWindow(_hwnd);
        var command = TrackPopupMenu(menu, TpmRightButton | TpmReturnCmd, point.X, point.Y, 0, _hwnd, IntPtr.Zero);
        DestroyMenu(menu);
        if (command != 0)
        {
            HandleCommand(command);
        }
    }

    private void HandleCommand(int command)
    {
        switch (command)
        {
            case CmdOpen:
                _window.ShowFromTray();
                break;
            case CmdActive:
                _ = ToggleActiveAsync();
                break;
            case CmdExit:
                _ = ExitAsync();
                break;
        }
    }

    private string TooltipFor(TrayStatus status, string message) => status switch
    {
        TrayStatus.Idle => $"{_profile.AppName} - bereit",
        TrayStatus.Paused => $"{_profile.AppName} - inaktiv",
        TrayStatus.Recording => $"{_profile.AppName} - Aufnahme läuft",
        TrayStatus.Processing => $"{_profile.AppName} - verarbeitet den Text",
        TrayStatus.Success => $"{_profile.AppName} - Text eingefügt",
        TrayStatus.Error => $"{_profile.AppName} - {message}",
        TrayStatus.ConfigurationRequired => $"{_profile.AppName} - Einrichtung erforderlich",
        _ => _profile.AppName
    };

    private static string IconPathFor(TrayStatus status)
    {
        var fileName = status switch
        {
            TrayStatus.Paused => "TrayPaused.ico",
            TrayStatus.Recording => "TrayRecording.ico",
            TrayStatus.Processing => "TrayProcessing.ico",
            TrayStatus.Success => "TraySuccess.ico",
            TrayStatus.Error => "TrayError.ico",
            TrayStatus.ConfigurationRequired => "TrayConfigurationRequired.ico",
            _ => "TrayIdle.ico"
        };
        return Path.Combine(AppContext.BaseDirectory, "Assets", fileName);
    }

    private static IntPtr LoadIconFor(TrayStatus status, out bool ownsIcon)
    {
        var path = IconPathFor(status);
        var icon = File.Exists(path)
            ? LoadImage(IntPtr.Zero, path, ImageIcon, 0, 0, LrLoadFromFile | LrDefaultSize)
            : IntPtr.Zero;
        if (icon != IntPtr.Zero)
        {
            ownsIcon = true;
            return icon;
        }

        ownsIcon = false;
        return LoadIcon(IntPtr.Zero, new IntPtr(IdiApplication));
    }

    private static string TrimTooltip(string tooltip) => tooltip.Length > 127 ? tooltip[..127] : tooltip;

    private static void DestroyIconIfOwned(IntPtr icon, bool ownsIcon)
    {
        if (icon != IntPtr.Zero && ownsIcon)
        {
            DestroyIcon(icon);
        }
    }

    private static async Task ExitAsync()
    {
        if (Microsoft.UI.Xaml.Application.Current is App app)
        {
            await app.ShutdownAsync();
        }

        Microsoft.UI.Xaml.Application.Current.Exit();
    }

    private static IntPtr CreateMessageWindow(WndProc wndProc)
    {
        var className = "MagicVoiceTrayWindow_" + Environment.ProcessId.ToString(CultureInfo.InvariantCulture);
        var moduleHandle = GetModuleHandle(null);
        var windowClass = new WindowClass
        {
            cbSize = (uint)Marshal.SizeOf<WindowClass>(),
            lpfnWndProc = Marshal.GetFunctionPointerForDelegate(wndProc),
            hInstance = moduleHandle,
            lpszClassName = className
        };
        if (RegisterClassEx(ref windowClass) == 0)
        {
            MagicVoiceStartupCrashLogger.WriteMessage($"Tray RegisterClassEx failed: {Marshal.GetLastWin32Error()}");
            return IntPtr.Zero;
        }

        var hwnd = CreateWindowEx(0, className, className, 0, 0, 0, 0, 0, IntPtr.Zero, IntPtr.Zero, moduleHandle, IntPtr.Zero);
        if (hwnd == IntPtr.Zero)
        {
            MagicVoiceStartupCrashLogger.WriteMessage($"Tray CreateWindowEx failed: {Marshal.GetLastWin32Error()}");
        }

        return hwnd;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _status.StatusChanged -= OnStatusChanged;
        var data = new NotifyIconData
        {
            cbSize = Marshal.SizeOf<NotifyIconData>(),
            hWnd = _hwnd,
            uID = TrayId
        };
        Shell_NotifyIcon(NimDelete, ref data);
        DestroyIconIfOwned(_icon, _ownsIcon);
        if (_hwnd != IntPtr.Zero)
        {
            DestroyWindow(_hwnd);
        }
    }

    private delegate IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NotifyIconData
    {
        public int cbSize;
        public IntPtr hWnd;
        public uint uID;
        public uint uFlags;
        public uint uCallbackMessage;
        public IntPtr hIcon;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szTip;
        public uint dwState;
        public uint dwStateMask;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string szInfo;
        public uint uVersionOrTimeout;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string szInfoTitle;
        public uint dwInfoFlags;
        public Guid guidItem;
        public IntPtr hBalloonIcon;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Point
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WindowClass
    {
        public uint cbSize;
        public uint style;
        public IntPtr lpfnWndProc;
        public int cbClsExtra;
        public int cbWndExtra;
        public IntPtr hInstance;
        public IntPtr hIcon;
        public IntPtr hCursor;
        public IntPtr hbrBackground;
        public string? lpszMenuName;
        public string lpszClassName;
        public IntPtr hIconSm;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern bool Shell_NotifyIcon(uint dwMessage, ref NotifyIconData lpData);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr LoadImage(IntPtr hinst, string lpszName, uint uType, int cxDesired, int cyDesired, uint fuLoad);

    [DllImport("user32.dll")]
    private static extern IntPtr LoadIcon(IntPtr hInstance, IntPtr lpIconName);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern ushort RegisterClassEx(ref WindowClass lpWndClass);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateWindowEx(
        int dwExStyle,
        string lpClassName,
        string lpWindowName,
        int dwStyle,
        int x,
        int y,
        int nWidth,
        int nHeight,
        IntPtr hWndParent,
        IntPtr hMenu,
        IntPtr hInstance,
        IntPtr lpParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr DefWindowProc(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr CreatePopupMenu();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern bool AppendMenu(IntPtr hMenu, int uFlags, int uIDNewItem, string lpNewItem);

    [DllImport("user32.dll")]
    private static extern bool DestroyMenu(IntPtr hMenu);

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out Point lpPoint);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int TrackPopupMenu(IntPtr hMenu, uint uFlags, int x, int y, int nReserved, IntPtr hWnd, IntPtr prcRect);
}

internal static class MagicVoiceStartupCrashLogger
{
    public static void WriteMessage(string message)
    {
        try
        {
            var logPath = LogPath();
            Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
            File.AppendAllText(logPath, $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}");
        }
        catch
        {
            // Startup logging must never create another startup failure.
        }
    }

    public static void Write(Exception exception)
    {
        try
        {
            var logPath = LogPath();
            Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
            File.AppendAllText(
                logPath,
                $"{DateTimeOffset.Now:O}{Environment.NewLine}{exception}{Environment.NewLine}{Environment.NewLine}");
        }
        catch
        {
            // Startup logging must never create another startup failure.
        }
    }

    private static string LogPath() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Magic-Voice",
        "logs",
        "startup-crash.log");
}
