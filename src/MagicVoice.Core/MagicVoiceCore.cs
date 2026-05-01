using System.Globalization;
using System.Text;
using System.Text.Json.Serialization;

namespace MagicVoice.Core;

public enum AssistantMode
{
    /// <summary>Transkript + Anweisung aus der UI (Korrektur, Umschreiben, Stil usw.).</summary>
    Transform,
    /// <summary>Nur gesprochene Anweisung, kein Quelltext.</summary>
    Generate,
    /// <summary>Zwischenablage als Quelltext + gesprochene Anweisung.</summary>
    AnswerClipboard
}

/// <summary>Wie stark der Text mit Absätzen gegliedert werden soll (UI).</summary>
public enum ParagraphDensity
{
    Compact,
    Balanced,
    Spacious
}

public enum WritingStyle
{
    Casual,
    Neutral,
    Professional,
    Academic
}

/// <summary>Wie stark Emojis und „Social“-Ausdruck genutzt werden sollen (UI).</summary>
public enum EmojiExpression
{
    None,
    Sparse,
    Balanced,
    Lively,
    Heavy
}

public enum TrayStatus
{
    Idle,
    Paused,
    Recording,
    Processing,
    Success,
    Error,
    ConfigurationRequired
}

public enum InsertMethod
{
    Clipboard,
    SendInput
}

public sealed record AssistantModeDefinition(
    AssistantMode Mode,
    string Name,
    string Description,
    string DefaultHotkey,
    string DefaultPrompt,
    bool RequiresClipboardSource = false);

public sealed record AudioBuffer(byte[] PcmBytes, int SampleRate, int Channels, TimeSpan Duration)
{
    public bool IsEmpty => PcmBytes.Length == 0 || Duration <= TimeSpan.Zero;
}

public sealed record AudioInputDevice(string Id, string Name, bool IsDefault)
{
    public override string ToString() => Name;
}

public sealed record LanguageOption(string Code, string Name)
{
    public override string ToString() => Name;
}

public sealed record SttRequest(AudioBuffer Audio, string Provider, string Endpoint, string Model, string Language, string ApiKey);

public sealed record LlmRequest(string Transcript, AssistantMode Mode, string Provider, string Endpoint, string Model, string ApiKey, string SystemPrompt, string ModePrompt, string? SourceText = null);

public sealed record PipelineResult(bool Success, string Message, string? Transcript = null, string? FinalText = null, Exception? Error = null)
{
    public static PipelineResult Ok(string message, string? transcript, string? finalText) => new(true, message, transcript, finalText);

    public static PipelineResult Failed(string message, Exception? error = null, string? transcript = null, string? finalText = null) =>
        new(false, message, transcript, finalText, error);
}

public sealed class AppSettings
{
    public string SttProvider { get; set; } = Defaults.OpenAiProviderName;
    public string SttModel { get; set; } = "gpt-4o-mini-transcribe";
    public string InputLanguage { get; set; } = "de";
    public string OutputLanguage { get; set; } = "same";
    public string AudioInputDeviceId { get; set; } = Defaults.DefaultAudioInputDeviceId;
    public string? SttApiKeyEncrypted { get; set; }
    public string LlmProvider { get; set; } = Defaults.OpenAiProviderName;
    public string LlmModel { get; set; } = Defaults.DefaultLlmModel;
    public string? LlmApiKeyEncrypted { get; set; }
    public int RecordingMaxSeconds { get; set; } = 60;
    public int MinimumRecordingMilliseconds { get; set; } = 300;
    public int ProcessingTimeoutSeconds { get; set; } = 45;
    public InsertMethod InsertMethod { get; set; } = InsertMethod.Clipboard;
    public bool RestoreClipboard { get; set; } = true;
    public string LogLevel { get; set; } = "Information";
    public bool LaunchMinimizedToTray { get; set; } = true;
    public bool MinimizeToTray { get; set; } = true;
    public bool PlayRecordingSounds { get; set; } = true;
    public bool ShowNotifications { get; set; } = true;
    public bool OpenSettingsOnConfigurationError { get; set; } = true;
    public string LastSelectedSettingsSection { get; set; } = "overview";
    public WindowBounds WindowBounds { get; set; } = new();
    public List<AssistantInstance> Assistants { get; set; } = new();

    [JsonIgnore]
    public bool HasEncryptedLlmApiKey => !string.IsNullOrWhiteSpace(LlmApiKeyEncrypted);

    [JsonIgnore]
    public bool HasEncryptedSttApiKey => !string.IsNullOrWhiteSpace(SttApiKeyEncrypted);
}

public sealed class AssistantInstance
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    [JsonConverter(typeof(AssistantModeJsonConverter))]
    public AssistantMode Type { get; set; } = AssistantMode.Transform;
    public string Name { get; set; } = string.Empty;
    public string Hotkey { get; set; } = string.Empty;
    public string Prompt { get; set; } = string.Empty;
    /// <summary>
    /// Optionaler Override für die System-Nachricht an das LLM.
    /// Null/leer bedeutet: globaler Standard (aus dem App-Profil).
    /// </summary>
    public string? SystemPromptOverride { get; set; }
    /// <summary>
    /// Optionaler Override für die STT-Eingabesprache (z. B. "de", "en", "auto").
    /// Null/leer bedeutet: globaler Standard aus den App-Einstellungen.
    /// </summary>
    public string? InputLanguageOverride { get; set; }
    /// <summary>
    /// Optionaler Override für die LLM-Ausgabesprache (z. B. "same", "de", "en").
    /// Null/leer bedeutet: globaler Standard aus den App-Einstellungen.
    /// </summary>
    public string? OutputLanguageOverride { get; set; }
    public int Intensity { get; set; } = Defaults.DefaultModeIntensity;
    public WritingStyle WritingStyle { get; set; } = WritingStyle.Neutral;
    /// <summary>Absatz- und Leerzeilennutzung für die KI (UI).</summary>
    public ParagraphDensity ParagraphDensity { get; set; } = ParagraphDensity.Balanced;
    /// <summary>Emoji- und Social-Media-Ausdrucksstärke (UI).</summary>
    public EmojiExpression EmojiExpression { get; set; } = EmojiExpression.Balanced;
}

public static class PromptComposition
{
    public static string EffectiveInputLanguage(AppSettings settings, AssistantInstance assistant) =>
        string.IsNullOrWhiteSpace(assistant.InputLanguageOverride) ? settings.InputLanguage : assistant.InputLanguageOverride!;

    public static string EffectiveOutputLanguage(AppSettings settings, AssistantInstance assistant) =>
        string.IsNullOrWhiteSpace(assistant.OutputLanguageOverride) ? settings.OutputLanguage : assistant.OutputLanguageOverride!;

    public static string EffectiveBaseSystemPrompt(IAppProfile profile, AssistantInstance assistant) =>
        string.IsNullOrWhiteSpace(assistant.SystemPromptOverride) ? profile.SystemPrompt : assistant.SystemPromptOverride!;

    public static string BuildSystemPrompt(string baseSystemPrompt, string effectiveInputLanguage, string? policyBlock = null)
    {
        var inputLanguageNote = Defaults.IsAutoLanguage(effectiveInputLanguage)
            ? "Die Eingabesprache ist nicht fest vorgegeben; orientiere dich an der erkannten Transkription und am Kontext."
            : $"Die vorgesehene Eingabesprache ist {Defaults.LanguageName(effectiveInputLanguage)}.";
        var baseWithLanguage = $"{baseSystemPrompt.TrimEnd()} {inputLanguageNote}";
        if (string.IsNullOrWhiteSpace(policyBlock))
        {
            return baseWithLanguage;
        }

        return string.Join(Environment.NewLine + Environment.NewLine, baseWithLanguage, policyBlock.Trim());
    }

    public static string OutputInstruction(string effectiveInputLanguage, string effectiveOutputLanguage)
    {
        if (effectiveOutputLanguage.Equals(Defaults.SameAsInputLanguageCode, StringComparison.OrdinalIgnoreCase))
        {
            return Defaults.IsAutoLanguage(effectiveInputLanguage)
                ? "Gib den finalen Text in der erkannten Eingabesprache aus."
                : $"Gib den finalen Text auf {Defaults.LanguageName(effectiveInputLanguage)} aus.";
        }

        return $"Gib den finalen Text auf {Defaults.LanguageName(effectiveOutputLanguage)} aus. Übersetze den Inhalt bei Bedarf natürlich und ohne Erklärung.";
    }

    public static string BuildPolicyBlock(IAppProfile profile, AssistantInstance assistant, string effectiveInputLanguage, string effectiveOutputLanguage)
    {
        var intensity = Math.Clamp(assistant.Intensity, Defaults.MinModeIntensity, Defaults.MaxModeIntensity);
        var intensityInstruction = profile.IntensityStepInstruction(assistant.Type, intensity);
        var intensityName = profile.IntensityStepName(assistant.Type, intensity);
        var intensityBlock = string.IsNullOrWhiteSpace(intensityInstruction)
            ? string.Empty
            : $"Intensität {intensity}/{Defaults.MaxModeIntensity} ({intensityName}): {intensityInstruction}";

        var styleBlock = profile.WritingStyleInstruction(assistant.WritingStyle);
        var emojiBlock = profile.EmojiExpressionInstruction(assistant.EmojiExpression);
        var formattingBlock = FormattingInstruction(assistant.ParagraphDensity);
        var outputBlock = OutputInstruction(effectiveInputLanguage, effectiveOutputLanguage);

        return string.Join(
            Environment.NewLine + Environment.NewLine,
            new[] { intensityBlock, styleBlock, emojiBlock, formattingBlock, outputBlock }.Where(part => !string.IsNullOrWhiteSpace(part)));
    }

    public static string FormattingInstruction(ParagraphDensity density) => density switch
    {
        ParagraphDensity.Compact =>
            "Halte den Text möglichst zusammenhängend; vermeide zusätzliche Absätze und Leerzeilen, außer sie sind inhaltlich nötig. Gib nur den finalen Text zurück.",
        ParagraphDensity.Spacious =>
            "Strukturiere längere Texte mit klaren Absätzen und ggf. Leerzeilen zwischen Abschnitten, wenn es die Lesbarkeit verbessert. Gib nur den finalen Text zurück.",
        _ =>
            "Setze sinnvolle Absätze und Zeilenumbrüche, wenn es der Lesbarkeit dient. Verwende keine künstlichen Zeilenumbrüche in sehr kurzen Texten. Gib nur den finalen Text zurück."
    };
}

public sealed class WindowBounds
{
    public double Width { get; set; } = 900;
    public double Height { get; set; } = 650;
    public double? X { get; set; }
    public double? Y { get; set; }
    public bool IsSet => X.HasValue && Y.HasValue && Width > 0 && Height > 0;
}

public sealed record ValidationIssue(string Field, string Message);

public sealed record ReadinessReport(TrayStatus Status, IReadOnlyList<ValidationIssue> Issues)
{
    public bool IsReady => Status == TrayStatus.Idle && Issues.Count == 0;
}

public sealed record HotkeyGesture(bool Control, bool Alt, bool Shift, bool Windows, string Key)
{
    public override string ToString()
    {
        var parts = new List<string>();
        if (Control) parts.Add("Ctrl");
        if (Alt) parts.Add("Alt");
        if (Shift) parts.Add("Shift");
        if (Windows) parts.Add("Win");
        parts.Add(Key);
        return string.Join("+", parts);
    }
}

/// <summary>App-spezifische Identität, verfügbare Assistenten-Typen und Texte. Wird per DI aus dem App-Projekt bereitgestellt.</summary>
public interface IAppProfile
{
    string AppName { get; }
    string DataFolderName { get; }
    string AuthorName { get; }
    string CopyrightText { get; }
    string LicenseName { get; }
    string MutexName { get; }
    string AumId { get; }
    string AutostartRegistryValueName { get; }
    string SystemPrompt { get; }
    /// <summary>Verfügbare Assistenten-Typen mit Default-Werten als Vorlage beim Hinzufügen neuer Assistenten.</summary>
    IReadOnlyList<AssistantModeDefinition> Modes { get; }
    string IntensityStepName(AssistantMode mode, int intensity);
    string IntensityStepInstruction(AssistantMode mode, int intensity);
    string WritingStyleInstruction(WritingStyle style);
    string EmojiExpressionInstruction(EmojiExpression level);
    /// <summary>Liste der Assistenten beim ersten Start (oder beim Reset).</summary>
    List<AssistantInstance> CreateDefaultAssistants();
}

public static class Defaults
{
    public const string OpenAiProviderName = "OpenAI";
    public const string DefaultAudioInputDeviceId = "default";
    public const string AutoLanguageCode = "auto";
    public const string SameAsInputLanguageCode = "same";
    public const int DefaultModeIntensity = 3;
    public const int MinModeIntensity = 1;
    public const int MaxModeIntensity = 5;
    /// <summary>Empfohlenes Standard-KI-Modell (Chat Completions); auch frei überschreibbar in der App.</summary>
    public const string DefaultLlmModel = "gpt-4o";
    private const string OpenAiCompatibleSttEndpoint = "https://api.openai.com/v1/audio/transcriptions";
    private const string OpenAiCompatibleLlmEndpoint = "https://api.openai.com/v1/chat/completions";
    public static readonly IReadOnlyList<string> KnownProviders = [OpenAiProviderName];
    public static readonly IReadOnlyList<string> OpenAiSttModels = ["gpt-4o-mini-transcribe", "gpt-4o-transcribe"];
    /// <summary>Bekannte Chat-Completions-Modelle (Auswahl); jede andere Modell-ID kann manuell eingetragen werden.</summary>
    public static readonly IReadOnlyList<string> OpenAiLlmModels =
    [
        "gpt-5.4-nano",
        "gpt-5.4-mini",
        "gpt-5-mini",
        "gpt-5-nano",
        "gpt-5.4",
        "gpt-5.4-pro",
        "gpt-5.5",
        "gpt-5.5-pro",
        "gpt-5",
        "gpt-5.1",
        "gpt-5.2",
        "gpt-5.3",
        "gpt-4o-mini",
        "gpt-4o",
        "gpt-4.1-nano",
        "gpt-4.1-mini",
        "gpt-4.1",
        "gpt-4-turbo",
        "gpt-4",
        "gpt-3.5-turbo"
    ];

    /// <summary>Sprachen mit Whisper/OpenAI-STT-kompatiblen Codes und Anzeigenamen (alphabetisch).</summary>
    private static readonly LanguageOption[] SelectableLanguages =
    [
        new("af", "Afrikaans"),
        new("sq", "Albanisch"),
        new("am", "Amharisch"),
        new("ar", "Arabisch"),
        new("hy", "Armenisch"),
        new("as", "Assamesisch"),
        new("az", "Aserbaidschanisch"),
        new("ba", "Baschkirisch"),
        new("eu", "Baskisch"),
        new("be", "Weißrussisch"),
        new("bn", "Bengalisch"),
        new("my", "Birmanisch"),
        new("bs", "Bosnisch"),
        new("br", "Bretonisch"),
        new("bg", "Bulgarisch"),
        new("zh", "Chinesisch (Mandarin)"),
        new("da", "Dänisch"),
        new("de", "Deutsch"),
        new("en", "Englisch"),
        new("et", "Estnisch"),
        new("fo", "Färöisch"),
        new("tl", "Filipino"),
        new("fi", "Finnisch"),
        new("fr", "Französisch"),
        new("gl", "Galicisch"),
        new("ka", "Georgisch"),
        new("el", "Griechisch"),
        new("gu", "Gujarati"),
        new("ht", "Haitianisch"),
        new("ha", "Hausa"),
        new("haw", "Hawaiisch"),
        new("he", "Hebräisch"),
        new("hi", "Hindi"),
        new("id", "Indonesisch"),
        new("is", "Isländisch"),
        new("it", "Italienisch"),
        new("ja", "Japanisch"),
        new("jw", "Javanisch"),
        new("yi", "Jiddisch"),
        new("kn", "Kannada"),
        new("yue", "Kantonesisch"),
        new("kk", "Kasachisch"),
        new("ca", "Katalanisch"),
        new("km", "Khmer"),
        new("ko", "Koreanisch"),
        new("hr", "Kroatisch"),
        new("lo", "Laotisch"),
        new("la", "Lateinisch"),
        new("lv", "Lettisch"),
        new("ln", "Lingala"),
        new("lt", "Litauisch"),
        new("lb", "Luxemburgisch"),
        new("mg", "Madagassisch"),
        new("ms", "Malaiisch"),
        new("ml", "Malayalam"),
        new("mt", "Maltesisch"),
        new("mi", "Maori"),
        new("mr", "Marathi"),
        new("mk", "Mazedonisch"),
        new("mn", "Mongolisch"),
        new("ne", "Nepalesisch"),
        new("nl", "Niederländisch"),
        new("nn", "Norwegisch (Nynorsk)"),
        new("no", "Norwegisch (Bokmål)"),
        new("oc", "Okzitanisch"),
        new("ps", "Paschtu"),
        new("fa", "Persisch"),
        new("pl", "Polnisch"),
        new("pt", "Portugiesisch"),
        new("pa", "Punjabi"),
        new("ro", "Rumänisch"),
        new("ru", "Russisch"),
        new("sa", "Sanskrit"),
        new("sv", "Schwedisch"),
        new("sr", "Serbisch"),
        new("sn", "Shona"),
        new("si", "Singhalesisch"),
        new("sd", "Sindhi"),
        new("sk", "Slowakisch"),
        new("sl", "Slowenisch"),
        new("so", "Somali"),
        new("es", "Spanisch"),
        new("sw", "Suaheli"),
        new("su", "Sundanesisch"),
        new("ta", "Tamil"),
        new("tg", "Tadschikisch"),
        new("tt", "Tatarisch"),
        new("te", "Telugu"),
        new("th", "Thai"),
        new("bo", "Tibetisch"),
        new("cs", "Tschechisch"),
        new("tr", "Türkisch"),
        new("tk", "Turkmenisch"),
        new("uk", "Ukrainisch"),
        new("hu", "Ungarisch"),
        new("ur", "Urdu"),
        new("uz", "Usbekisch"),
        new("vi", "Vietnamesisch"),
        new("cy", "Walisisch"),
        new("yo", "Yoruba")
    ];

    public static readonly IReadOnlyList<LanguageOption> InputLanguages =
    [
        new(AutoLanguageCode, "automatisch erkennen"),
        ..SelectableLanguages
    ];

    public static readonly IReadOnlyList<LanguageOption> OutputLanguages =
    [
        new(SameAsInputLanguageCode, "wie Eingabe"),
        ..SelectableLanguages
    ];

    public static bool TryGetSttEndpoint(string provider, out string endpoint) =>
        TryGetProviderEndpoint(provider, OpenAiCompatibleSttEndpoint, out endpoint);

    public static bool TryGetLlmEndpoint(string provider, out string endpoint) =>
        TryGetProviderEndpoint(provider, OpenAiCompatibleLlmEndpoint, out endpoint);

    public static bool IsAutoLanguage(string languageCode) =>
        string.IsNullOrWhiteSpace(languageCode) || languageCode.Equals(AutoLanguageCode, StringComparison.OrdinalIgnoreCase);

    public static string LanguageName(string languageCode)
    {
        var option = InputLanguages.Concat(OutputLanguages)
            .FirstOrDefault(language => language.Code.Equals(languageCode, StringComparison.OrdinalIgnoreCase));
        return option?.Name ?? languageCode;
    }

    private static bool TryGetProviderEndpoint(string provider, string endpoint, out string resolvedEndpoint)
    {
        if (provider.Equals(OpenAiProviderName, StringComparison.OrdinalIgnoreCase))
        {
            resolvedEndpoint = endpoint;
            return true;
        }

        resolvedEndpoint = string.Empty;
        return false;
    }
}

public interface ISettingsService
{
    string SettingsPath { get; }
    string DataDirectory { get; }
    string LogDirectory { get; }
    Task<AppSettings> LoadAsync(CancellationToken cancellationToken = default);
    Task SaveAsync(AppSettings settings, CancellationToken cancellationToken = default);
    ReadinessReport Validate(AppSettings settings);
}

public interface ISecretProtector
{
    string Protect(string secret);
    string Unprotect(string protectedSecret);
}

public interface IHotkeyService : IAsyncDisposable
{
    event EventHandler<HotkeyPressedEventArgs>? HotkeyDown;
    event EventHandler<HotkeyPressedEventArgs>? HotkeyUp;
    bool IsPaused { get; }
    Task<IReadOnlyList<ValidationIssue>> RegisterAsync(AppSettings settings, CancellationToken cancellationToken = default);
    void Pause();
    void Resume();
}

public sealed class HotkeyPressedEventArgs(string assistantId) : EventArgs
{
    public string AssistantId { get; } = assistantId;
}

public interface IAudioRecorder
{
    Task StartAsync(CancellationToken cancellationToken = default);
    Task<AudioBuffer> StopAsync(CancellationToken cancellationToken = default);
    Task AbortAsync(CancellationToken cancellationToken = default);
}

public interface IAudioDeviceService
{
    IReadOnlyList<AudioInputDevice> GetInputDevices();
}

public interface ISttService
{
    Task<string> TranscribeAsync(SttRequest request, CancellationToken cancellationToken = default);
}

public interface ILlmService
{
    Task<string> ProcessAsync(LlmRequest request, CancellationToken cancellationToken = default);
}

public interface IInputInjector
{
    Task InsertTextAsync(string text, InsertMethod method, bool restoreClipboard, CancellationToken cancellationToken = default);
}

public interface ITrayStatusService
{
    TrayStatus CurrentStatus { get; }
    string Message { get; }
    event EventHandler<TrayStatusChangedEventArgs>? StatusChanged;
    void SetStatus(TrayStatus status, string message);
}

public sealed record TrayStatusChangedEventArgs(TrayStatus Status, string Message);

/// <summary>Letzter Verarbeitungsfehler für die Diagnose-Ansicht (thread-sicher, kurz gehalten).</summary>
public interface IProcessingFailureLog
{
    void Record(string headline, Exception? exception = null);

    void Clear();

    string? LastEntry { get; }
}

public sealed class InMemoryProcessingFailureLog : IProcessingFailureLog
{
    private const int MaxChars = 4500;
    private readonly object _gate = new();
    private string? _last;

    public string? LastEntry
    {
        get
        {
            lock (_gate)
            {
                return _last;
            }
        }
    }

    public void Clear()
    {
        lock (_gate)
        {
            _last = null;
        }
    }

    public void Record(string headline, Exception? exception = null)
    {
        var sb = new StringBuilder();
        sb.Append('[').Append(DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)).AppendLine("]");
        sb.AppendLine(headline);
        if (exception is not null)
        {
            sb.Append(FormatExceptionForDiagnostics(exception));
        }

        var text = sb.ToString().TrimEnd();
        if (text.Length > MaxChars)
        {
            text = text[..MaxChars] + "…";
        }

        lock (_gate)
        {
            _last = text;
        }
    }

    private static string FormatExceptionForDiagnostics(Exception ex)
    {
        // Ziel: brauchbare Fehlersuche im Diagnose-Panel ohne riesige Dumps.
        // - Zeige pro Exception: Typ, Message (oder Fallback), relevante Properties (z.B. FileName)
        // - Zeige Stacktrace (gekürzt), weil ohne den die Ursache oft nicht nachvollziehbar ist.
        const int maxBodyChars = 3200;
        const int maxStackLines = 18;

        var sb = new StringBuilder();

        var index = 0;
        for (var e = ex; e is not null; e = e.InnerException)
        {
            if (index > 0)
            {
                sb.AppendLine("InnerException:");
            }

            sb.Append(e.GetType().Name).Append(": ");
            var message = string.IsNullOrWhiteSpace(e.Message) ? "(keine Message)" : e.Message.Trim();
            sb.AppendLine(message);

            if (e is FileNotFoundException fnf)
            {
                if (!string.IsNullOrWhiteSpace(fnf.FileName))
                {
                    sb.Append("FileName: ").AppendLine(fnf.FileName);
                }
            }
            else if (e is DirectoryNotFoundException)
            {
                // keine zusätzlichen Properties
            }
            else if (e is DllNotFoundException dll && !string.IsNullOrWhiteSpace(dll.Message))
            {
                // Message enthält oft bereits den DLL-Namen; bleibt oben sichtbar.
            }

            sb.Append("HResult: 0x").AppendLine(e.HResult.ToString("X8", CultureInfo.InvariantCulture));

            var stack = e.StackTrace;
            if (!string.IsNullOrWhiteSpace(stack))
            {
                sb.AppendLine("Stacktrace:");
                var lines = stack.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                for (var i = 0; i < Math.Min(lines.Length, maxStackLines); i++)
                {
                    sb.AppendLine(lines[i]);
                }

                if (lines.Length > maxStackLines)
                {
                    sb.AppendLine("…");
                }
            }

            sb.AppendLine();

            if (sb.Length > maxBodyChars)
            {
                sb.AppendLine("…");
                break;
            }

            index++;
        }

        return sb.ToString().TrimEnd() + Environment.NewLine;
    }
}

public sealed class NullProcessingFailureLog : IProcessingFailureLog
{
    public void Record(string headline, Exception? exception = null)
    {
    }

    public void Clear()
    {
    }

    public string? LastEntry => null;
}

public interface IAutostartService
{
    bool IsEnabled();
    void SetEnabled(bool enabled);
}

public interface IFeedbackSoundService
{
    void PlayRecordingStart();
    void PlayRecordingStop();
}

/// <summary>Liest den aktuellen Zwischenablage-Inhalt (Unicode-Text) für Modi, die einen Quelltext benötigen.</summary>
public interface IClipboardSourceCapture
{
    string? TryGetText();
}

public static class HotkeyParser
{
    private static readonly HashSet<string> Modifiers = new(StringComparer.OrdinalIgnoreCase)
    {
        "Ctrl", "Control", "Alt", "Shift", "Win", "Windows"
    };

    public static bool TryParse(string? text, out HotkeyGesture gesture, out string error)
    {
        gesture = new HotkeyGesture(false, false, false, false, string.Empty);
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(text))
        {
            error = "Tastenkürzel darf nicht leer sein.";
            return false;
        }

        var parts = text.Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
        {
            error = "Bitte mindestens eine Zusatztaste und eine Haupttaste verwenden.";
            return false;
        }

        var control = false;
        var alt = false;
        var shift = false;
        var windows = false;
        string? key = null;

        foreach (var part in parts)
        {
            if (Modifiers.Contains(part))
            {
                control |= part.Equals("Ctrl", StringComparison.OrdinalIgnoreCase) || part.Equals("Control", StringComparison.OrdinalIgnoreCase);
                alt |= part.Equals("Alt", StringComparison.OrdinalIgnoreCase);
                shift |= part.Equals("Shift", StringComparison.OrdinalIgnoreCase);
                windows |= part.Equals("Win", StringComparison.OrdinalIgnoreCase) || part.Equals("Windows", StringComparison.OrdinalIgnoreCase);
                continue;
            }

            if (key is not null)
            {
                error = "Nur eine Haupttaste ist erlaubt.";
                return false;
            }

            key = NormalizeKey(part);
        }

        if (!control && !alt && !shift && !windows)
        {
            error = "Bitte mindestens eine Zusatztaste wie Strg, Alt, Umschalt oder Win verwenden.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(key))
        {
            error = "Bitte eine Haupttaste angeben.";
            return false;
        }

        gesture = new HotkeyGesture(control, alt, shift, windows, key);
        return true;
    }

    private static string NormalizeKey(string key)
    {
        key = key.Trim();
        return key.Length == 1 ? key.ToUpperInvariant() : key;
    }
}

public sealed class PromptService(IAppProfile profile)
{
    public string GetPrompt(AssistantInstance assistant) =>
        !string.IsNullOrWhiteSpace(assistant.Prompt)
            ? assistant.Prompt
            : profile.Modes.FirstOrDefault(definition => definition.Mode == assistant.Type)?.DefaultPrompt ?? string.Empty;
}

public sealed class SpeechPipeline(
    IAppProfile profile,
    IAudioRecorder audioRecorder,
    ISttService sttService,
    ILlmService llmService,
    IInputInjector inputInjector,
    ISettingsService settingsService,
    ISecretProtector secretProtector,
    IHotkeyService hotkeyService,
    ITrayStatusService trayStatusService,
    IProcessingFailureLog failureLog)
{
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly PromptService _promptService = new(profile);

    private PipelineResult Fail(string message, Exception? ex = null, string? transcript = null, string? finalText = null)
    {
        failureLog.Record(message, ex);
        return PipelineResult.Failed(message, ex, transcript, finalText);
    }

    public async Task<PipelineResult> RunAsync(string assistantId, string? sourceText, CancellationToken cancellationToken = default)
    {
        if (!await _gate.WaitAsync(0, cancellationToken).ConfigureAwait(false))
        {
            trayStatusService.SetStatus(TrayStatus.Processing, "Text wird bereits verarbeitet...");
            return Fail("Eine Verarbeitung läuft bereits.");
        }

        try
        {
            var settings = await settingsService.LoadAsync(cancellationToken).ConfigureAwait(false);
            var readiness = settingsService.Validate(settings);
            if (!readiness.IsReady)
            {
                trayStatusService.SetStatus(TrayStatus.ConfigurationRequired, "Einrichtung erforderlich. Bitte prüfe die markierten Einstellungen.");
                return Fail("Einrichtung erforderlich.");
            }

            var assistant = settings.Assistants.FirstOrDefault(a => string.Equals(a.Id, assistantId, StringComparison.Ordinal));
            if (assistant is null)
            {
                trayStatusService.SetStatus(TrayStatus.Error, "Assistent nicht gefunden. Bitte Tastenkürzel neu zuweisen.");
                return Fail($"Assistent mit Id '{assistantId}' nicht in den Einstellungen gefunden.");
            }

            var effectiveInputLanguage = PromptComposition.EffectiveInputLanguage(settings, assistant);
            var effectiveOutputLanguage = PromptComposition.EffectiveOutputLanguage(settings, assistant);
            trayStatusService.SetStatus(TrayStatus.Processing, "Text wird verarbeitet...");
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(settings.ProcessingTimeoutSeconds));

            var audio = await audioRecorder.StopAsync(timeout.Token).ConfigureAwait(false);
            if (audio.IsEmpty || audio.Duration < TimeSpan.FromMilliseconds(settings.MinimumRecordingMilliseconds))
            {
                trayStatusService.SetStatus(TrayStatus.Idle, "Keine Sprache erkannt. Es wurde nichts eingefügt.");
                return Fail("Keine Sprache erkannt.");
            }

            var sttKey = secretProtector.Unprotect(settings.SttApiKeyEncrypted!);
            Defaults.TryGetSttEndpoint(settings.SttProvider, out var sttEndpoint);

            string transcript;
            try
            {
                transcript = await sttService.TranscribeAsync(
                    new SttRequest(audio, settings.SttProvider, sttEndpoint, settings.SttModel, effectiveInputLanguage, sttKey),
                    timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                trayStatusService.SetStatus(TrayStatus.Error, "Transkription fehlgeschlagen. Bitte API-Schlüssel, Modell und Netzwerk prüfen.");
                return Fail("Transkription fehlgeschlagen.", ex);
            }

            if (string.IsNullOrWhiteSpace(transcript))
            {
                trayStatusService.SetStatus(TrayStatus.Idle, "Kein Text erkannt. Es wurde nichts eingefügt.");
                return Fail("Transkription lieferte keinen Text (leeres Ergebnis).");
            }

            var llmKey = secretProtector.Unprotect(settings.LlmApiKeyEncrypted!);
            Defaults.TryGetLlmEndpoint(settings.LlmProvider, out var llmEndpoint);
            string finalText;
            try
            {
                var baseSystemPrompt = PromptComposition.EffectiveBaseSystemPrompt(profile, assistant);
                var policyBlock = PromptComposition.BuildPolicyBlock(profile, assistant, effectiveInputLanguage, effectiveOutputLanguage);
                var systemPrompt = PromptComposition.BuildSystemPrompt(baseSystemPrompt, effectiveInputLanguage, policyBlock);
                var modePrompt = BuildModePrompt(assistant, sourceText);
                finalText = await llmService.ProcessAsync(
                    new LlmRequest(transcript, assistant.Type, settings.LlmProvider, llmEndpoint, settings.LlmModel, llmKey, systemPrompt, modePrompt, sourceText),
                    timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                failureLog.Record("KI-Verarbeitung (Chat) fehlgeschlagen; es wird das Transkript eingefügt.", ex);
                finalText = transcript;
                trayStatusService.SetStatus(TrayStatus.Error, "KI-Verarbeitung fehlgeschlagen. Das Transkript wird eingefügt.");
            }

            if (string.IsNullOrWhiteSpace(finalText))
            {
                trayStatusService.SetStatus(TrayStatus.Idle, "Kein Text erkannt. Es wurde nichts eingefügt.");
                return Fail("Kein finaler Text nach Transkription/KI.", transcript: transcript);
            }

            var insertUsedClipboardFallback = false;
            try
            {
                insertUsedClipboardFallback = await InsertFinalTextAsync(
                    finalText,
                    settings.InsertMethod,
                    settings.RestoreClipboard,
                    timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                trayStatusService.SetStatus(TrayStatus.Error, "Einfügen fehlgeschlagen. Details siehe Diagnose.");
                return Fail($"Einfügen fehlgeschlagen ({settings.InsertMethod}).", ex, transcript: transcript, finalText: finalText);
            }

            failureLog.Clear();
            if (insertUsedClipboardFallback)
            {
                failureLog.Record(
                    "Text eingefügt. Die Zwischenablage ließ sich nicht nutzen; es wurde automatisch per Tastatur eingegeben (Fallback).");
            }

            trayStatusService.SetStatus(
                TrayStatus.Success,
                insertUsedClipboardFallback
                    ? "Text eingefügt (Fallback: direktes Tippen, da die Zwischenablage blockiert war)."
                    : "Text eingefügt.");
            _ = ResetTransientStatusAsync(settingsService, hotkeyService, trayStatusService, TimeSpan.FromSeconds(2), cancellationToken);
            return PipelineResult.Ok("Text eingefügt.", transcript, finalText);
        }
        catch (OperationCanceledException ex)
        {
            trayStatusService.SetStatus(TrayStatus.Error, "Zeitüberschreitung. Es wurde nichts eingefügt.");
            return Fail("Zeitüberschreitung bei der Verarbeitung.", ex);
        }
        catch (Exception ex)
        {
            trayStatusService.SetStatus(TrayStatus.Error, "Unerwarteter Fehler. Bitte Diagnose prüfen.");
            return Fail("Unerwarteter Pipeline-Fehler.", ex);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>
    /// Liefert <see langword="true"/>, wenn nach einem Zwischenablage-Fehler per <see cref="InsertMethod.SendInput"/> eingefügt wurde.
    /// Bei <see cref="InsertMethod.Clipboard"/> und einem Fehler (z. B. gesperrte Zwischenablage) einmaliger Fallback auf SendInput.
    /// Ob die Zielanwendung Einfügen wirklich übernommen hat, lässt sich ohne UI-Prüfung nicht erkennen.
    /// </summary>
    private async Task<bool> InsertFinalTextAsync(
        string finalText,
        InsertMethod insertMethod,
        bool restoreClipboard,
        CancellationToken cancellationToken)
    {
        if (insertMethod != InsertMethod.Clipboard)
        {
            await inputInjector.InsertTextAsync(finalText, insertMethod, restoreClipboard, cancellationToken).ConfigureAwait(false);
            return false;
        }

        try
        {
            await inputInjector.InsertTextAsync(finalText, InsertMethod.Clipboard, restoreClipboard, cancellationToken).ConfigureAwait(false);
            return false;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception)
        {
            trayStatusService.SetStatus(TrayStatus.Processing, "Zwischenablage nicht nutzbar – Text wird eingegeben …");
            await inputInjector.InsertTextAsync(finalText, InsertMethod.SendInput, restoreClipboard: false, cancellationToken).ConfigureAwait(false);
            return true;
        }
    }

    private string BuildModePrompt(AssistantInstance assistant, string? sourceText)
    {
        var prompt = _promptService.GetPrompt(assistant);
        var sourceBlock = !string.IsNullOrWhiteSpace(sourceText)
            ? $"Quelltext (Bezug für die Antwort, nicht in die Ausgabe übernehmen):{Environment.NewLine}\"\"\"{Environment.NewLine}{sourceText.Trim()}{Environment.NewLine}\"\"\""
            : string.Empty;
        return string.Join(
            Environment.NewLine + Environment.NewLine,
            new[] { prompt, sourceBlock }.Where(part => !string.IsNullOrWhiteSpace(part)));
    }

    private static async Task ResetTransientStatusAsync(
        ISettingsService settingsService,
        IHotkeyService hotkeyService,
        ITrayStatusService trayStatusService,
        TimeSpan delay,
        CancellationToken cancellationToken)
    {
        try
        {
            await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
            if (trayStatusService.CurrentStatus != TrayStatus.Success)
            {
                return;
            }

            var settings = await settingsService.LoadAsync(cancellationToken).ConfigureAwait(false);
            var readiness = settingsService.Validate(settings);
            if (hotkeyService.IsPaused)
            {
                trayStatusService.SetStatus(TrayStatus.Paused, "Inaktiv. Aktiviere die App im Tray-Menü, um Tastenkürzel zu nutzen.");
            }
            else if (readiness.IsReady)
            {
                trayStatusService.SetStatus(TrayStatus.Idle, "Bereit. Halte ein Tastenkürzel gedrückt, um zu diktieren.");
            }
            else
            {
                trayStatusService.SetStatus(TrayStatus.ConfigurationRequired, "Einrichtung erforderlich. Bitte prüfe die markierten Einstellungen.");
            }
        }
        catch (OperationCanceledException)
        {
        }
    }
}

public sealed class InMemoryTrayStatusService : ITrayStatusService
{
    public TrayStatus CurrentStatus { get; private set; } = TrayStatus.Idle;
    public string Message { get; private set; } = "Bereit. Halte ein Tastenkürzel gedrückt, um zu diktieren.";
    public event EventHandler<TrayStatusChangedEventArgs>? StatusChanged;

    public void SetStatus(TrayStatus status, string message)
    {
        CurrentStatus = status;
        Message = message;
        StatusChanged?.Invoke(this, new TrayStatusChangedEventArgs(status, message));
    }
}
