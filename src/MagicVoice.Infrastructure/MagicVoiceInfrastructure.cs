using System.Diagnostics;
using System.Globalization;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using NAudio.Wave;
using MagicVoice.Core;

namespace MagicVoice.Infrastructure;

public sealed class DpapiSecretProtector : ISecretProtector
{
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("Magic-Voice.v1");

    public string Protect(string secret)
    {
        if (string.IsNullOrEmpty(secret))
        {
            return string.Empty;
        }

        var bytes = Encoding.UTF8.GetBytes(secret);
        var protectedBytes = ProtectedData.Protect(bytes, Entropy, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(protectedBytes);
    }

    public string Unprotect(string protectedSecret)
    {
        if (string.IsNullOrEmpty(protectedSecret))
        {
            return string.Empty;
        }

        var bytes = Convert.FromBase64String(protectedSecret);
        var plainBytes = ProtectedData.Unprotect(bytes, Entropy, DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(plainBytes);
    }
}

public sealed class SettingsService : ISettingsService
{
    private readonly IAppProfile _profile;
    private readonly JsonSerializerOptions _jsonOptions = new(JsonSerializerDefaults.Web)
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    public string DataDirectory { get; }
    public string LogDirectory { get; }
    public string SettingsPath { get; }

    public SettingsService(IAppProfile profile, string? dataDirectory = null)
    {
        _profile = profile;
        DataDirectory = dataDirectory ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), profile.DataFolderName);
        LogDirectory = Path.Combine(DataDirectory, "logs");
        SettingsPath = Path.Combine(DataDirectory, "settings.json");
    }

    public async Task<AppSettings> LoadAsync(CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(DataDirectory);
        Directory.CreateDirectory(LogDirectory);

        if (!File.Exists(SettingsPath))
        {
            var defaults = new AppSettings { Assistants = _profile.CreateDefaultAssistants() };
            await SaveAsync(defaults, cancellationToken).ConfigureAwait(false);
            return defaults;
        }

        try
        {
            await using var stream = File.OpenRead(SettingsPath);
            var settings = await JsonSerializer.DeserializeAsync<AppSettings>(stream, _jsonOptions, cancellationToken).ConfigureAwait(false) ?? new AppSettings();
            Normalize(settings);
            return settings;
        }
        catch (JsonException)
        {
            var backupPath = $"{SettingsPath}.defekt-{DateTimeOffset.Now:yyyyMMddHHmmss}.bak";
            File.Move(SettingsPath, backupPath, overwrite: true);
            var defaults = new AppSettings { Assistants = _profile.CreateDefaultAssistants() };
            await SaveAsync(defaults, cancellationToken).ConfigureAwait(false);
            return defaults;
        }
    }

    public async Task SaveAsync(AppSettings settings, CancellationToken cancellationToken = default)
    {
        Directory.CreateDirectory(DataDirectory);
        Directory.CreateDirectory(LogDirectory);
        Normalize(settings);

        var tempPath = $"{SettingsPath}.tmp";
        await using (var stream = File.Create(tempPath))
        {
            await JsonSerializer.SerializeAsync(stream, settings, _jsonOptions, cancellationToken).ConfigureAwait(false);
        }

        File.Move(tempPath, SettingsPath, overwrite: true);
    }

    public ReadinessReport Validate(AppSettings settings)
    {
        var issues = new List<ValidationIssue>();

        if (string.IsNullOrWhiteSpace(settings.SttProvider))
        {
            issues.Add(new("sttProvider", "Bitte einen Transkriptionsanbieter auswählen."));
        }
        else if (!Defaults.TryGetSttEndpoint(settings.SttProvider, out _))
        {
            issues.Add(new("sttProvider", "Der gewählte Transkriptionsanbieter wird noch nicht unterstützt."));
        }

        if (string.IsNullOrWhiteSpace(settings.SttModel))
        {
            issues.Add(new("sttModel", "Bitte ein Transkriptionsmodell wählen."));
        }

        if (!settings.HasEncryptedSttApiKey)
        {
            issues.Add(new("sttApiKey", "Bitte einen API-Schlüssel für die Transkription speichern."));
        }

        if (string.IsNullOrWhiteSpace(settings.LlmProvider))
        {
            issues.Add(new("llmProvider", "Bitte einen KI-Anbieter auswählen."));
        }
        else if (!Defaults.TryGetLlmEndpoint(settings.LlmProvider, out _))
        {
            issues.Add(new("llmProvider", "Der gewählte KI-Anbieter wird noch nicht unterstützt."));
        }

        if (string.IsNullOrWhiteSpace(settings.LlmModel))
        {
            issues.Add(new("llmModel", "Bitte ein KI-Modell wählen."));
        }

        if (!settings.HasEncryptedLlmApiKey)
        {
            issues.Add(new("llmApiKey", "Bitte einen API-Schlüssel für die KI-Verarbeitung speichern."));
        }

        if (settings.Assistants.Count == 0)
        {
            issues.Add(new("assistants", "Es ist kein Assistent konfiguriert. Bitte mindestens einen anlegen."));
        }
        else
        {
            foreach (var assistant in settings.Assistants)
            {
                var label = string.IsNullOrWhiteSpace(assistant.Name) ? assistant.Type.ToString() : assistant.Name;
                if (!HotkeyParser.TryParse(assistant.Hotkey, out _, out var hotkeyError))
                {
                    issues.Add(new($"hotkey.{assistant.Id}", $"{label}: {hotkeyError}"));
                }
            }
        }

        return new ReadinessReport(issues.Count == 0 ? TrayStatus.Idle : TrayStatus.ConfigurationRequired, issues);
    }

    private void Normalize(AppSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.InputLanguage))
        {
            settings.InputLanguage = "de";
        }

        if (string.IsNullOrWhiteSpace(settings.OutputLanguage))
        {
            settings.OutputLanguage = Defaults.SameAsInputLanguageCode;
        }

        if (string.IsNullOrWhiteSpace(settings.AudioInputDeviceId))
        {
            settings.AudioInputDeviceId = Defaults.DefaultAudioInputDeviceId;
        }

        if (string.IsNullOrWhiteSpace(settings.LastSelectedSettingsSection))
        {
            settings.LastSelectedSettingsSection = "overview";
        }

        // Eine leere Liste (`[]`) ist faktisch wie "kein Assistent konfiguriert" — die App
        // ist dann ohne Wirkung und der User würde im UI eine leere Hotkey-Seite sehen.
        // ??= würde nur null ersetzen, nicht eine leere Liste, daher hier explizit prüfen.
        if (settings.Assistants is null || settings.Assistants.Count == 0)
        {
            settings.Assistants = _profile.CreateDefaultAssistants();
        }
        foreach (var assistant in settings.Assistants)
        {
            if (string.IsNullOrWhiteSpace(assistant.Id))
            {
                assistant.Id = Guid.NewGuid().ToString("N");
            }

            var typeDefinition = _profile.Modes.FirstOrDefault(m => m.Mode == assistant.Type);
            if (typeDefinition is null)
            {
                assistant.Type = AssistantMode.Transform;
                typeDefinition = _profile.Modes.FirstOrDefault(m => m.Mode == assistant.Type);
                if (typeDefinition is null)
                {
                    continue;
                }
            }

            if (!Enum.IsDefined(assistant.ParagraphDensity))
            {
                assistant.ParagraphDensity = ParagraphDensity.Balanced;
            }

            if (!Enum.IsDefined(assistant.EmojiExpression))
            {
                assistant.EmojiExpression = EmojiExpression.Balanced;
            }

            if (string.IsNullOrWhiteSpace(assistant.Name))
            {
                assistant.Name = typeDefinition.Name;
            }

            if (string.IsNullOrWhiteSpace(assistant.Prompt))
            {
                assistant.Prompt = typeDefinition.DefaultPrompt;
            }

            assistant.SystemPromptOverride = string.IsNullOrWhiteSpace(assistant.SystemPromptOverride)
                ? null
                : assistant.SystemPromptOverride.Trim();

            assistant.InputLanguageOverride = NormalizeLanguageOverride(
                assistant.InputLanguageOverride,
                Defaults.InputLanguages.Select(l => l.Code));

            assistant.OutputLanguageOverride = NormalizeLanguageOverride(
                assistant.OutputLanguageOverride,
                Defaults.OutputLanguages.Select(l => l.Code));

            assistant.Intensity = Math.Clamp(assistant.Intensity, Defaults.MinModeIntensity, Defaults.MaxModeIntensity);
        }
    }

    private static string? NormalizeLanguageOverride(string? value, IEnumerable<string> validCodes)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        value = value.Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return validCodes.Any(code => code.Equals(value, StringComparison.OrdinalIgnoreCase)) ? value : null;
    }

}

public sealed class NAudioRecorder(ISettingsService settingsService, ILogger<NAudioRecorder>? logger = null) : IAudioRecorder
{
    private WaveInEvent? _waveIn;
    private MemoryStream? _buffer;
    private DateTimeOffset _startedAt;
    private readonly WaveFormat _format = new(16000, 16, 1);

    public async Task StartAsync(CancellationToken cancellationToken = default)
    {
        if (_waveIn is not null)
        {
            return;
        }

        var settings = await settingsService.LoadAsync(cancellationToken).ConfigureAwait(false);
        _buffer = new MemoryStream();
        _startedAt = DateTimeOffset.UtcNow;
        _waveIn = new WaveInEvent
        {
            DeviceNumber = ResolveDeviceNumber(settings.AudioInputDeviceId),
            WaveFormat = _format,
            BufferMilliseconds = 50,
            NumberOfBuffers = 3
        };
        _waveIn.DataAvailable += (_, args) => _buffer?.Write(args.Buffer, 0, args.BytesRecorded);
        _waveIn.RecordingStopped += (_, args) =>
        {
            if (args.Exception is not null)
            {
                logger?.LogError(args.Exception, "Audioaufnahme wurde mit Fehler beendet.");
            }
        };
        try
        {
            _waveIn.StartRecording();
        }
        catch (Exception ex) when (IsMicrophoneStartupException(ex))
        {
            _waveIn.Dispose();
            _waveIn = null;
            _buffer.Dispose();
            _buffer = null;
            throw new InvalidOperationException("Mikrofonzugriff fehlgeschlagen. Bitte prüfe Mikrofon, Windows-Datenschutzeinstellungen und ob ein anderes Programm das Gerät blockiert.", ex);
        }

    }

    private static int ResolveDeviceNumber(string deviceId)
    {
        if (string.Equals(deviceId, Defaults.DefaultAudioInputDeviceId, StringComparison.OrdinalIgnoreCase))
        {
            return -1;
        }

        return int.TryParse(deviceId, NumberStyles.Integer, CultureInfo.InvariantCulture, out var deviceNumber)
            ? deviceNumber
            : -1;
    }

    private static bool IsMicrophoneStartupException(Exception ex) =>
        ex is InvalidOperationException or UnauthorizedAccessException
        || ex.GetType().Name.Contains("MmException", StringComparison.OrdinalIgnoreCase);

    public Task<AudioBuffer> StopAsync(CancellationToken cancellationToken = default)
    {
        if (_waveIn is null || _buffer is null)
        {
            return Task.FromResult(new AudioBuffer([], 16000, 1, TimeSpan.Zero));
        }

        _waveIn.StopRecording();
        _waveIn.Dispose();
        _waveIn = null;

        var bytes = _buffer.ToArray();
        _buffer.Dispose();
        _buffer = null;
        return Task.FromResult(new AudioBuffer(bytes, 16000, 1, DateTimeOffset.UtcNow - _startedAt));
    }

    public Task AbortAsync(CancellationToken cancellationToken = default)
    {
        _waveIn?.StopRecording();
        _waveIn?.Dispose();
        _waveIn = null;
        _buffer?.Dispose();
        _buffer = null;
        return Task.CompletedTask;
    }
}

public sealed class NAudioDeviceService : IAudioDeviceService
{
    public IReadOnlyList<AudioInputDevice> GetInputDevices()
    {
        var devices = new List<AudioInputDevice>
        {
            new(Defaults.DefaultAudioInputDeviceId, "Systemstandard", true)
        };

        for (var index = 0; index < WaveIn.DeviceCount; index++)
        {
            var capabilities = WaveIn.GetCapabilities(index);
            devices.Add(new(index.ToString(CultureInfo.InvariantCulture), capabilities.ProductName, false));
        }

        return devices;
    }
}

public sealed class OpenAiCompatibleSttService(HttpClient httpClient) : ISttService
{
    public async Task<string> TranscribeAsync(SttRequest request, CancellationToken cancellationToken = default)
    {
        using var form = new MultipartFormDataContent();
        form.Add(new StringContent(request.Model), "model");
        if (!Defaults.IsAutoLanguage(request.Language))
        {
            form.Add(new StringContent(request.Language), "language");
        }

        form.Add(new StringContent("json"), "response_format");

        var wavBytes = WavWriter.ToWav(request.Audio);
        var audioContent = new ByteArrayContent(wavBytes);
        audioContent.Headers.ContentType = MediaTypeHeaderValue.Parse("audio/wav");
        form.Add(audioContent, "file", "aufnahme.wav");

        using var message = new HttpRequestMessage(HttpMethod.Post, request.Endpoint);
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", request.ApiKey);
        message.Content = form;

        using var response = await httpClient.SendAsync(message, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        var payload = await JsonSerializer.DeserializeAsync<SttResponse>(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return payload?.Text?.Trim() ?? string.Empty;
    }

    private sealed record SttResponse([property: JsonPropertyName("text")] string? Text);
}

public sealed class OpenAiCompatibleLlmService(HttpClient httpClient) : ILlmService
{
    public async Task<string> ProcessAsync(LlmRequest request, CancellationToken cancellationToken = default)
    {
        var body = new
        {
            model = request.Model,
            temperature = 0.2,
            messages = new object[]
            {
                new { role = "system", content = request.SystemPrompt },
                new { role = "user", content = $"{request.ModePrompt}\n\nTranskript:\n{request.Transcript}" }
            }
        };

        using var message = new HttpRequestMessage(HttpMethod.Post, request.Endpoint);
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", request.ApiKey);
        message.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        using var response = await httpClient.SendAsync(message, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        var payload = await JsonSerializer.DeserializeAsync<ChatResponse>(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return payload?.Choices?.FirstOrDefault()?.Message?.Content?.Trim() ?? string.Empty;
    }

    private sealed record ChatResponse([property: JsonPropertyName("choices")] Choice[]? Choices);
    private sealed record Choice([property: JsonPropertyName("message")] ChatMessage? Message);
    private sealed record ChatMessage([property: JsonPropertyName("content")] string? Content);
}

public sealed class ClipboardInputInjector : IInputInjector
{
    public async Task InsertTextAsync(string text, InsertMethod method, bool restoreClipboard, CancellationToken cancellationToken = default)
    {
        if (method == InsertMethod.SendInput)
        {
            NativeInput.SendUnicodeText(text);
            return;
        }

        var previousText = restoreClipboard ? NativeClipboard.TryGetText() : null;

        NativeClipboard.SetText(text);
        NativeInput.SendPasteShortcut();
        await Task.Delay(150, cancellationToken).ConfigureAwait(true);

        if (restoreClipboard && previousText is not null)
        {
            try
            {
                NativeClipboard.SetText(previousText);
            }
            catch
            {
                // Die Einfügung war erfolgreich; ein Restore-Fehler soll die Pipeline nicht nachträglich scheitern lassen.
            }
        }
    }
}

internal static class NativeClipboard
{
    private const uint CfUnicodeText = 13;
    private const uint GmemMoveable = 0x0002;

    public static string? TryGetText()
    {
        if (!OpenClipboard(IntPtr.Zero))
        {
            return null;
        }

        try
        {
            var handle = GetClipboardData(CfUnicodeText);
            if (handle == IntPtr.Zero)
            {
                return null;
            }

            var pointer = GlobalLock(handle);
            if (pointer == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                return Marshal.PtrToStringUni(pointer);
            }
            finally
            {
                GlobalUnlock(handle);
            }
        }
        finally
        {
            CloseClipboard();
        }
    }

    public static void SetText(string text)
    {
        if (!OpenClipboard(IntPtr.Zero))
        {
            throw new InvalidOperationException("Zwischenablage konnte nicht geöffnet werden.");
        }

        try
        {
            EmptyClipboard();
            var bytes = (text.Length + 1) * 2;
            var handle = GlobalAlloc(GmemMoveable, (UIntPtr)bytes);
            if (handle == IntPtr.Zero)
            {
                throw new InvalidOperationException("Zwischenablage-Speicher konnte nicht reserviert werden.");
            }

            var pointer = GlobalLock(handle);
            if (pointer == IntPtr.Zero)
            {
                throw new InvalidOperationException("Zwischenablage-Speicher konnte nicht gesperrt werden.");
            }

            try
            {
                var data = Encoding.Unicode.GetBytes(text + '\0');
                Marshal.Copy(data, 0, pointer, data.Length);
            }
            finally
            {
                GlobalUnlock(handle);
            }

            if (SetClipboardData(CfUnicodeText, handle) == IntPtr.Zero)
            {
                throw new InvalidOperationException("Zwischenablage konnte nicht gesetzt werden.");
            }
        }
        finally
        {
            CloseClipboard();
        }
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr hWndNewOwner);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool CloseClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool EmptyClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr GetClipboardData(uint uFormat);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GlobalUnlock(IntPtr hMem);
}

public sealed class WindowsClipboardSourceCapture : IClipboardSourceCapture
{
    public string? TryGetText() => NativeClipboard.TryGetText();
}

internal static class NativeInput
{
    private const ushort VkControl = 0x11;
    private const ushort VkV = 0x56;
    private const ushort VkReturn = 0x0D;
    private const uint KeyEventFKeyUp = 0x0002;
    private const uint KeyEventFUnicode = 0x0004;
    private const int InputKeyboard = 1;

    public static void SendPasteShortcut()
    {
        var inputs = new[]
        {
            KeyboardInput(VkControl, 0),
            KeyboardInput(VkV, 0),
            KeyboardInput(VkV, KeyEventFKeyUp),
            KeyboardInput(VkControl, KeyEventFKeyUp)
        };
        Send(inputs);
    }

    public static void SendUnicodeText(string text)
    {
        // Zeilenumbrüche normalisieren und als echten Enter-Tastendruck schicken.
        // Sonst landen \n und \r als Unicode-Codepoints im Text – Word stellt die als Kästchen dar,
        // weil es nur \r (Absatz) bzw. VK_RETURN als Zeilentrenner akzeptiert.
        var normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
        var inputs = new List<Input>(normalized.Length * 2);
        foreach (var character in normalized)
        {
            if (character == '\n')
            {
                inputs.Add(KeyboardInput(VkReturn, 0));
                inputs.Add(KeyboardInput(VkReturn, KeyEventFKeyUp));
                continue;
            }

            inputs.Add(UnicodeInput(character, 0));
            inputs.Add(UnicodeInput(character, KeyEventFKeyUp));
        }

        Send(inputs.ToArray());
    }

    private static void Send(Input[] inputs)
    {
        if (inputs.Length == 0)
        {
            return;
        }

        var sent = SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<Input>());
        if (sent != inputs.Length)
        {
            var error = Marshal.GetLastWin32Error();
            throw new InvalidOperationException(
                $"Tastatureingabe konnte nicht gesendet werden ({sent}/{inputs.Length} Ereignisse, Win32-Fehler {error}). " +
                "Mögliche Ursachen: Die Zielanwendung läuft mit höheren Rechten (UIPI blockiert Eingaben), " +
                "die Eingabe wurde durch das Betriebssystem blockiert oder das Zielfenster akzeptiert kein SendInput. " +
                "Tipp: Wechsle in den Einstellungen unter „Einfügen“ die Methode auf „über die Zwischenablage einfügen“.");
        }
    }

    private static Input KeyboardInput(ushort key, uint flags) => new()
    {
        Type = InputKeyboard,
        Data = new InputUnion { Keyboard = new KeyboardInputData { VirtualKey = key, Flags = flags } }
    };

    private static Input UnicodeInput(char character, uint flags) => new()
    {
        Type = InputKeyboard,
        Data = new InputUnion { Keyboard = new KeyboardInputData { ScanCode = character, Flags = flags | KeyEventFUnicode } }
    };

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint cInputs, Input[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Sequential)]
    private struct Input
    {
        public int Type;
        public InputUnion Data;
    }

    // Die Win32-INPUT-Struktur ist eine Union aus MOUSEINPUT, KEYBDINPUT und HARDWAREINPUT.
    // Damit Marshal.SizeOf<Input>() die korrekte Größe (40 Bytes auf x64) liefert,
    // muss die Union alle drei Member enthalten – sonst lehnt SendInput mit Win32-Fehler 87
    // (ERROR_INVALID_PARAMETER) ab.
    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MouseInputData Mouse;
        [FieldOffset(0)]
        public KeyboardInputData Keyboard;
        [FieldOffset(0)]
        public HardwareInputData Hardware;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardInputData
    {
        public ushort VirtualKey;
        public ushort ScanCode;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInputData
    {
        public int Dx;
        public int Dy;
        public uint MouseData;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct HardwareInputData
    {
        public uint Msg;
        public ushort ParamL;
        public ushort ParamH;
    }
}

/// <summary>
/// Spielt zwei kurze Sinus-Beeps (mit kurzer Hüllkurve, damit es nicht klickt) für
/// Aufnahme-Start und Aufnahme-Ende. Jeder Aufruf erzeugt eine eigene WaveOutEvent-Instanz und
/// kann sich überlappen; Audio-Fehler werden geschluckt, damit die Aufnahme dadurch nicht stoppt.
/// </summary>
public sealed class NAudioFeedbackSoundService(ISettingsService settingsService) : IFeedbackSoundService
{
    public async void PlayRecordingStart()
    {
        if (await ShouldPlayAsync().ConfigureAwait(false))
        {
            Play(880, 80);
        }
    }

    public async void PlayRecordingStop()
    {
        if (await ShouldPlayAsync().ConfigureAwait(false))
        {
            Play(587, 110);
        }
    }

    private async Task<bool> ShouldPlayAsync()
    {
        try
        {
            var settings = await settingsService.LoadAsync().ConfigureAwait(false);
            return settings.PlayRecordingSounds;
        }
        catch
        {
            return false;
        }
    }

    private static void Play(double frequencyHz, int durationMs)
    {
        try
        {
            var provider = BuildBeep(frequencyHz, durationMs, gain: 0.18);
            var output = new WaveOutEvent();
            output.Init(provider);
            output.PlaybackStopped += (_, _) => output.Dispose();
            output.Play();
        }
        catch
        {
            // Kein Audio-Ausgabegerät verfügbar oder exklusiv blockiert – darf die Aufnahme nicht behindern.
        }
    }

    private static NAudio.Wave.ISampleProvider BuildBeep(double frequencyHz, int durationMs, double gain)
    {
        const int sampleRate = 44100;
        var totalSamples = sampleRate * durationMs / 1000;
        var fadeSamples = sampleRate * 8 / 1000;
        var samples = new float[totalSamples];
        for (var i = 0; i < totalSamples; i++)
        {
            var sine = Math.Sin(2 * Math.PI * frequencyHz * i / sampleRate);
            var envelope = 1.0;
            if (i < fadeSamples)
            {
                envelope = (double)i / fadeSamples;
            }
            else if (i > totalSamples - fadeSamples)
            {
                envelope = (double)(totalSamples - i) / fadeSamples;
            }

            samples[i] = (float)(sine * envelope * gain);
        }

        return new BufferedSampleProvider(samples, sampleRate);
    }
}

internal sealed class BufferedSampleProvider : NAudio.Wave.ISampleProvider
{
    private readonly float[] _buffer;
    private int _position;

    public NAudio.Wave.WaveFormat WaveFormat { get; }

    public BufferedSampleProvider(float[] buffer, int sampleRate)
    {
        _buffer = buffer;
        WaveFormat = NAudio.Wave.WaveFormat.CreateIeeeFloatWaveFormat(sampleRate, 1);
    }

    public int Read(float[] buffer, int offset, int count)
    {
        var available = Math.Min(count, _buffer.Length - _position);
        if (available <= 0)
        {
            return 0;
        }

        Array.Copy(_buffer, _position, buffer, offset, available);
        _position += available;
        return available;
    }
}

public sealed class WindowsAutostartService(IAppProfile profile) : IAutostartService
{
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

    public bool IsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: false);
        return key?.GetValue(profile.AutostartRegistryValueName) is string;
    }

    public void SetEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKey, writable: true) ?? Registry.CurrentUser.CreateSubKey(RunKey);
        if (!enabled)
        {
            key.DeleteValue(profile.AutostartRegistryValueName, throwOnMissingValue: false);
            return;
        }

        var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (!string.IsNullOrWhiteSpace(exePath))
        {
            key.SetValue(profile.AutostartRegistryValueName, $"\"{exePath}\"");
        }
    }
}

public sealed class LowLevelKeyboardHotkeyService : IHotkeyService
{
    private readonly Dictionary<string, HotkeyGesture> _gestures = [];
    private readonly HashSet<string> _pressedAssistants = [];
    private readonly LowLevelKeyboardProc _proc;
    private IntPtr _hookId;
    private bool _paused;

    public event EventHandler<HotkeyPressedEventArgs>? HotkeyDown;
    public event EventHandler<HotkeyPressedEventArgs>? HotkeyUp;
    public bool IsPaused => _paused;

    public LowLevelKeyboardHotkeyService()
    {
        _proc = HookCallback;
    }

    public Task<IReadOnlyList<ValidationIssue>> RegisterAsync(AppSettings settings, CancellationToken cancellationToken = default)
    {
        _gestures.Clear();
        var issues = new List<ValidationIssue>();
        foreach (var assistant in settings.Assistants)
        {
            var label = string.IsNullOrWhiteSpace(assistant.Name) ? assistant.Type.ToString() : assistant.Name;
            if (!HotkeyParser.TryParse(assistant.Hotkey, out var gesture, out var error))
            {
                issues.Add(new($"hotkey.{assistant.Id}", $"{label}: {error}"));
                continue;
            }

            _gestures[assistant.Id] = gesture;
        }

        if (_hookId == IntPtr.Zero)
        {
            _hookId = SetHook(_proc);
            if (_hookId == IntPtr.Zero)
            {
                issues.Add(new("hotkeys", "Globale Tastenkürzel konnten nicht registriert werden."));
            }
        }

        return Task.FromResult<IReadOnlyList<ValidationIssue>>(issues);
    }

    public void Pause()
    {
        _paused = true;
        _pressedAssistants.Clear();
    }

    public void Resume() => _paused = false;

    public ValueTask DisposeAsync()
    {
        if (_hookId != IntPtr.Zero)
        {
            UnhookWindowsHookEx(_hookId);
            _hookId = IntPtr.Zero;
        }

        return ValueTask.CompletedTask;
    }

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && !_paused)
        {
            var message = (int)wParam;
            var keyCode = Marshal.ReadInt32(lParam);
            var key = VirtualKeyToName(keyCode);
            var isDown = message is WmKeyDown or WmSysKeyDown;
            var isUp = message is WmKeyUp or WmSysKeyUp;

            // Snapshot, da HotkeyUp-Handler den Service während der Iteration verändern könnten.
            foreach (var (assistantId, gesture) in _gestures.ToArray())
            {
                if (isDown)
                {
                    if (IsDownMatch(gesture, key) && _pressedAssistants.Add(assistantId))
                    {
                        HotkeyDown?.Invoke(this, new HotkeyPressedEventArgs(assistantId));
                    }
                }
                else if (isUp && _pressedAssistants.Contains(assistantId) && IsReleaseRelevant(gesture, keyCode, key))
                {
                    _pressedAssistants.Remove(assistantId);
                    HotkeyUp?.Invoke(this, new HotkeyPressedEventArgs(assistantId));
                }
            }
        }

        return CallNextHookEx(_hookId, nCode, wParam, lParam);
    }

    private static bool IsDownMatch(HotkeyGesture gesture, string key)
    {
        return key.Equals(gesture.Key, StringComparison.OrdinalIgnoreCase)
            && IsPressed(VkControl) == gesture.Control
            && IsPressed(VkMenu) == gesture.Alt
            && IsPressed(VkShift) == gesture.Shift
            && (IsPressed(VkLWin) || IsPressed(VkRWin)) == gesture.Windows;
    }

    /// <summary>
    /// Beim Loslassen reicht es, dass entweder die Haupttaste oder eine der erforderlichen Zusatztasten
    /// der laufenden Geste losgelassen wird. So endet die Aufnahme auch dann zuverlässig, wenn der Nutzer
    /// die Tasten in beliebiger Reihenfolge oder nahezu gleichzeitig loslässt.
    /// </summary>
    private static bool IsReleaseRelevant(HotkeyGesture gesture, int virtualKey, string keyName)
    {
        if (keyName.Equals(gesture.Key, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return virtualKey switch
        {
            VkControl or VkLControl or VkRControl => gesture.Control,
            VkMenu or VkLMenu or VkRMenu => gesture.Alt,
            VkShift or VkLShift or VkRShift => gesture.Shift,
            VkLWin or VkRWin => gesture.Windows,
            _ => false
        };
    }

    private static string VirtualKeyToName(int virtualKey)
    {
        if (virtualKey is >= 0x30 and <= 0x39)
        {
            return ((char)virtualKey).ToString();
        }

        if (virtualKey is >= 0x41 and <= 0x5A)
        {
            return ((char)virtualKey).ToString();
        }

        if (virtualKey is >= 0x70 and <= 0x87)
        {
            return $"F{virtualKey - 0x6F}";
        }

        return virtualKey switch
        {
            0x20 => "Space",
            0x0D => "Enter",
            0x09 => "Tab",
            0x1B => "Esc",
            _ => virtualKey.ToString(CultureInfo.InvariantCulture)
        };
    }

    private static bool IsPressed(int key) => (GetKeyState(key) & 0x8000) != 0;

    private static IntPtr SetHook(LowLevelKeyboardProc proc)
    {
        using var process = Process.GetCurrentProcess();
        using var module = process.MainModule!;
        return SetWindowsHookEx(WhKeyboardLl, proc, GetModuleHandle(module.ModuleName), 0);
    }

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private const int WhKeyboardLl = 13;
    private const int WmKeyDown = 0x0100;
    private const int WmKeyUp = 0x0101;
    private const int WmSysKeyDown = 0x0104;
    private const int WmSysKeyUp = 0x0105;
    private const int VkShift = 0x10;
    private const int VkControl = 0x11;
    private const int VkMenu = 0x12;
    private const int VkLWin = 0x5B;
    private const int VkRWin = 0x5C;
    private const int VkLShift = 0xA0;
    private const int VkRShift = 0xA1;
    private const int VkLControl = 0xA2;
    private const int VkRControl = 0xA3;
    private const int VkLMenu = 0xA4;
    private const int VkRMenu = 0xA5;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int nVirtKey);
}

public static class WavWriter
{
    public static byte[] ToWav(AudioBuffer audio)
    {
        using var output = new MemoryStream();
        using (var writer = new BinaryWriter(output, Encoding.ASCII, leaveOpen: true))
        {
            var byteRate = audio.SampleRate * audio.Channels * 2;
            writer.Write("RIFF"u8.ToArray());
            writer.Write(36 + audio.PcmBytes.Length);
            writer.Write("WAVE"u8.ToArray());
            writer.Write("fmt "u8.ToArray());
            writer.Write(16);
            writer.Write((short)1);
            writer.Write((short)audio.Channels);
            writer.Write(audio.SampleRate);
            writer.Write(byteRate);
            writer.Write((short)(audio.Channels * 2));
            writer.Write((short)16);
            writer.Write("data"u8.ToArray());
            writer.Write(audio.PcmBytes.Length);
            writer.Write(audio.PcmBytes);
        }

        return output.ToArray();
    }
}

public sealed class FileLogger
{
    private readonly string _logDirectory;
    public string CurrentLogPath => Path.Combine(_logDirectory, $"{DateTimeOffset.Now:yyyy-MM-dd}.log");

    public FileLogger(ISettingsService settingsService)
    {
        _logDirectory = settingsService.LogDirectory;
        Directory.CreateDirectory(_logDirectory);
    }

    public Task WriteAsync(string message, CancellationToken cancellationToken = default)
    {
        var line = $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}";
        return File.AppendAllTextAsync(CurrentLogPath, line, cancellationToken);
    }
}
