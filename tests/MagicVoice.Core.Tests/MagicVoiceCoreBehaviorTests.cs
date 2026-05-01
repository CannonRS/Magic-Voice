using MagicVoice.Core;

namespace MagicVoice.Core.Tests;

public class CoreBehaviorTests
{
    [Fact]
    public void HotkeyParser_accepts_modifier_plus_key()
    {
        var ok = HotkeyParser.TryParse("Ctrl+Shift+1", out var gesture, out var error);

        Assert.True(ok, error);
        Assert.True(gesture.Control);
        Assert.True(gesture.Shift);
        Assert.Equal("1", gesture.Key);
    }

    [Fact]
    public void HotkeyParser_rejects_key_without_modifier()
    {
        var ok = HotkeyParser.TryParse("A", out _, out var error);

        Assert.False(ok);
        Assert.Contains("Zusatztaste", error);
    }

    [Fact]
    public void Provider_endpoints_are_resolved_internally()
    {
        var sttOk = Defaults.TryGetSttEndpoint("OpenAI", out var sttEndpoint);
        var llmOk = Defaults.TryGetLlmEndpoint("OpenAI", out var llmEndpoint);

        Assert.True(sttOk);
        Assert.True(llmOk);
        Assert.Contains("/audio/transcriptions", sttEndpoint);
        Assert.Contains("/chat/completions", llmEndpoint);
    }

    [Fact]
    public async Task Pipeline_returns_configuration_required_when_settings_are_incomplete()
    {
        var status = new InMemoryTrayStatusService();
        var profile = new TestProfile();
        var pipeline = new SpeechPipeline(
            profile,
            new FakeRecorder(new AudioBuffer([1, 2, 3], 16000, 1, TimeSpan.FromSeconds(1))),
            new FakeStt("Hallo"),
            new FakeLlm("Hallo."),
            new FakeInjector(),
            new FakeSettingsService(new AppSettings()),
            new PlainSecretProtector(),
            new FakeHotkeys(),
            status,
            new NullProcessingFailureLog());

        var result = await pipeline.RunAsync("any-id", null);

        Assert.False(result.Success);
        Assert.Equal(TrayStatus.ConfigurationRequired, status.CurrentStatus);
    }

    [Fact]
    public async Task Pipeline_inserts_transcript_when_llm_fails()
    {
        var injector = new FakeInjector();
        var profile = new TestProfile();
        var settings = ValidSettings(profile);
        var pipeline = new SpeechPipeline(
            profile,
            new FakeRecorder(new AudioBuffer([1, 2, 3, 4], 16000, 1, TimeSpan.FromSeconds(1))),
            new FakeStt("Das ist ein Test"),
            new ThrowingLlm(),
            injector,
            new FakeSettingsService(settings),
            new PlainSecretProtector(),
            new FakeHotkeys(),
            new InMemoryTrayStatusService(),
            new NullProcessingFailureLog());

        var result = await pipeline.RunAsync(settings.Assistants[0].Id, null);

        Assert.True(result.Success);
        Assert.Equal("Das ist ein Test", injector.InsertedText);
    }

    [Fact]
    public async Task Pipeline_falls_back_to_send_input_when_clipboard_insert_throws()
    {
        var injector = new ClipboardFailsOnceInjector();
        var profile = new TestProfile();
        var settings = ValidSettings(profile);
        settings.InsertMethod = InsertMethod.Clipboard;
        var pipeline = new SpeechPipeline(
            profile,
            new FakeRecorder(new AudioBuffer([1, 2, 3, 4], 16000, 1, TimeSpan.FromSeconds(1))),
            new FakeStt("Hallo"),
            new FakeLlm("Fertig."),
            injector,
            new FakeSettingsService(settings),
            new PlainSecretProtector(),
            new FakeHotkeys(),
            new InMemoryTrayStatusService(),
            new NullProcessingFailureLog());

        var result = await pipeline.RunAsync(settings.Assistants[0].Id, null);

        Assert.True(result.Success);
        Assert.Equal(2, injector.CallCount);
        Assert.Equal(InsertMethod.Clipboard, injector.Methods[0]);
        Assert.Equal(InsertMethod.SendInput, injector.Methods[1]);
        Assert.Equal("Fertig.", injector.LastText);
    }

    [Fact]
    public async Task Pipeline_passes_clipboard_source_to_llm()
    {
        var capturingLlm = new CapturingLlm("Antwort");
        var profile = new TestProfile();
        var settings = ValidSettings(profile);
        var editAssistant = settings.Assistants.First(a => a.Type == AssistantMode.AnswerClipboard);
        var pipeline = new SpeechPipeline(
            profile,
            new FakeRecorder(new AudioBuffer([1, 2, 3, 4], 16000, 1, TimeSpan.FromSeconds(1))),
            new FakeStt("Bitte umformulieren"),
            capturingLlm,
            new FakeInjector(),
            new FakeSettingsService(settings),
            new PlainSecretProtector(),
            new FakeHotkeys(),
            new InMemoryTrayStatusService(),
            new NullProcessingFailureLog());

        var result = await pipeline.RunAsync(editAssistant.Id, "Originaltext aus der Zwischenablage");

        Assert.True(result.Success);
        Assert.Equal("Originaltext aus der Zwischenablage", capturingLlm.LastRequest?.SourceText);
        Assert.Contains("Originaltext aus der Zwischenablage", capturingLlm.LastRequest!.ModePrompt);
    }

    private static AppSettings ValidSettings(IAppProfile profile) => new()
    {
        LlmApiKeyEncrypted = "secret",
        SttApiKeyEncrypted = "secret",
        Assistants = profile.CreateDefaultAssistants()
    };

    private sealed class TestProfile : IAppProfile
    {
        public string AppName => "Test";
        public string DataFolderName => "Test";
        public string AuthorName => "Test";
        public string CopyrightText => "Test";
        public string LicenseName => "Test";
        public string MutexName => "Test";
        public string AumId => "Test.App";
        public string AutostartRegistryValueName => "Test";
        public string SystemPrompt => "Test-Systemprompt.";
        public IReadOnlyList<AssistantModeDefinition> Modes { get; } =
        [
            new(AssistantMode.Transform, "Text", "Test", "Ctrl+Shift+1", "Bearbeite."),
            new(AssistantMode.AnswerClipboard, "Bearbeiten", "Test", "Ctrl+Shift+5", "Bearbeite den Quelltext.", RequiresClipboardSource: true)
        ];
        public string IntensityStepName(AssistantMode mode, int intensity) => intensity.ToString();
        public string IntensityStepInstruction(AssistantMode mode, int intensity) => $"Stufe {intensity}";
        public string WritingStyleInstruction(WritingStyle style) => string.Empty;
        public string EmojiExpressionInstruction(EmojiExpression level) => string.Empty;
        public List<AssistantInstance> CreateDefaultAssistants() =>
            Modes.Select(m => new AssistantInstance
            {
                Id = Guid.NewGuid().ToString("N"),
                Type = m.Mode,
                Name = m.Name,
                Hotkey = m.DefaultHotkey,
                Prompt = m.DefaultPrompt,
                Intensity = Defaults.DefaultModeIntensity
            }).ToList();
    }

    private sealed class CapturingLlm(string text) : ILlmService
    {
        public LlmRequest? LastRequest { get; private set; }
        public Task<string> ProcessAsync(LlmRequest request, CancellationToken cancellationToken = default)
        {
            LastRequest = request;
            return Task.FromResult(text);
        }
    }

    private sealed class FakeRecorder(AudioBuffer audio) : IAudioRecorder
    {
        public Task AbortAsync(CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task StartAsync(CancellationToken cancellationToken = default) => Task.CompletedTask;
        public Task<AudioBuffer> StopAsync(CancellationToken cancellationToken = default) => Task.FromResult(audio);
    }

    private sealed class FakeStt(string transcript) : ISttService
    {
        public Task<string> TranscribeAsync(SttRequest request, CancellationToken cancellationToken = default) => Task.FromResult(transcript);
    }

    private sealed class FakeLlm(string text) : ILlmService
    {
        public Task<string> ProcessAsync(LlmRequest request, CancellationToken cancellationToken = default) => Task.FromResult(text);
    }

    private sealed class ThrowingLlm : ILlmService
    {
        public Task<string> ProcessAsync(LlmRequest request, CancellationToken cancellationToken = default) => throw new InvalidOperationException("LLM kaputt");
    }

    private sealed class FakeInjector : IInputInjector
    {
        public string? InsertedText { get; private set; }

        public Task InsertTextAsync(string text, InsertMethod method, bool restoreClipboard, CancellationToken cancellationToken = default)
        {
            InsertedText = text;
            return Task.CompletedTask;
        }
    }

    private sealed class ClipboardFailsOnceInjector : IInputInjector
    {
        public int CallCount { get; private set; }
        public string? LastText { get; private set; }
        public List<InsertMethod> Methods { get; } = [];

        public Task InsertTextAsync(string text, InsertMethod method, bool restoreClipboard, CancellationToken cancellationToken = default)
        {
            CallCount++;
            LastText = text;
            Methods.Add(method);
            if (method == InsertMethod.Clipboard)
            {
                throw new InvalidOperationException("Zwischenablage gesperrt");
            }

            return Task.CompletedTask;
        }
    }

    private sealed class FakeSettingsService(AppSettings settings) : ISettingsService
    {
        public string SettingsPath => "settings.json";
        public string DataDirectory => ".";
        public string LogDirectory => ".";
        public Task<AppSettings> LoadAsync(CancellationToken cancellationToken = default) => Task.FromResult(settings);
        public Task SaveAsync(AppSettings settings, CancellationToken cancellationToken = default) => Task.CompletedTask;

        public ReadinessReport Validate(AppSettings settings)
        {
            var issues = new List<ValidationIssue>();
            if (!settings.HasEncryptedLlmApiKey) issues.Add(new("llmApiKey", "Bitte einen API-Schlüssel für die KI-Verarbeitung speichern."));
            if (!settings.HasEncryptedSttApiKey) issues.Add(new("sttApiKey", "Bitte einen API-Schlüssel für die Transkription speichern."));
            return new ReadinessReport(issues.Count == 0 ? TrayStatus.Idle : TrayStatus.ConfigurationRequired, issues);
        }
    }

    private sealed class PlainSecretProtector : ISecretProtector
    {
        public string Protect(string secret) => secret;
        public string Unprotect(string protectedSecret) => protectedSecret;
    }

    private sealed class FakeHotkeys : IHotkeyService
    {
#pragma warning disable CS0067
        public event EventHandler<HotkeyPressedEventArgs>? HotkeyDown;
        public event EventHandler<HotkeyPressedEventArgs>? HotkeyUp;
#pragma warning restore CS0067
        public bool IsPaused { get; private set; }
        public Task<IReadOnlyList<ValidationIssue>> RegisterAsync(AppSettings settings, CancellationToken cancellationToken = default) => Task.FromResult<IReadOnlyList<ValidationIssue>>([]);
        public void Pause() => IsPaused = true;
        public void Resume() => IsPaused = false;
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
