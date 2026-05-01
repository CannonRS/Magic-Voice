using MagicVoice.Core;
using MagicVoice.Infrastructure;

namespace MagicVoice.Infrastructure.Tests;

public class InfrastructureBehaviorTests
{
    private static IAppProfile CreateProfile() => new TestProfile();

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
        public string SystemPrompt => "Test.";
        public IReadOnlyList<AssistantModeDefinition> Modes { get; } =
        [
            new(AssistantMode.Transform, "Korrektur", "Test", "Ctrl+Shift+1", "Test.")
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

    [Fact]
    public void DpapiSecretProtector_roundtrips_secret_for_current_user()
    {
        var protector = new DpapiSecretProtector();

        var encrypted = protector.Protect("geheim");
        var plain = protector.Unprotect(encrypted);

        Assert.NotEqual("geheim", encrypted);
        Assert.Equal("geheim", plain);
    }

    [Fact]
    public void WavWriter_creates_riff_wave_payload()
    {
        var bytes = WavWriter.ToWav(new AudioBuffer([1, 0, 2, 0], 16000, 1, TimeSpan.FromMilliseconds(1)));

        Assert.Equal((byte)'R', bytes[0]);
        Assert.Equal((byte)'I', bytes[1]);
        Assert.Equal((byte)'F', bytes[2]);
        Assert.Equal((byte)'F', bytes[3]);
        Assert.Contains((byte)'W', bytes);
    }

    [Fact]
    public async Task SettingsService_creates_defaults_and_reports_missing_configuration()
    {
        var temp = Path.Combine(Path.GetTempPath(), $"magic-s2t-tests-{Guid.NewGuid():N}");
        var service = new SettingsService(CreateProfile(), temp);
        var settings = await service.LoadAsync();
        var readiness = service.Validate(settings);

        Assert.Contains(readiness.Issues, issue => issue.Field == "llmApiKey");
        Assert.Contains(readiness.Issues, issue => issue.Field == "sttApiKey" || issue.Field == "llmApiKey");
        Directory.Delete(temp, recursive: true);
    }

    [Fact]
    public void SettingsService_requires_supported_providers()
    {
        var service = new SettingsService(CreateProfile(), Path.Combine(Path.GetTempPath(), $"magic-s2t-tests-{Guid.NewGuid():N}"));
        var settings = new AppSettings
        {
            LlmApiKeyEncrypted = "secret",
            SttApiKeyEncrypted = "secret",
            SttProvider = "Unbekanntes STT",
            LlmProvider = "Unbekannte KI"
        };

        var readiness = service.Validate(settings);

        Assert.Contains(readiness.Issues, issue => issue.Field == "sttProvider");
        Assert.Contains(readiness.Issues, issue => issue.Field == "llmProvider");
    }

    [Fact]
    public async Task SettingsService_persists_window_bounds()
    {
        var temp = Path.Combine(Path.GetTempPath(), $"magic-s2t-tests-{Guid.NewGuid():N}");
        var service = new SettingsService(CreateProfile(), temp);
        var settings = new AppSettings();
        settings.WindowBounds.X = 100;
        settings.WindowBounds.Y = 120;
        settings.WindowBounds.Width = 1024;
        settings.WindowBounds.Height = 720;

        await service.SaveAsync(settings);
        var loaded = await service.LoadAsync();

        Assert.True(loaded.WindowBounds.IsSet);
        Assert.Equal(1024, loaded.WindowBounds.Width);
        Directory.Delete(temp, recursive: true);
    }
}
