using System.Text.Json;
using System.Text.Json.Serialization;

namespace MagicVoice.Core;

/// <summary>Liest gespeicherte Assistenten-Typen aus älteren Versionen und mappt sie auf die aktuellen Werte.</summary>
public sealed class AssistantModeJsonConverter : JsonConverter<AssistantMode>
{
    public override AssistantMode Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        return reader.TokenType switch
        {
            JsonTokenType.String => MapLegacy(reader.GetString()),
            JsonTokenType.Number => reader.TryGetInt32(out var n) ? MapLegacyNumber(n) : AssistantMode.Transform,
            _ => AssistantMode.Transform
        };
    }

    public override void Write(Utf8JsonWriter writer, AssistantMode value, JsonSerializerOptions options) =>
        writer.WriteStringValue(value.ToString());

    private static AssistantMode MapLegacyNumber(int value) => value switch
    {
        0 or 1 or 2 => AssistantMode.Transform,
        3 => AssistantMode.Generate,
        4 => AssistantMode.AnswerClipboard,
        _ => AssistantMode.Transform
    };

    private static AssistantMode MapLegacy(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return AssistantMode.Transform;
        }

        if (Enum.TryParse<AssistantMode>(s, ignoreCase: true, out var direct) && IsCurrentMode(direct))
        {
            return direct;
        }

        return s.Trim() switch
        {
            "Correction" or "Content" or "Social" => AssistantMode.Transform,
            "AnswerGenerate" => AssistantMode.Generate,
            "AnswerEdit" => AssistantMode.AnswerClipboard,
            _ => AssistantMode.Transform
        };
    }

    private static bool IsCurrentMode(AssistantMode mode) =>
        mode is AssistantMode.Transform or AssistantMode.Generate or AssistantMode.AnswerClipboard;
}
