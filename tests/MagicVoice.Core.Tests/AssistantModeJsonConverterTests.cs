using System.Text.Json;
using System.Text.Json.Serialization;
using MagicVoice.Core;
using Xunit;

namespace MagicVoice.Core.Tests;

public class AssistantModeJsonConverterTests
{
    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web)
    {
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };

    [Theory]
    [InlineData("Correction", AssistantMode.Transform)]
    [InlineData("Content", AssistantMode.Transform)]
    [InlineData("Social", AssistantMode.Transform)]
    [InlineData("AnswerGenerate", AssistantMode.Generate)]
    [InlineData("AnswerEdit", AssistantMode.AnswerClipboard)]
    [InlineData("Transform", AssistantMode.Transform)]
    [InlineData("Generate", AssistantMode.Generate)]
    [InlineData("AnswerClipboard", AssistantMode.AnswerClipboard)]
    public void Deserialisiert_gespeicherte_Typnamen(string typeName, AssistantMode expected)
    {
        var json =
            $$"""{"id":"a","type":"{{typeName}}","name":"","hotkey":"","prompt":"","intensity":3,"writingStyle":"Neutral","paragraphDensity":"Balanced"}""";
        var assistant = JsonSerializer.Deserialize<AssistantInstance>(json, Options);
        Assert.NotNull(assistant);
        Assert.Equal(expected, assistant.Type);
    }

    [Fact]
    public void Deserialisiert_alte_numerische_Typwerte()
    {
        var json = """{"id":"a","type":4,"name":"","hotkey":"","prompt":"","intensity":3,"writingStyle":"Neutral","paragraphDensity":"Balanced"}""";
        var assistant = JsonSerializer.Deserialize<AssistantInstance>(json, Options);
        Assert.Equal(AssistantMode.AnswerClipboard, assistant!.Type);
    }
}
