using System.Globalization;
using MagicVoice.Core;

namespace MagicVoice;

public sealed class MagicVoiceProfile : IAppProfile
{
    public string AppName => "Magic-Voice";
    public string DataFolderName => "Magic-Voice";
    public string AuthorName => "Ronny Schulz";
    public string CopyrightText => "Copyright 2026 Ronny Schulz";
    public string LicenseName => "MIT-Lizenz";
    public string MutexName => "Global\\Magic-Voice.Singleton";
    public string AumId => "RonnySchulz.MagicVoice";
    public string AutostartRegistryValueName => "Magic-Voice";

    public string SystemPrompt =>
        "Du bist ein mehrsprachiger Schreib- und Textassistent. Antworte ausschließlich mit dem finalen Text, ohne Markdown, Erklärung oder Anführungszeichen.";

    public IReadOnlyList<AssistantModeDefinition> Modes { get; } =
    [
        new(
            AssistantMode.Transform,
            "Text bearbeiten",
            "Wendet deine Anweisung, den Schreibstil, die Intensität und die Absatz-Optionen aus der UI auf das Transkript an.",
            "Ctrl+Shift+1",
            "Bearbeite das Transkript gemäß dieser Anweisung. Ohne weitere Vorgabe: Rechtschreibung und Grammatik korrigieren, leicht glätten, Aussage beibehalten."),
        new(
            AssistantMode.Generate,
            "Generieren",
            "Erzeugt Text nur aus deiner gesprochenen Anweisung (Schreibstil, Intensität und Absätze kommen aus der UI).",
            "Ctrl+Shift+2",
            "Erzeuge einen Text gemäß der gesprochenen Anweisung. Schreibe direkt den fertigen Text und gib keine Vorbemerkungen aus."),
        new(
            AssistantMode.AnswerClipboard,
            "Antwort (Zwischenablage)",
            "Nutzt den Text in der Zwischenablage als Quelle und deine gesprochene Anweisung; Stil, Intensität und Absätze steuerst du in der UI.",
            "Ctrl+Shift+3",
            "Beantworte den unten angegebenen Quelltext gemäß der gesprochenen Anweisung. Gib ausschließlich die Antwort aus. Wiederhole, zitiere oder kopiere den Quelltext nicht in die Ausgabe.",
            RequiresClipboardSource: true)
    ];

    public string IntensityStepName(AssistantMode mode, int intensity)
    {
        var step = Math.Clamp(intensity, Defaults.MinModeIntensity, Defaults.MaxModeIntensity);
        return mode switch
        {
            AssistantMode.Transform => step switch
            {
                1 => "minimal",
                2 => "leicht",
                3 => "ausgewogen",
                4 => "deutlich",
                5 => "stark",
                _ => "ausgewogen"
            },
            AssistantMode.Generate => step switch
            {
                1 => "streng wörtlich",
                2 => "nah am Auftrag",
                3 => "ausgewogen",
                4 => "freier",
                5 => "sehr kreativ",
                _ => "ausgewogen"
            },
            AssistantMode.AnswerClipboard => step switch
            {
                1 => "sehr knapp",
                2 => "kompakt",
                3 => "ausgewogen",
                4 => "ausführlich",
                5 => "sehr ausführlich",
                _ => "ausgewogen"
            },
            _ => step.ToString(CultureInfo.InvariantCulture)
        };
    }

    public string IntensityStepInstruction(AssistantMode mode, int intensity)
    {
        var step = Math.Clamp(intensity, Defaults.MinModeIntensity, Defaults.MaxModeIntensity);
        return mode switch
        {
            AssistantMode.Transform => step switch
            {
                1 => "Bleibe sehr nah am Transkript: nur offensichtliche Tippfehler und minimale Zeichensetzung.",
                2 => "Korrigiere Rechtschreibung, Grammatik und Zeichensetzung; erhalte Ton, Wortwahl und Satzbau.",
                3 => "Korrigiere und glätte unklare Stellen; bleibe inhaltlich beim Original.",
                4 => "Formuliere deutlich klarer und lesbarer; du darfst Sätze umbauen, wenn die Aussage gleich bleibt.",
                5 => "Darf kräftig umarbeiten und polieren, solange die Anweisung aus dem Auftrag nicht widersprochen wird.",
                _ => string.Empty
            },
            AssistantMode.Generate => step switch
            {
                1 => "Halte dich strikt an die Anweisung und ergänze keine eigenen Inhalte.",
                2 => "Folge der Anweisung eng und ergänze nur, wo unbedingt nötig.",
                3 => "Setze die Anweisung sinnvoll um und ergänze passend, ohne abzuschweifen.",
                4 => "Setze die Anweisung großzügig um und ergänze hilfreiche Details.",
                5 => "Setze die Anweisung sehr frei und kreativ um, ergänze passende Ideen.",
                _ => string.Empty
            },
            AssistantMode.AnswerClipboard => step switch
            {
                1 => "Antworte sehr knapp und strikt entlang der Anweisung.",
                2 => "Antworte kompakt und nah am Auftrag, ohne unnötige Details.",
                3 => "Antworte ausgewogen mit den nötigen Details.",
                4 => "Antworte ausführlich und ergänze hilfreiche Details.",
                5 => "Antworte sehr ausführlich und decke das Thema breit ab.",
                _ => string.Empty
            },
            _ => string.Empty
        };
    }

    public string WritingStyleInstruction(WritingStyle style) => style switch
    {
        WritingStyle.Casual => "Schreibstil: locker und alltagssprachlich, gerne kurze Sätze und persönliche Ansprache.",
        WritingStyle.Neutral => "Schreibstil: neutral und sachlich, weder formell noch zu locker.",
        WritingStyle.Professional => "Schreibstil: professionell und höflich, klare Struktur, ohne Floskeln.",
        WritingStyle.Academic => "Schreibstil: wissenschaftlich-präzise, vollständige Sätze, Fachbegriffe wo passend, neutraler Ton.",
        _ => string.Empty
    };

    public string EmojiExpressionInstruction(EmojiExpression level) => level switch
    {
        EmojiExpression.None =>
            "Emojis und Ausdruck: verwende keine Emojis und keine Text-Emoticons (z. B. :-) ). Halte die Ausgabe sachlich und rein textbasiert.",
        EmojiExpression.Sparse =>
            "Emojis und Ausdruck: setze Emojis oder Text-Emoticons höchstens sehr sparsam und nur, wenn sie den Inhalt klar unterstützen.",
        EmojiExpression.Balanced =>
            "Emojis und Ausdruck: du darfst Emojis gelegentlich nutzen, wenn sie zum Kontext passen; vermeide einen überladenen Emoji-Stil.",
        EmojiExpression.Lively =>
            "Emojis und Ausdruck: nutze Emojis und einen lebhaften Social-Media-Ton; mehrere Emojis pro Abschnitt sind in Ordnung, wenn es zur Stimmung passt.",
        EmojiExpression.Heavy =>
            "Emojis und Ausdruck: nutze Emojis großzügig und häufig für einen starken Social-Media-Charakter; der Text soll ausdrucksvoll wirken, aber lesbar bleiben.",
        _ => string.Empty
    };

    public List<AssistantInstance> CreateDefaultAssistants() =>
        Modes.Select(mode => new AssistantInstance
        {
            Id = Guid.NewGuid().ToString("N"),
            Type = mode.Mode,
            Name = mode.Name,
            Hotkey = mode.DefaultHotkey,
            Prompt = mode.DefaultPrompt,
            Intensity = Defaults.DefaultModeIntensity,
            WritingStyle = WritingStyle.Neutral,
            ParagraphDensity = ParagraphDensity.Balanced,
            EmojiExpression = EmojiExpression.Balanced
        }).ToList();
}
