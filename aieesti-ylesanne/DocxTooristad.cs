using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

public class DocxTooristad
{
    public static int LeiaPealkirjaIndeks(List<Paragraph> p, string pealkiri)
    {
        for (int i = 0; i < p.Count; i++)
        {
            var text = LeiaTekst(p[i]);
            if (text.Equals(pealkiri, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }
        return -1;
    }

    public static int LeiaParagraafiIndeks(List<Paragraph> p, int indeks)
    {
        for (int i = indeks; i < p.Count; i++)
        {
            var text = LeiaTekst(p[i]);
            if (!string.IsNullOrWhiteSpace(text))
            {
                var pPr = p[i].ParagraphProperties;
                if (pPr?.ParagraphStyleId == null)
                {
                    return i;
                }
            }
        }
        return -1;
    }

    public static string LeiaTekst(Paragraph p)
    {
        return string.Join(
            "",
            p
            .Elements<Run>()
            .SelectMany(run => run.Elements<Text>())
            .Select(text => text.Text)
        );
    }

    public static Paragraph LooUusParagraaf(string t)
    {
        RunProperties omadused = new RunProperties(
            new RunFonts()
            {
                Ascii = "Arial",
                HighAnsi = "Arial",
                ComplexScript = "Arial"
            }
        );

        return new Paragraph(
            new Run(
                omadused,
                new Text(t) { Space = SpaceProcessingModeValues.Preserve }
            )
        );
    }

    public static string TombaValjaParagraafiTekst(Body b)
    {
        try
        {
            foreach (var p in b.Elements<Paragraph>())
            {
                var text = LeiaTekst(p);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                var pPr = p.ParagraphProperties;
                if (pPr?.ParagraphStyleId == null)
                {
                    return text;
                }
            }

            return string.Empty;
        }
        catch (Exception ex)
        {
            throw new Exception("Viga andmebaasi töötlemisel", ex);
        }
    }

    public static int? AastaArvTekstist(string s)
    {
        var leitud = Regex.Match(s, @"202\d");
        return leitud.Success ? int.Parse(leitud.Value) : null;
    }
}
