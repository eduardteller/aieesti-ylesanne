using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

public class DokumentiRedaktor
{
    public static void UuendaDokumentiAndmed(
        Body dok,
        string loiguPealkiri,
        string uusTekst
    )
    {
        if (dok == null)
            throw new ArgumentNullException(nameof(dok), "Dokument on tühi");
        if (string.IsNullOrWhiteSpace(loiguPealkiri))
            throw new ArgumentException("Pealikiri on tühi", nameof(loiguPealkiri));
        if (string.IsNullOrWhiteSpace(uusTekst))
            throw new ArgumentException("Uus paragraaf on tühi", nameof(uusTekst));

        try
        {
            var paragraafid = dok.Elements<Paragraph>().ToList();
            int pealkirjaIndeks = LeiaPealkirjaIndeks(paragraafid, loiguPealkiri);

            if (pealkirjaIndeks == -1)
            {
                throw new ArgumentException($"Pealkiri '{loiguPealkiri}' ei eksisteeri");
            }

            int paragraafiIndeks = LeiaParagraafiIndeks(paragraafid, pealkirjaIndeks + 1);

            if (paragraafiIndeks == -1)
            {
                throw new ArgumentException($"Paragraaf mida peaks muutma, ei leitud");
            }
            else
            {
                var uusParagraaf = LooUusParagraaf(uusTekst);
                dok.ReplaceChild(uusParagraaf, paragraafid[paragraafiIndeks]);

            }
        }
        catch (Exception ex)
        {
            throw new Exception("Tekkis viga dokumenti andmete uuendamisel", ex);
        }
    }

    private static int LeiaPealkirjaIndeks(List<Paragraph> p, string pealkiri)
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

    private static int LeiaParagraafiIndeks(List<Paragraph> p, int indeks)
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

    private static string LeiaTekst(Paragraph p)
    {
        return string.Join(
            "",
            p
            .Elements<Run>()
            .SelectMany(run => run.Elements<Text>())
            .Select(text => text.Text)
        );
    }

    private static Paragraph LooUusParagraaf(string t)
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
}
