using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class Teadmistebaas
{
    public string MakromajanduslikTaust { get; private set; } = string.Empty;
    public string EestiKinnisvaraturg { get; private set; } = string.Empty;
    public string Linnaosa { get; private set; } = string.Empty;

    public Teadmistebaas(string lokatsioon)
    {

        if (string.IsNullOrWhiteSpace(lokatsioon))
        {
            throw new ArgumentNullException(nameof(lokatsioon), "Lokatsioon ei saa olla tühi");
        }

        ParsiAndmebaasiFailid(lokatsioon);
    }

    private void ParsiAndmebaasiFailid(string lokatsioon)
    {

        string programmiLokatsioon = AppContext.BaseDirectory;
        string andmebaasiKaust = Path.Combine(programmiLokatsioon, "database");
        try
        {
            using (var wordDoc = WordprocessingDocument.Open(Path.Combine(andmebaasiKaust, "teadmistebaas_majandus_2025.docx"), false))
            {
                var t = TombaValjaTekst(wordDoc.MainDocumentPart.Document.Body);
                MakromajanduslikTaust = t;
            }

            using (var wordDoc = WordprocessingDocument.Open(Path.Combine(andmebaasiKaust, "teadmistebaas_Üldülevaated_Eesti_Tln_Harjumaa_2025.docx"), false))
            {
                var t = TombaValjaTekst(wordDoc.MainDocumentPart.Document.Body);
                EestiKinnisvaraturg = t;
            }

            var lokatsioonDoc = $"teadmistebaas_{lokatsioon.ToLower()}_2025.docx";

            using (var wordDoc = WordprocessingDocument.Open(Path.Combine(andmebaasiKaust, lokatsioonDoc), false))
            {
                var t = TombaValjaTekst(wordDoc.MainDocumentPart.Document.Body);
                Linnaosa = t;
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Viga andmebaasi failide parsimisel", ex);
        }

    }

    private string TombaValjaTekst(Body b)
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

    private string LeiaTekst(Paragraph p)
    {
        try
        {
            return string.Join(
                "",
                p
                    .Elements<Run>()
                    .SelectMany(run => run.Elements<Text>())
                    .Select(text => text.Text)
            );
        }
        catch (Exception ex)
        {
            throw new Exception("Viga teksti extractimisel", ex);
        }
    }
}
