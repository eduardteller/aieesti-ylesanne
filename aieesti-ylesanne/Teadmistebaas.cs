using DocumentFormat.OpenXml.Packaging;

public class Teadmistebaas
{
    public string MakromajanduslikTaust { get; private set; } = string.Empty;
    public string EestiKinnisvaraturg { get; private set; } = string.Empty;
    public string PiirkondlikYlevaade { get; private set; } = string.Empty;
    private string lokatsioonLocal { get; set; } = string.Empty;
    public List<string> teadmistebaasiFailideNimed { get; private set; } = new List<string>();

    public Teadmistebaas(string lokatsioon)
    {

        if (string.IsNullOrWhiteSpace(lokatsioon))
        {
            throw new ArgumentNullException(nameof(lokatsioon), "Lokatsioon ei saa olla tühi");
        }

        lokatsioonLocal = lokatsioon;

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
                teadmistebaasiFailideNimed.Add("teadmistebaas_majandus_2025.docx");
                var t = DocxTooristad.TombaValjaParagraafiTekst(wordDoc.MainDocumentPart.Document.Body);
                MakromajanduslikTaust = t;
            }

            using (var wordDoc = WordprocessingDocument.Open(Path.Combine(andmebaasiKaust, "teadmistebaas_Üldülevaated_Eesti_Tln_Harjumaa_2025.docx"), false))
            {
                teadmistebaasiFailideNimed.Add("teadmistebaas_Üldülevaated_Eesti_Tln_Harjumaa_2025.docx");
                var t = DocxTooristad.TombaValjaParagraafiTekst(wordDoc.MainDocumentPart.Document.Body);
                EestiKinnisvaraturg = t;
            }

            var lokatsioonDoc = $"teadmistebaas_{lokatsioon.ToLower()}_2025.docx";

            using (var wordDoc = WordprocessingDocument.Open(Path.Combine(andmebaasiKaust, lokatsioonDoc), false))
            {
                teadmistebaasiFailideNimed.Add(lokatsioonDoc);
                var t = DocxTooristad.TombaValjaParagraafiTekst(wordDoc.MainDocumentPart.Document.Body);
                PiirkondlikYlevaade = t;
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Viga andmebaasi failide parsimisel", ex);
        }
    }

    public void UuendaPiirkondlikYlevaade(string uusTekst)
    {

        string programmiLokatsioon = AppContext.BaseDirectory;
        string andmebaasiKaust = Path.Combine(programmiLokatsioon, "database");
        try
        {
            var lokatsioonDoc = $"teadmistebaas_{lokatsioonLocal.ToLower()}_2025.docx";

            using (var wordDoc = WordprocessingDocument.Open(Path.Combine(andmebaasiKaust, lokatsioonDoc), false))
            {
                DokumentiRedaktor.UuendaDokumentiAndmed(
                    wordDoc.MainDocumentPart.Document.Body,
                    $"{lokatsioonLocal} linnaosa korteriturg",
                    uusTekst
                );

                wordDoc.MainDocumentPart.Document.Save();
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Viga teadmistebaasi uuendamisel", ex);
        }
    }
}
