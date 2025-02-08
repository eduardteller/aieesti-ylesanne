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
            int pealkirjaIndeks = DocxTooristad.LeiaPealkirjaIndeks(paragraafid, loiguPealkiri);

            if (pealkirjaIndeks == -1)
            {
                throw new ArgumentException($"Pealkiri '{loiguPealkiri}' ei eksisteeri");
            }

            int paragraafiIndeks = DocxTooristad.LeiaParagraafiIndeks(paragraafid, pealkirjaIndeks + 1);

            if (paragraafiIndeks == -1)
            {
                throw new ArgumentException($"Paragraaf mida peaks muutma, ei leitud");
            }
            else
            {
                var uusParagraaf = DocxTooristad.LooUusParagraaf(uusTekst);
                dok.ReplaceChild(uusParagraaf, paragraafid[paragraafiIndeks]);

            }
        }
        catch (Exception ex)
        {
            throw new Exception("Tekkis viga dokumenti andmete uuendamisel", ex);
        }
    }


}
