using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace OpenXmlUtil
{
    public static class CoreControl
    {

        public static void ProcessDocuments(string pModel, string pNewDocument)
        {

            var document = CoreControl.CreateDocument(pModel, pNewDocument);

            AppendText(document);

            ProcessMergefield(document);


            document.Close(); //fecha
            document.Dispose(); //tira da memoria 
        }

            public static WordprocessingDocument CreateDocument(string pModel, string pNewDocument)
        {
            System.IO.File.Copy(pModel, pNewDocument, true);
            WordprocessingDocument document = WordprocessingDocument.Open(pNewDocument, true);

            document.ChangeDocumentType(WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = document.MainDocumentPart;

            return document;

        }


        public static void AppendText(WordprocessingDocument document)
        {
            document.ChangeDocumentType(WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = document.MainDocumentPart;
            var mergeFields = mainPart.RootElement.Descendants<FieldCode>(); //Pega todos os MergeFields do corpo do documento

            Body body = document.MainDocumentPart.Document.Body;

            Paragraph p = new Paragraph();
            Run r = new Run();
            string t = "TextoAppend"; //Passar texto por parametro

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(t));

            mainPart.Document.Save();

        }


        public static void ProcessMergefield(WordprocessingDocument document)
        {
            #region Teste Vanessa
            var mergeFields = document.MainDocumentPart.RootElement.Descendants<FieldCode>(); //Pega todos os MergeFields do corpo do documento
                                                                                              ///TODO: ALTERAR PARA BUSCAR DE UMA TABELA DE PARAMETROS QUE INDIQUE O QUE DEVE SER SUBSTITUIDO 

            var mergeFieldName = "SenderFullName";
            var replacementText = string.Empty;

            ClasseTeste teste = new ClasseTeste { SenderFullName = "NomeTeste" }; //mudar aqui pra usar uma classe chamada Bonus - vai ser a entidade que teremos na base de dados

            var propertyInfo = teste.GetType().GetProperty(mergeFieldName).GetValue(teste);
            replacementText = propertyInfo.ToString();

            //essa parte tem que estar num for que percorra a lista de parametros 

            var lstFields = mergeFields
                .Where(f => f.InnerText.Contains(mergeFieldName))
                .ToList(); //Retorna os mergefields com o nome que passamos no corpo do documento TODO: VER COMO TRABALHAR COM MERGEFIELDS NO HEADER E NO FOOTER

            if (lstFields != null)
            {

                foreach (var field in lstFields)
                {
                    Run rFldCode = (Run)field.Parent;

                    Run rBegin = rFldCode.PreviousSibling<Run>(); //Pega o primeiro elemento do mergefield - sempre fldChartType Begin
                    Run rSep = rFldCode.NextSibling<Run>();//Pega o segundo elemento do mergefield - sempre fldChartType  Separate
                    Run rText = rSep.NextSibling<Run>(); //Pega o terceiro nó com o texto do mergefield - nao é um fldChartType, mas um InstrText
                    Run rEnd = rText.NextSibling<Run>();//Pega o quarto elemento do mergefield - sempre fldChartType End

                    Text t = rText.GetFirstChild<Text>();
                    t.Text = replacementText;

                    rFldCode.Remove();
                    rBegin.Remove();
                    rSep.Remove();
                    rEnd.Remove();
                }
            }



            document.MainDocumentPart.Document.Save();

            #endregion

        }
    }
}
