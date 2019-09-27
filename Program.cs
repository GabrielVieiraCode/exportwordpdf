using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;

namespace Repositorio_exporwordpdf
{
    class Program
    {
        static void Main(string[] args)
        {
            //Criando um novo documento com o nome documento
            Document Documento = new Document();

            //Criando um seção dento do documento
            Section secaoCapa = Documento.AddSection();

            //Insere um tiutlo na primeira pagina
            Paragraph titulo = secaoCapa.AddParagraph();

            //Aqui eu insiro na minha variavel titulo o valor da string "Titulo do meu documento"
            // No meu documento aparecera "Titulo do meu documento"
            titulo.AppendText("Titulo do meu documento\n\n");
            
            //Linha horizontalmente o titulo
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;
             

            // instaciando a classe ParagraphStyle dentro do nosso documento
            ParagraphStyle estilo01 = new ParagraphStyle(Documento);

            //Define o nome da classe estilo01
            estilo01.Name = "Cor do titulo";
            
            //Colore a propriedade TextColor de azul escuro
            estilo01.CharacterFormat.TextColor = Color.DarkBlue;

            //Transformar a propriedade bold em true
            estilo01.CharacterFormat.Bold = true;

            //Adicionar e colocar com usavel no nosso documento 
            Documento.Styles.Add(estilo01);
                      
            titulo.ApplyStyle(estilo01.Name);

           
            Paragraph paragrafo = secaoCapa.AddParagraph();
            paragrafo.AppendText("Texto para lorem lorem lorem lorem lorem lorem\n\nlorem lore lorem lorem lorem");
            

            Documento.SaveToFile(@"Saida\exemploWord.docx",FileFormat.Docx);





            

        }
    }
}
