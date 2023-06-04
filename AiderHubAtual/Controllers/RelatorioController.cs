using Aspose.Slides;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using MimeKit;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using AiderHubAtual.Models;
using DocumentFormat.OpenXml;
using Paragraph = iText.Layout.Element.Paragraph;

namespace AiderHubAtual.Controllers
{
    public class RelatorioController : Controller
    {
        private readonly Context _context;
        private readonly IConfiguration _configuration;

        public RelatorioController(Context context, IConfiguration configuration)
        {
            _context = context;
            _configuration = configuration;
        }

        public IActionResult GerarCertificado(int id)
        {
            var certificado = _context.Relatorios.Find(id);

            if (certificado == null)
            {
                // Certificado não encontrado, retorne uma mensagem de erro ou redirecione para uma página de erro
                return NotFound();
            }

            // Gerar o certificado em PDF
            byte[] pdfCertificado = GerarCertificadoPDF(certificado);

            // Salvar o certificado em um arquivo temporário
            var tempFilePath = Path.GetTempFileName() + ".pdf";
            System.IO.File.WriteAllBytes(tempFilePath, pdfCertificado);

            // Abrir o arquivo PowerPoint
            var pptxPath = Path.Combine(Directory.GetCurrentDirectory(), "Models", "certificado_de_conclusão_modelo.pptx");
            using (var presentation = new Presentation(pptxPath))
            {
                var slide = presentation.Slides.First();

                // Substituir os marcadores de lugar no slide pelos dados do certificado
                foreach (var shape in slide.Shapes)
                {
                    if (shape is AutoShape autoShape)
                    {

                        if (autoShape.TextFrame.Text.Contains("<<Nome>>"))
                            autoShape.TextFrame.Text = autoShape.TextFrame.Text.Replace("<<Nome>>", certificado.NomeVoluntario);

                        if (autoShape.TextFrame.Text.Contains("<<Email>>"))
                            autoShape.TextFrame.Text = autoShape.TextFrame.Text.Replace("<<Email>>", certificado.Email);
                    }
                }

                // Salvar o certificado modificado em um arquivo temporário
                string pptxTempFilePath = Path.GetTempFileName() + ".pptx";
                using (PresentationDocument presentationDocument = PresentationDocument.Create(pptxTempFilePath, PresentationDocumentType.Presentation))

                // Enviar o certificado por e-mail
                EnviarCertificadoPorEmail(certificado, pdfCertificado, pptxTempFilePath);
            }

            // Retorne uma mensagem de sucesso ou redirecione para uma página de confirmação
            ViewBag.Mensagem = "Certificado gerado e enviado com sucesso!";
            return View("Relatorio");
        }

        private byte[] GerarCertificadoPDF(Relatorio certificado)
        {
            using (var memoryStream = new MemoryStream())
            {
                var writer = new PdfWriter(memoryStream);
                var pdfDocument = new PdfDocument(writer);
                var document = new Document(pdfDocument);

                // Adicionar o conteúdo do certificado
                document.Add(new Paragraph("Certificado"));
                document.Add(new Paragraph($"Nome: {certificado.NomeVoluntario}"));
                document.Add(new Paragraph($"Email: {certificado.Email}"));

                // Fechar o documento PDF
                document.Close();

                return memoryStream.ToArray();
            }
        }

        private void EnviarCertificadoPorEmail(Relatorio certificado, byte[] pdfCertificado, string pptxFilePath)
        {
            var emailSettings = _configuration.GetSection("EmailSettings").Get<EmailSettings>();

            // Configurar o cliente SMTP
            var smtpClient = new SmtpClient();
            smtpClient.Connect(emailSettings.Host, emailSettings.Port, SecureSocketOptions.StartTls);
            smtpClient.Authenticate(emailSettings.Username, emailSettings.Password);

            // Crie a mensagem de e-mail
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(emailSettings.SenderName, emailSettings.SenderEmail));
            message.To.Add(new MailboxAddress(certificado.NomeVoluntario, certificado.Email));
            message.Subject = "Certificado";

            // Corpo do e-mail em texto
            var textBody = new TextPart("plain")
            {
                Text = "Segue o certificado em anexo."
            };

            // Anexar o certificado em PDF
            var pdfAttachment = new MimePart("application", "pdf")
            {
                Content = new MimeContent(new MemoryStream(pdfCertificado)),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = "Certificado.pdf"
            };

            // Anexar o arquivo PowerPoint
            var pptxAttachment = new MimePart("application", "octet-stream")
            {
                Content = new MimeContent(new FileStream(pptxFilePath, FileMode.Open)),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = "Certificado.pptx"
            };

            // Adicionar as partes ao corpo da mensagem
            var multipart = new Multipart("mixed");
            multipart.Add(textBody);
            multipart.Add(pdfAttachment);
            multipart.Add(pptxAttachment);

            message.Body = multipart;

            // Enviar o e-mail
            smtpClient.Send(message);

            // Desconectar o cliente SMTP
            smtpClient.Disconnect(true);
        }
    }
}
