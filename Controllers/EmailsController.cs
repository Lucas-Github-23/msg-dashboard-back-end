using Microsoft.AspNetCore.Mvc;
using MsgReader.Outlook;
using MsgDashboardBackend.Models;
using System.IO;

namespace MsgDashboardBackend.Controllers;

// Essas anotações dizem pro C# que essa classe é uma rota de API
[ApiController]
[Route("api/[controller]")] 
public class EmailsController : ControllerBase
{
    [HttpGet] // Isso indica que a rota vai responder a requisições do tipo GET (buscar dados)
    public IActionResult GetEmails()
    {
        // 1. Aponta para a pasta "Emails" que você criou na raiz do projeto
        var pastaEmails = Path.Combine(Directory.GetCurrentDirectory(), "Emails");
        var listaDeEmails = new List<EmailDto>();

        // Se a pasta não existir, a gente só retorna uma lista vazia e evita um erro
        if (!Directory.Exists(pastaEmails))
        {
            return Ok(listaDeEmails);
        }

        // 2. Pega todos os arquivos terminados em .msg dentro da pasta
        var arquivosMsg = Directory.GetFiles(pastaEmails, "*.msg");

        // 3. Passa por cada arquivo extraindo os dados
        foreach (var arquivo in arquivosMsg)
        {
            // O MsgReader abre o arquivo binário aqui
            using var msg = new Storage.Message(arquivo);
            var emailId = Path.GetFileNameWithoutExtension(arquivo);

            var corpoHtml = msg.BodyHtml ?? msg.BodyText ?? "Conteúdo vazio";
            if (!string.IsNullOrEmpty(msg.BodyHtml))
            {
                corpoHtml = System.Text.RegularExpressions.Regex.Replace(msg.BodyHtml, @"(?i)cid:([^""'\s>]+)", match =>
                {
                    var cid = match.Groups[1].Value;
                    return $"/api/Emails/{emailId}/attachment/{Uri.EscapeDataString(cid)}";
                });
            }

            var emailDto = new EmailDto
            {
                // Usamos o nome do arquivo como ID para facilitar
                Id = emailId, 
                Assunto = msg.Subject ?? "Sem Assunto",
                // Tenta pegar o e-mail, se não tiver, pega o nome de exibição
                Remetente = msg.Sender?.Email ?? msg.Sender?.DisplayName ?? "Desconhecido", 
                Para = msg.GetEmailRecipients(RecipientType.To, false, false),
                Cc = msg.GetEmailRecipients(RecipientType.Cc, false, false),
                Data = msg.SentOn?.ToString("dd/MM/yyyy HH:mm") ?? "Data desconhecida",
                Corpo = corpoHtml,
                Preview = msg.BodyText ?? ""
            };

            // 4. Verifica e extrai os anexos
            foreach (var anexo in msg.Attachments)
            {
                if (anexo is Storage.Attachment fileAttachment)
                {
                    // Ignora imagens inline (usadas no corpo do e-mail como rodapés e assinaturas)
                    var isInline = fileAttachment.IsInline;
                    var hasContentId = !string.IsNullOrEmpty(fileAttachment.ContentId);
                    
                    if (isInline || hasContentId)
                    {
                        continue;
                    }

                    // Calcula o tamanho do arquivo de Bytes para Kilobytes
                    var tamanhoBytes = fileAttachment.Data?.Length ?? 0;
                    var tamanhoKb = (int)Math.Max(1, tamanhoBytes / 1024.0);

                    emailDto.Anexos.Add(new AnexoDto
                    {
                        Id = Guid.NewGuid().ToString(),
                        Nome = fileAttachment.FileName ?? "anexo_sem_nome",
                        Tamanho = tamanhoKb
                    });
                }
            }

            // Adiciona o e-mail processado na nossa lista final
            listaDeEmails.Add(emailDto);
        }

        // 5. Retorna a lista pronta (o C# transforma isso em JSON automaticamente)
        return Ok(listaDeEmails);
    }

    [HttpGet("{emailId}/attachment/{contentId}")]
    public IActionResult GetAttachment(string emailId, string contentId)
    {
        var pastaEmails = Path.Combine(Directory.GetCurrentDirectory(), "Emails");
        var filePath = Path.Combine(pastaEmails, emailId + ".msg");
        if (!System.IO.File.Exists(filePath))
        {
            return NotFound("E-mail não encontrado.");
        }

        var decodificadoContentId = System.Net.WebUtility.UrlDecode(contentId);

        using var msg = new Storage.Message(filePath);
        foreach (var anexo in msg.Attachments)
        {
            if (anexo is Storage.Attachment fileAttachment)
            {
                if (fileAttachment.ContentId == decodificadoContentId || 
                    fileAttachment.FileName == decodificadoContentId ||
                    fileAttachment.ContentId == contentId ||
                    fileAttachment.FileName == contentId)
                {
                    if (fileAttachment.Data == null)
                    {
                        return NotFound("Dados do anexo vazios.");
                    }
                    var stream = new MemoryStream(fileAttachment.Data);
                    var contentType = GetContentType(fileAttachment.FileName);
                    return File(stream, contentType, fileAttachment.FileName);
                }
            }
        }

        return NotFound("Anexo não encontrado.");
    }

    private string GetContentType(string fileName)
    {
        var extension = Path.GetExtension(fileName).ToLowerInvariant();
        return extension switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".pdf" => "application/pdf",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".msg" => "application/vnd.ms-outlook",
            _ => "application/octet-stream"
        };
    }
}