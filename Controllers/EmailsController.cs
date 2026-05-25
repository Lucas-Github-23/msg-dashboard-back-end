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
            
            var emailDto = new EmailDto
            {
                // Usamos o nome do arquivo como ID para facilitar
                Id = Path.GetFileNameWithoutExtension(arquivo), 
                Assunto = msg.Subject ?? "Sem Assunto",
                // Tenta pegar o e-mail, se não tiver, pega o nome de exibição
                Remetente = msg.Sender?.Email ?? msg.Sender?.DisplayName ?? "Desconhecido", 
                Cc = "",
                Data = msg.SentOn?.ToString("dd/MM/yyyy HH:mm") ?? "Data desconhecida",
                // Prioriza o HTML, mas se for um e-mail simples, pega o texto puro
                Corpo = msg.BodyHtml ?? msg.BodyText ?? "Conteúdo vazio" 
            };

            // 4. Verifica e extrai os anexos
            foreach (var anexo in msg.Attachments)
            {
                if (anexo is Storage.Attachment fileAttachment)
                {
                    // Calcula o tamanho do arquivo de Bytes para Kilobytes
                    var tamanhoBytes = fileAttachment.Data?.Length ?? 0;
                    var tamanhoKb = tamanhoBytes / 1024;

                    emailDto.Anexos.Add(new AnexoDto
                    {
                        Id = Guid.NewGuid().ToString(),
                        Nome = fileAttachment.FileName,
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
}