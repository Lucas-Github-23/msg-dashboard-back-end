namespace MsgDashboardBackend.Models;
public class EmailDto
    {
    public string Id { get; set; } = string.Empty;
    public string Remetente { get; set; } = string.Empty;
    public string? Cc { get; set; } // A interrogação significa que pode ser nulo (nem todo e-mail tem Cc)
    public string Assunto { get; set; } = string.Empty;
    public string Data { get; set; } = string.Empty;
    public string Corpo { get; set; } = string.Empty;
    public List<AnexoDto> Anexos { get; set; } = new();
    }
public class AnexoDto
    {
    public string Id { get; set; } = string.Empty;
    public string Nome { get; set; } = string.Empty;
    public long Tamanho { get; set; } = 0;
    }

