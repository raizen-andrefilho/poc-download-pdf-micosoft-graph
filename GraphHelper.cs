using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using IO = System.IO;


// https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0&tabs=http
// https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http

class GraphHelper
{
    private static Settings? _settings;
    private static DeviceCodeCredential? _deviceCodeCredential;
    private static GraphServiceClient? _userClient;

    /// <summary>
    /// Inicializa o Graph Service Client utilizando o Device Code Credential
    /// que permitirá que o dispositivo tenha acesso aos métodos e informações
    /// da api do usuário que se autenticar;
    /// </summary>
    /// <param name="settings">Parametros de configuração para o processo de inicialização</param>
    /// <param name="deviceCodePrompt">Função que irá exibir a mensagem de autenticação</param>
    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.TenantId, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }

    /// <summary>
    /// Recupera o token do usuário gerado para o escopo
    /// </summary>
    /// <returns>String contendo o token</returns>
    public static async Task<string> GetUserTokenAsync()
    {
        if(_deviceCodeCredential == null)
            throw new NullReferenceException("O 'Graph Service Client' não foi inicializado");

        if(_settings?.GraphUserScopes == null)
            throw new ArgumentNullException("A configuração 'GraphUserScopes' não pode ser nula");

        // Recupera o token gerado para o escopo
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }

    /// <summary>
    /// Recupera as informações do usuário
    /// </summary>
    public static Task<User> GetUserAsync()
    {
        if(_userClient == null)
            throw new NullReferenceException("O 'Graph Service Client' não foi inicializado");

        return _userClient.Me
            .Request()
            .Select(u => new
            {
                u.DisplayName,
                u.Mail,
                u.UserPrincipalName
            })
            .GetAsync();
    }
    
    public async static Task<string> UploadFileOnDrive(string file)
    {
        if(_userClient == null)
            throw new NullReferenceException("O 'Graph Service Client' não foi inicializado");

        var fileStream = System.IO.File.OpenRead(file);

        var uploadProps = new DriveItemUploadableProperties
        {
            AdditionalData = new Dictionary<string, object>
            {
                { "@microsoft.graph.conflictBehavior", "replace" }
            }
        };

        var uploadSession = await _userClient.Me.Drive.Root
            .ItemWithPath(Guid.NewGuid() + ".docx")
            .CreateUploadSession(uploadProps)
            .Request()
            .PostAsync();

        int maxSliceSize = 320 * 1024;
        var fileUploadTask =
            new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

        var totalLength = fileStream.Length;
        IProgress<long> progress = new Progress<long>(prog => {
            Console.WriteLine($"Enviado {prog} bytes de {totalLength} bytes");
        });

        try
        {
            var uploadResult = await fileUploadTask.UploadAsync(progress);

            Console.WriteLine(uploadResult.UploadSucceeded ?
                $"Envio concluído, item ID: {uploadResult.ItemResponse.Id}" :
                "Erro no envio");

            return uploadResult.ItemResponse.Id;
        }
        catch (ServiceException ex)
        {
            Console.WriteLine($"Error no envio: {ex.ToString()}");
            throw ex;
        }
    }

    public async static Task<string> DownloadFileFromDrive(string itemId, string format = "pdf")
    {
        if(_userClient == null)
            throw new NullReferenceException("O 'Graph Service Client' não foi inicializado");

        var queryOptions = new List<QueryOption>() { new QueryOption("format", $"{format}") };

        var stream = await _userClient.Me.Drive.Items[$"{itemId}"].Content
            .Request( queryOptions )
            .GetAsync();

        string newFileNew = Guid.NewGuid() + "." + format;
        using (var fileStream = IO.File.Create(newFileNew))
        {
            stream.CopyTo(fileStream);
        }

        return newFileNew;
    }
}