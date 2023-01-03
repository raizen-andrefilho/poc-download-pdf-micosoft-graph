Splash.Show();

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

// Greet the user by name
await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.Clear();
    Console.WriteLine("Selecione a opção da POC:");
    Console.WriteLine("0. Sair");
    Console.WriteLine("1. Exibir o Token");
    Console.WriteLine("2. Converter Arquivo");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        choice = -1;
    }

    switch(choice)
    {
        case 0:
            Console.WriteLine("Tchau...");
            break;
        case 1:
            await DisplayAccessTokenAsync();
            break;
        case 2:
            await ConverterFile();
            break;
        default:
            Console.WriteLine("Opção inválida.");
            break;
    }
}

void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForUserAuth(settings,
        (info, cancel) =>
        {
            Console.WriteLine(info.Message);
            return Task.FromResult(0);
        });
}

async Task GreetUserAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        Console.WriteLine($"Olá, {user?.DisplayName}!");
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Erro ao recuperar o usuário: {ex.Message}");
    }
}

async Task DisplayAccessTokenAsync()
{
    try
    {
        var userToken = await GraphHelper.GetUserTokenAsync();
        
        Console.Clear();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"Token: {userToken}");
        Console.ResetColor();
        Console.WriteLine($"Pressione <enter> para voltar ao menu");
        Console.ReadLine();
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Erro ao recuprar o token: {ex.Message}");
    }
}

async Task ConverterFile()
{
    try
    {
        Console.Clear();

string fileName = "docx-teste.docx";
        var fileId = await GraphHelper.UploadFileOnDrive(fileName);
        var newFileName = await GraphHelper.DownloadFileFromDrive(fileId, "pdf");

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Arquivo convertido: {newFileName}");
        Console.ResetColor();

        Console.WriteLine($"Pressione <enter> para voltar ao menu");
        Console.ReadLine();

    }
    catch (Exception ex)
    {
        Console.WriteLine($"Erro ao convrter o arquivo: {ex.Message}");
    }
}

