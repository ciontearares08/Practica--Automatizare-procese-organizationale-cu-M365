using Azure.Core;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DotNetEnv; // ⚠️ trebuie instalat cu NuGet: DotNetEnv

namespace TeamsAutomation
{
    class Program
    {
        // Vor fi setate din environment
        static string accessToken;
        static string teamId;

        static string inputFilePath = "input.txt";
        static string logFilePath = "automation.log";

        static async Task Main(string[] args)
        {
            try
            {
                Console.WriteLine("=== Teams Automation - Pornire ===");

                // Încarcă variabilele din .env
                Env.Load();
                accessToken = Environment.GetEnvironmentVariable("AZURE_TOKEN");
                teamId = Environment.GetEnvironmentVariable("TEAM_ID");

                if (string.IsNullOrEmpty(accessToken) || string.IsNullOrEmpty(teamId))
                {
                    Console.WriteLine("❌ Te rog setează AZURE_TOKEN și TEAM_ID în fișierul .env!");
                    return;
                }

                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Eroare: Fișierul {inputFilePath} nu există!");
                    CreateSampleInputFile();
                    Console.WriteLine($"Am creat un fișier exemplu {inputFilePath}. Completează-l și rulează din nou.");
                    return;
                }

                var (channelName, members, filePath) = ReadInputFile(inputFilePath);

                Console.WriteLine($"Canal: {channelName}");
                Console.WriteLine($"Membri/Guests: {members.Count}");
                Console.WriteLine($"Fișier: {filePath}");

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Eroare: Fișierul {filePath} nu există!");
                    return;
                }

                var client = GetGraphClient();

                using (var log = new StreamWriter(logFilePath, append: true, Encoding.UTF8))
                {
                    await log.WriteLineAsync($"\n=== Sesiune nouă - {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");

                    var channel = await EnsureChannel(client, teamId, channelName, log);
                    await AddExistingGuestsToTeam(client, members, log);
                    await UploadFileAndPostMessage(client, teamId, channel, filePath, log);
                    await ListTeamMembersWithRoles(client, log);

                    await log.WriteLineAsync("=== Toate operațiunile au fost finalizate ===");
                }

                Console.WriteLine("Operațiuni finalizate cu succes! Verifică fișierul automation.log pentru detalii.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Eroare generală: {ex.Message}");
                using (var log = new StreamWriter(logFilePath, append: true, Encoding.UTF8))
                {
                    await log.WriteLineAsync($"EROARE CRITICĂ [{DateTime.Now:yyyy-MM-dd HH:mm:ss}]: {ex.Message}");
                    await log.WriteLineAsync($"Stack trace: {ex.StackTrace}");
                }
            }

            Console.WriteLine("Apasă orice tastă pentru a închide...");
            Console.ReadKey();
        }

        static GraphServiceClient GetGraphClient()
        {
            if (string.IsNullOrEmpty(accessToken))
                throw new InvalidOperationException("Te rog setează access token-ul în variabila de mediu 'AZURE_TOKEN'!");

            var credential = new HardcodedTokenCredential(accessToken);
            return new GraphServiceClient(credential);
        }

        // --- Restul codului rămâne la fel ---
        static (string channelName, List<string> members, string filePath) ReadInputFile(string path)
        {
            var lines = File.ReadAllLines(path, Encoding.UTF8)
                            .Where(l => !string.IsNullOrWhiteSpace(l))
                            .Select(l => l.Trim())
                            .ToList();

            if (lines.Count < 3)
                throw new InvalidOperationException("Fișierul de input trebuie să aibă cel puțin 3 linii: nume canal, cel puțin un membru, calea către fișier!");

            string channelName = lines[0];
            string filePath = lines.Last();
            var members = lines.Skip(1).Take(lines.Count - 2).ToList();

            return (channelName, members, filePath);
        }

        static void CreateSampleInputFile()
        {
            var sampleContent = @"Test Canal Automatizare
guest1@example.com
guest2@example.com
C:\temp\test_file.txt";

            File.WriteAllText(inputFilePath, sampleContent, Encoding.UTF8);
        }

        static async Task<Channel> EnsureChannel(GraphServiceClient client, string teamId, string channelName, StreamWriter log)
        {
            Console.WriteLine($"Verifică existența canalului '{channelName}'...");

            var channels = await client.Teams[teamId].Channels.GetAsync();
            var existingChannel = channels?.Value?.FirstOrDefault(c =>
                string.Equals(c.DisplayName, channelName, StringComparison.OrdinalIgnoreCase));

            if (existingChannel != null)
            {
                Console.WriteLine($"✓ Canalul '{channelName}' există deja.");
                await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Canalul '{channelName}' există deja (ID: {existingChannel.Id}).");
                return existingChannel;
            }

            Console.WriteLine($"Creează canalul '{channelName}'...");
            var channel = new Channel
            {
                DisplayName = channelName,
                Description = $"Canal creat automat la {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                MembershipType = ChannelMembershipType.Standard
            };

            var createdChannel = await client.Teams[teamId].Channels.PostAsync(channel);

            Console.WriteLine($"✓ Canalul '{channelName}' a fost creat cu succes!");
            await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Canalul '{channelName}' a fost creat (ID: {createdChannel?.Id}).");

            return createdChannel;
        }

        static async Task ListTeamMembersWithRoles(GraphServiceClient client, StreamWriter log)
        {
            Console.WriteLine("Se extrag membrii și rolurile lor din team...");

            var members = await client.Teams[teamId].Members.GetAsync();

            if (members?.Value == null || members.Value.Count == 0)
            {
                Console.WriteLine("⚠️ Nu există membri în acest team!");
                await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Nu există membri în team.");
                return;
            }

            await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Lista membrilor și rolurile lor în team:");

            foreach (var m in members.Value.OfType<AadUserConversationMember>())
            {
                string displayName = m.DisplayName ?? "Necunoscut";
                string email = m.Email ?? "N/A";
                string roles = (m.Roles != null && m.Roles.Count > 0) ? string.Join(",", m.Roles) : "member";

                Console.WriteLine($" - {displayName} ({email}) | Roluri: {roles}");
                await log.WriteLineAsync($"   - {displayName} ({email}) | Roluri: {roles}");
            }
        }

        static async Task AddExistingGuestsToTeam(GraphServiceClient client, List<string> guests, StreamWriter log)
        {
            Console.WriteLine($"Adaugă direct {guests.Count} guest-uri existente în team...");
            int successCount = 0, errorCount = 0;

            foreach (var email in guests)
            {
                try
                {
                    var users = await client.Users.GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Filter = $"userPrincipalName eq '{email}'";
                    });

                    var guestUser = users?.Value?.FirstOrDefault();
                    if (guestUser == null)
                    {
                        Console.WriteLine($"⚠️ Guest-ul {email} nu există în tenant!");
                        await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Guest-ul {email} nu există în tenant.");
                        errorCount++;
                        continue;
                    }

                    var member = new AadUserConversationMember
                    {
                        Roles = new List<string> { "guest" },
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{guestUser.Id}')" }
                        }
                    };
                    await client.Teams[teamId].Members.PostAsync(member);

                    Console.WriteLine($"✓ Guest-ul {email} a fost adăugat direct în team!");
                    await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Guest-ul {email} a fost adăugat direct în team.");
                    successCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"✗ Eroare la adăugarea guest-ului {email}: {ex.Message}");
                    await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] EROARE: {ex.Message}");
                    errorCount++;
                }
            }

            Console.WriteLine($"Guest-uri adăugate: {successCount} succese, {errorCount} erori");
            await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Sumar adăugări guest-uri: {successCount} succese, {errorCount} erori");
        }

        static async Task UploadFileAndPostMessage(GraphServiceClient client, string teamId, Channel channel, string filePath, StreamWriter log)
        {
            try
            {
                Console.WriteLine($"Încarcă fișierul '{Path.GetFileName(filePath)}' în canalul '{channel.DisplayName}'...");
                var drive = await client.Groups[teamId].Drive.GetAsync();
                if (drive?.Id == null)
                    throw new Exception("Nu s-a putut accesa drive-ul asociat team-ului.");

                var fileName = Path.GetFileName(filePath);
                var uploadPath = $"{channel.DisplayName}/{fileName}";

                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    await client.Drives[drive.Id].Root.ItemWithPath(uploadPath).Content.PutAsync(fileStream);
                }

                var uploadedItem = await client.Drives[drive.Id].Root.ItemWithPath(uploadPath).GetAsync();
                string fileUrl = uploadedItem?.WebUrl;

                Console.WriteLine($"✓ Fișier încărcat cu succes! URL: {fileUrl}");
                await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Fișier încărcat: {fileName} în canalul '{channel.DisplayName}'");

                var message = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = $"Am adăugat fișierul <a href='{fileUrl}'>{fileName}</a> în canal."
                    },
                    Attachments = new List<ChatMessageAttachment>
                    {
                        new ChatMessageAttachment
                        {
                            ContentType = "reference",
                            ContentUrl = fileUrl,
                            Name = fileName
                        }
                    }
                };

                await client.Teams[teamId].Channels[channel.Id].Messages.PostAsync(message);
                Console.WriteLine("✓ Mesaj postat în Teams cu fișierul atașat!");
                await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] Mesaj postat în Teams cu fișierul {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Eroare la încărcarea fișierului sau postarea mesajului: {ex.Message}");
                await log.WriteLineAsync($"[{DateTime.Now:HH:mm:ss}] EROARE: {ex.Message}");
                throw;
            }
        }
    }

    public class HardcodedTokenCredential : TokenCredential
    {
        private readonly string _token;
        public HardcodedTokenCredential(string token)
        {
            if (string.IsNullOrEmpty(token))
                throw new ArgumentException("Token-ul nu poate fi null sau gol!", nameof(token));
            _token = token;
        }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            return new AccessToken(_token, DateTimeOffset.UtcNow.AddHours(1));
        }

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            return new ValueTask<AccessToken>(GetToken(requestContext, cancellationToken));
        }
    }
}
