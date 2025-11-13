using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Serilog;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // Load configuration
        var config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
            .AddEnvironmentVariables()
            .Build();

        var settings = config.GetSection("AppSettings").Get<AppSettings>();
        if (settings == null)
        {
            Console.Error.WriteLine("Missing AppSettings in configuration.");
            return 1;
        }

        // Read connection string (required for SQL Server)
        var connString = config.GetConnectionString("DefaultConnection");
        if (string.IsNullOrWhiteSpace(connString))
        {
            Console.Error.WriteLine("Missing ConnectionStrings:DefaultConnection in configuration. Please set your SQL Server connection string.");
            return 1;
        }

        // Configure Serilog
        var logFile = string.IsNullOrWhiteSpace(settings.LogFilePath) ? "logs/app.log" : settings.LogFilePath;
        var level = Serilog.Events.LogEventLevel.Information;
        if (!string.IsNullOrWhiteSpace(settings.LogLevel) &&
            Enum.TryParse<Serilog.Events.LogEventLevel>(settings.LogLevel, true, out var parsed))
        {
            level = parsed;
        }

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Is(level)
            .Enrich.FromLogContext()
            .WriteTo.Console()
            .WriteTo.File(logFile, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 14)
            .CreateLogger();

        using var loggerFactory = LoggerFactory.Create(builder =>
        {
            builder.AddSerilog(dispose: true);
        });

        var logger = loggerFactory.CreateLogger<Program>();
        try
        {
            logger.LogInformation("Application starting");

            // Ensure database connection to SQL Server
            using var db = new Database(connString, loggerFactory.CreateLogger<Database>());
            db.EnsureSchema();

            // Graph auth provider
            var authProvider = new GraphAuthProvider(settings.TenantId, settings.ClientId, settings.ClientSecret, settings.Scopes);

            // Create Graph client
            var graphClient = GraphHelper.CreateGraphClient(authProvider);

            // Graph helper with logging
            var helper = new GraphHelper(graphClient, db, settings, loggerFactory.CreateLogger<GraphHelper>());

            logger.LogInformation("Starting email fetch for {Mailbox}", settings.MailboxUserPrincipalName);
            await helper.FetchAndStoreEmailsAsync();
            logger.LogInformation("Email fetch finished");

            Log.CloseAndFlush();
            return 0;
        }
        catch (Exception ex)
        {
            logger.LogCritical(ex, "Fatal error");
            Log.CloseAndFlush();
            return 2;
        }
    }
}

// App settings POCO
public class AppSettings
{
    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ClientSecret { get; set; } = "";
    public string MailboxUserPrincipalName { get; set; } = "";
    public string DatabaseFilePath { get; set; } = "emails.db"; // kept for backward compatibility, not used when using SQL Server via ConnectionStrings
    public string[] Scopes { get; set; } = new[] { "https://graph.microsoft.com/.default" };
    public int PageSize { get; set; } = 50;
    // Logging
    public string LogFilePath { get; set; } = "logs/app.log";
    public string LogLevel { get; set; } = "Information";
}