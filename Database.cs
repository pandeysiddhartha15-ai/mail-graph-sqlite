using System;
using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;

public class Database : IDisposable
{
    private readonly string _connectionString;
    private readonly SqlConnection _conn;
    private readonly ILogger<Database>? _logger;

    // Accept a full SQL Server connection string (e.g. "Server=...;Database=...;User Id=...;Password=...;")
    public Database(string connectionString, ILogger<Database>? logger = null)
    {
        if (string.IsNullOrEmpty(connectionString)) throw new ArgumentNullException(nameof(connectionString));
        _connectionString = connectionString;
        _conn = new SqlConnection(_connectionString);
        _conn.Open();
        _logger = logger;
    }

    public void EnsureSchema()
    {
        // Create Emails and Metadata tables if not exist (SQL Server compatible)
        var createSql = @"
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Emails]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.Emails (
        Id NVARCHAR(50) PRIMARY KEY,
        MessageId NVARCHAR(200) UNIQUE,
        InternetMessageId NVARCHAR(500),
        Subject NVARCHAR(1000),
        BodyHtml NVARCHAR(MAX),
        BodyText NVARCHAR(MAX),
        SentTo NVARCHAR(MAX),
        SentFrom NVARCHAR(500),
        Cc NVARCHAR(MAX),
        Bcc NVARCHAR(MAX),
        SentDateTimeUtc DATETIME2,
        ReceivedDateTimeUtc DATETIME2,
        RawJson NVARCHAR(MAX)
    );
END

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Metadata]') AND type in (N'U'))
BEGIN
    CREATE TABLE dbo.Metadata (
        [Key] NVARCHAR(200) PRIMARY KEY,
        [Value] NVARCHAR(MAX)
    );
END

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Emails_ReceivedDateTimeUtc' AND object_id = OBJECT_ID('dbo.Emails'))
BEGIN
    CREATE INDEX IX_Emails_ReceivedDateTimeUtc ON dbo.Emails (ReceivedDateTimeUtc);
END
";
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = createSql;
        cmd.CommandType = CommandType.Text;
        cmd.ExecuteNonQuery();

        _logger?.LogInformation("Database schema ensured (Server={Server};Database={Database})", _conn.DataSource, _conn.Database);
    }

    public void InsertEmail(EmailRecord r)
    {
        // Idempotent insert: only insert if a row with same MessageId does not exist.
        var sql = @"
INSERT INTO dbo.Emails (Id, MessageId, InternetMessageId, Subject, BodyHtml, BodyText, SentTo, SentFrom, Cc, Bcc, SentDateTimeUtc, ReceivedDateTimeUtc, RawJson)
SELECT @Id, @MessageId, @InternetMessageId, @Subject, @BodyHtml, @BodyText, @SentTo, @SentFrom, @Cc, @Bcc, @SentDateTimeUtc, @ReceivedDateTimeUtc, @RawJson
WHERE NOT EXISTS (SELECT 1 FROM dbo.Emails WHERE MessageId = @MessageId);
";
        using var tran = _conn.BeginTransaction();
        using var cmd = _conn.CreateCommand();
        cmd.Transaction = tran;
        cmd.CommandText = sql;
        cmd.Parameters.AddWithValue("@Id", (object)r.Id ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@MessageId", (object)r.MessageId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@InternetMessageId", (object)r.InternetMessageId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Subject", (object)r.Subject ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@BodyHtml", (object)r.BodyHtml ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@BodyText", (object)r.BodyText ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SentTo", (object)r.SentTo ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SentFrom", (object)r.SentFrom ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Cc", (object)r.Cc ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Bcc", (object)r.Bcc ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SentDateTimeUtc", (object?) (r.SentDateTimeUtc.HasValue ? (object)r.SentDateTimeUtc.Value : DBNull.Value) ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@ReceivedDateTimeUtc", (object?) (r.ReceivedDateTimeUtc.HasValue ? (object)r.ReceivedDateTimeUtc.Value : DBNull.Value) ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@RawJson", (object)r.RawJson ?? DBNull.Value);

        cmd.ExecuteNonQuery();
        tran.Commit();

        _logger?.LogDebug("InsertEmail executed for MessageId={MessageId}", r.MessageId);
    }

    public string? GetMetadata(string key)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = "SELECT [Value] FROM dbo.Metadata WHERE [Key] = @Key";
        cmd.Parameters.AddWithValue("@Key", key);
        using var rdr = cmd.ExecuteReader();
        if (rdr.Read())
        {
            return rdr.IsDBNull(0) ? null : rdr.GetString(0);
        }
        return null;
    }

    public void SetMetadata(string key, string value)
    {
        // Upsert: update if exists else insert
        using var tran = _conn.BeginTransaction();
        using (var cmd = _conn.CreateCommand())
        {
            cmd.Transaction = tran;
            cmd.CommandText = @"
IF EXISTS (SELECT 1 FROM dbo.Metadata WHERE [Key] = @Key)
    UPDATE dbo.Metadata SET [Value] = @Value WHERE [Key] = @Key;
ELSE
    INSERT INTO dbo.Metadata ([Key], [Value]) VALUES (@Key, @Value);
";
            cmd.Parameters.AddWithValue("@Key", key);
            cmd.Parameters.AddWithValue("@Value", value ?? (object)DBNull.Value);
            cmd.ExecuteNonQuery();
        }
        tran.Commit();
        _logger?.LogDebug("Metadata set {Key}={Value}", key, value);
    }

    public void Dispose()
    {
        _conn?.Dispose();
    }
}