```markdown
# mail-graph-sqlserver

C# Console app that reads a service mailbox via Microsoft Graph (app-only) and stores emails into Microsoft SQL Server (tested with SQL Server 2019).

Features
- App-only (client credentials) OAuth 2.0 via MSAL
- Reads messages for a specified mailbox and stores:
  - unique DB id, Graph message id, internet message id, subject
  - body HTML and plain-text extracted body
  - sentTo, sentFrom, cc, bcc, sent/received timestamps
  - raw Graph JSON of the message for diagnostics
- Checkpointing with last_run_utc to process only new messages (safe for scheduled Task Scheduler runs)
- Structured logging (Serilog) to console and rolling file
- Retry/backoff for transient Graph errors
- Idempotent inserts (avoids duplicates on overlapping runs) based on MessageId UNIQUE constraint

Important configuration
- Put your SQL Server connection string into ConnectionStrings:DefaultConnection in appsettings.json.
  Example:
  "ConnectionStrings": {
    "DefaultConnection": "Server=your-server;Database=MailDb;User Id=dbuser;Password=DB_PASSWORD;TrustServerCertificate=True;Encrypt=True;"
  }

- Make sure the DB exists (create the database MailDb) and the SQL user has CREATE TABLE or appropriate schema permissions for the first run, or create the tables manually.

Security
- Do not check secrets (client secret, DB password) into source control.
- For production, consider storing secrets in Azure Key Vault or using managed identities.

Build & Run
- Build:
  dotnet build
- Run:
  dotnet run
- Publish for Windows Task Scheduler:
  dotnet publish -c Release -r win-x64 --self-contained false -o publish

Task Scheduler notes
- Use an absolute connection string and ensure the scheduled task runs as a service account that can read the publish folder.
- The app uses a metadata key "last_run_utc" stored in the Metadata table in the database. Each run:
  - Reads last_run_utc (defaults to UTC now - 7 days on first run),
  - Queries messages receivedDateTime >= last_run_utc and < nowUtc,
  - Inserts messages idempotently,
  - Updates last_run_utc to nowUtc on successful completion.
```