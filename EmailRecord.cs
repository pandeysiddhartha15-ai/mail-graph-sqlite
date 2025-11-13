using System;

public class EmailRecord
{
    // Unique ID for our DB record
    public string Id { get; set; } = Guid.NewGuid().ToString();

    // Graph message id (unique)
    public string MessageId { get; set; } = "";

    // Internet message-id header value
    public string InternetMessageId { get; set; } = "";

    public string Subject { get; set; } = "";

    // HTML body
    public string BodyHtml { get; set; } = "";

    // Plain text body
    public string BodyText { get; set; } = "";

    public string SentTo { get; set; } = "";
    public string SentFrom { get; set; } = "";
    public string Cc { get; set; } = "";
    public string Bcc { get; set; } = "";

    public DateTime? SentDateTimeUtc { get; set; }
    public DateTime? ReceivedDateTimeUtc { get; set; }

    // Raw JSON of the Graph Message object for diagnostics / full capture
    public string RawJson { get; set; } = "";
}