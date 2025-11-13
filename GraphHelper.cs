using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using HtmlAgilityPack;
using Microsoft.Extensions.Logging;
using System.Text.Json;

public class GraphHelper
{
    private readonly GraphServiceClient _graph;
    private readonly Database _db;
    private readonly AppSettings _settings;
    private readonly ILogger<GraphHelper> _logger;
    private const int MaxRetries = 5;

    public GraphHelper(GraphServiceClient graph, Database db, AppSettings settings, ILogger<GraphHelper> logger)
    {
        _graph = graph ?? throw new ArgumentNullException(nameof(graph));
        _db = db ?? throw new ArgumentNullException(nameof(db));
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    // Create GraphServiceClient using an HttpClient with an auth handler that injects the Bearer token.
    public static GraphServiceClient CreateGraphClient(GraphAuthProvider authProvider)
    {
        if (authProvider == null) throw new ArgumentNullException(nameof(authProvider));

        var authHandler = new AuthHttpHandler(authProvider);
        var httpClient = new HttpClient(authHandler)
        {
            BaseAddress = new Uri("https://graph.microsoft.com/")
        };

        return new GraphServiceClient(httpClient);
    }

    public async Task FetchAndStoreEmailsAsync()
    {
        // Get last run time (UTC). Default to 7 days ago to avoid huge initial sync.
        var lastRunStr = _db.GetMetadata("last_run_utc");
        DateTime lastRunUtc;
        if (!string.IsNullOrEmpty(lastRunStr) && DateTime.TryParse(lastRunStr, out var parsed))
        {
            lastRunUtc = DateTime.SpecifyKind(parsed, DateTimeKind.Utc);
        }
        else
        {
            lastRunUtc = DateTime.UtcNow.AddDays(-7);
            _logger.LogInformation("No checkpoint found. Defaulting last_run_utc to {LastRun}", lastRunUtc);
        }

        var nowUtc = DateTime.UtcNow;

        _logger.LogInformation("Querying messages for {Mailbox} from {From} to {To}", _settings.MailboxUserPrincipalName, lastRunUtc.ToString("o"), nowUtc.ToString("o"));

        var selectFields = new[]
        {
            "id",
            "internetMessageId",
            "subject",
            "body",
            "bodyPreview",
            "sentDateTime",
            "receivedDateTime",
            "from",
            "toRecipients",
            "ccRecipients",
            "bccRecipients"
        };

        try
        {
            // Initial page request using v5 SDK pattern
            var filterValue = $"receivedDateTime ge {lastRunUtc.ToString("o")}";
            var page = await _graph.Users[_settings.MailboxUserPrincipalName].Messages.GetAsync(config =>
            {
                config.QueryParameters.Filter = filterValue;
                config.QueryParameters.Select = selectFields;
                config.QueryParameters.Top = _settings.PageSize;
                config.QueryParameters.Orderby = new[] { "receivedDateTime asc" };
            });

            // Loop through pages: use OdataNextLink (page.OdataNextLink) and create a new MessagesRequestBuilder with that URL
            while (page != null && page.Value != null)
            {
                foreach (var msg in page.Value.OfType<Message>())
                {
                    if (msg.ReceivedDateTime == null)
                    {
                        _logger.LogWarning("Skipping message with null ReceivedDateTime, id={Id}", msg?.Id);
                        continue;
                    }

                    var receivedUtc = msg.ReceivedDateTime.Value.UtcDateTime;
                    if (!(receivedUtc >= lastRunUtc && receivedUtc < nowUtc))
                    {
                        // skip outside the interval
                        _logger.LogDebug("Skipping message {Id} outside interval ({Received})", msg.Id, receivedUtc);
                        continue;
                    }

                    try
                    {
                        string rawJson = SerializeMessage(msg);

                        var record = new EmailRecord
                        {
                            Id = Guid.NewGuid().ToString(),
                            MessageId = msg.Id ?? "",
                            InternetMessageId = msg.InternetMessageId ?? "",
                            Subject = msg.Subject ?? "",
                            BodyHtml = msg.Body?.Content ?? "",
                            BodyText = ExtractPlainText(msg.Body?.Content ?? ""),
                            SentTo = RecipientsToString(msg.ToRecipients),
                            SentFrom = RecipientToString(msg.From),
                            Cc = RecipientsToString(msg.CcRecipients),
                            Bcc = RecipientsToString(msg.BccRecipients),
                            SentDateTimeUtc = msg.SentDateTime?.UtcDateTime,
                            ReceivedDateTimeUtc = msg.ReceivedDateTime?.UtcDateTime,
                            RawJson = rawJson
                        };

                        _db.InsertEmail(record);
                        _logger.LogInformation("Saved message {MessageId} received {ReceivedDate}", record.MessageId, record.ReceivedDateTimeUtc);
                    }
                    catch (Exception exMsg)
                    {
                        _logger.LogError(exMsg, "Failed saving message id {MessageId}", msg?.Id);
                        // continue processing others
                    }
                }

                // If there is a next link, prepare the next request using the SDK's MessagesRequestBuilder with the nextLink and the same RequestAdapter
                if (string.IsNullOrEmpty(page.OdataNextLink))
                {
                    break;
                }

                // Create a new request builder for the nextLink and fetch the next page
                var nextBuilder = new Microsoft.Graph.Users.Item.Messages.MessagesRequestBuilder(page.OdataNextLink, _graph.RequestAdapter);
                page = await nextBuilder.GetAsync();
            }

            // All good - update checkpoint to nowUtc
            _db.SetMetadata("last_run_utc", nowUtc.ToString("o"));
            _logger.LogInformation("Checkpoint updated to {Now}", nowUtc.ToString("o"));
        }
        catch (ServiceException sEx)
        {
            _logger.LogError(sEx, "Graph API error");
            throw;
        }
    }

    private static string RecipientsToString(IEnumerable<Recipient>? recipients)
    {
        if (recipients == null) return "";
        return string.Join(";", recipients.Select(r => RecipientToString(r)));
    }

    private static string RecipientToString(Recipient? recipient)
    {
        if (recipient == null) return "";
        return $"{recipient.EmailAddress?.Name ?? ""} <{recipient.EmailAddress?.Address ?? ""}>";
    }

    private static string ExtractPlainText(string html)
    {
        if (string.IsNullOrEmpty(html)) return "";

        try
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            var text = doc.DocumentNode.InnerText;
            var normalized = System.Text.RegularExpressions.Regex.Replace(text, @"\s{2,}", " ").Trim();
            return normalized;
        }
        catch
        {
            var noTags = System.Text.RegularExpressions.Regex.Replace(html, "<.*?>", string.Empty);
            return noTags;
        }
    }

    private string SerializeMessage(Message msg)
    {
        try
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = false,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
            };
            return JsonSerializer.Serialize(msg, options);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to serialize message {Id} to JSON; storing minimal metadata", msg?.Id);
            return $"{{\"id\":\"{msg?.Id}\",\"subject\":\"{msg?.Subject?.Replace("\"", "\\\"")}\"}}";
        }
    }

    private async Task<T> ExecuteWithRetryAsync<T>(Func<Task<T>> action)
    {
        int attempt = 0;
        while (true)
        {
            try
            {
                return await action().ConfigureAwait(false);
            }
            catch (ServiceException sEx) when (IsRetriable(sEx) && attempt < MaxRetries)
            {
                attempt++;
                var delaySeconds = GetRetryDelaySeconds(sEx, attempt);
                _logger.LogWarning("Transient Graph error (attempt {Attempt}). Waiting {Delay}s and retrying. Error: {Message}", attempt, delaySeconds, sEx.Message);
                await Task.Delay(TimeSpan.FromSeconds(delaySeconds)).ConfigureAwait(false);
            }
        }
    }

    // Robust retriable check that works across Graph SDK versions by using reflection to extract status code / headers
    private static bool IsRetriable(ServiceException ex)
    {
        try
        {
            var status = TryGetStatusCode(ex);
            if (status.HasValue)
            {
                var code = status.Value;
                if (code == 429 || code == 503 || code == 504) return true;
            }

            var headers = GetResponseHeaders(ex);
            if (headers != null)
            {
                if (headers.TryGetValue("Retry-After", out var retryVals) && retryVals.Any()) return true;
            }
        }
        catch
        {
            // if introspection fails, be conservative and return false
        }

        return false;
    }

    private static int GetRetryDelaySeconds(ServiceException ex, int attempt)
    {
        try
        {
            var headers = GetResponseHeaders(ex);
            if (headers != null && headers.TryGetValue("Retry-After", out var vals) && int.TryParse(vals.FirstOrDefault(), out var headerSeconds))
            {
                return headerSeconds;
            }
        }
        catch
        {
            // ignore and fallback
        }

        // Exponential backoff base
        return (int)Math.Pow(2, attempt) + 1;
    }

    // Try to obtain a numeric HTTP status code from ServiceException across SDK versions
    private static int? TryGetStatusCode(ServiceException ex)
    {
        if (ex == null) return null;
        var t = ex.GetType();

        // possible property names across SDK versions
        var candidates = new[] { "StatusCode", "RawStatusCode", "ResponseStatusCode", "HttpStatusCode" };

        foreach (var name in candidates)
        {
            var prop = t.GetProperty(name);
            if (prop == null) continue;
            var val = prop.GetValue(ex);
            if (val == null) continue;

            // common types: System.Net.HttpStatusCode or int
            if (val is HttpStatusCode sc) return (int)sc;
            if (val is int i) return i;
            if (val is long l) return (int)l;
        }

        return null;
    }

    // Try to extract response headers from ServiceException (may be IDictionary<string,IEnumerable<string>> or HttpResponseMessage in different SDKs)
    private static IDictionary<string, IEnumerable<string>>? GetResponseHeaders(ServiceException ex)
    {
        if (ex == null) return null;
        var t = ex.GetType();

        // Common property names
        var headerProps = new[] { "ResponseHeaders", "Response", "Headers" };

        foreach (var name in headerProps)
        {
            var prop = t.GetProperty(name);
            if (prop == null) continue;
            var val = prop.GetValue(ex);
            if (val == null) continue;

            // If the property is already IDictionary<string,IEnumerable<string>>
            if (val is IDictionary<string, IEnumerable<string>> dict) return dict;

            // If it's an HttpResponseMessage, extract headers
            if (val is HttpResponseMessage httpResp)
            {
                var dict2 = new Dictionary<string, IEnumerable<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (var h in httpResp.Headers) dict2[h.Key] = h.Value;
                if (httpResp.Content != null)
                {
                    foreach (var h in httpResp.Content.Headers) dict2[h.Key] = h.Value;
                }
                return dict2;
            }

            // Some SDK versions expose Headers as IEnumerable<KeyValuePair<string,string[]>> or similar
            // Try to coerce enumerable of pairs
            if (val is IEnumerable<KeyValuePair<string, IEnumerable<string>>> kvEnum)
            {
                return kvEnum.ToDictionary(k => k.Key, k => k.Value, StringComparer.OrdinalIgnoreCase);
            }

            if (val is IEnumerable<KeyValuePair<string, string[]>> kvEnum2)
            {
                return kvEnum2.ToDictionary(k => k.Key, k => (IEnumerable<string>)k.Value, StringComparer.OrdinalIgnoreCase);
            }
        }

        return null;
    }

    // DelegatingHandler that injects the Bearer token into outgoing HttpRequestMessage
    private class AuthHttpHandler : DelegatingHandler
    {
        private readonly GraphAuthProvider _authProvider;

        public AuthHttpHandler(GraphAuthProvider authProvider, HttpMessageHandler? innerHandler = null)
            : base(innerHandler ?? new HttpClientHandler())
        {
            _authProvider = authProvider ?? throw new ArgumentNullException(nameof(authProvider));
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            // Acquire token via GraphAuthProvider helper
            var token = await _authProvider.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            return await base.SendAsync(request, cancellationToken).ConfigureAwait(false);
        }
    }
}