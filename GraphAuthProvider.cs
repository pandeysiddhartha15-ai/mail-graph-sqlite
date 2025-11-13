using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

public class GraphAuthProvider
{
    private readonly IConfidentialClientApplication _confidentialClient;
    private readonly string[] _scopes;

    public GraphAuthProvider(string tenantId, string clientId, string clientSecret, string[] scopes)
    {
        if (string.IsNullOrEmpty(tenantId)) throw new ArgumentNullException(nameof(tenantId));
        if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException(nameof(clientId));
        if (string.IsNullOrEmpty(clientSecret)) throw new ArgumentNullException(nameof(clientSecret));
        _scopes = scopes ?? throw new ArgumentNullException(nameof(scopes));

        _confidentialClient = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithTenantId(tenantId)
            .Build();
    }

    /// <summary>
    /// Acquires an app-only access token and applies it to the HttpRequestMessage.
    /// This method is used by the DelegateAuthenticationProvider when creating GraphServiceClient.
    /// </summary>
    public async Task AuthenticateRequestAsync(HttpRequestMessage request)
    {
        var result = await _confidentialClient.AcquireTokenForClient(_scopes).ExecuteAsync().ConfigureAwait(false);
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", result.AccessToken);
    }

    /// <summary>
    /// Overload with CancellationToken for compatibility with different SDK versions.
    /// </summary>
    public async Task AuthenticateRequestAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var result = await _confidentialClient.AcquireTokenForClient(_scopes).ExecuteAsync(cancellationToken).ConfigureAwait(false);
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", result.AccessToken);
    }

    /// <summary>
    /// Helper to retrieve raw access token if you need it elsewhere.
    /// </summary>
    public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default)
    {
        var result = await _confidentialClient.AcquireTokenForClient(_scopes).ExecuteAsync(cancellationToken).ConfigureAwait(false);
        return result.AccessToken;
    }
}