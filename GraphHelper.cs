using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
namespace SharePointInterface
{
	internal class GraphHelper
	{
		// Settings object
		private static Settings? _settings;
		// App-ony auth token credential
		private static ClientSecretCredential? _clientSecretCredential;
		// Client configured with app-only authentication
		private static GraphServiceClient? _appClient;

		public static GraphServiceClient InitializeGraphForAppOnlyAuth(Settings settings)
		{
			_settings = settings;

			// Ensure settings isn't null
			_ = settings ??
				throw new NullReferenceException("Settings cannot be null");

			_settings = settings;

			_clientSecretCredential ??= new ClientSecretCredential(
					_settings.TenantId, _settings.ClientId, _settings.ClientSecret);

			_appClient ??= new GraphServiceClient(_clientSecretCredential, ["https://graph.microsoft.com/.default"]);
			return _appClient;
		}

		public static async Task<string> GetAppOnlyTokenAsync()
		{
			// Ensure credential isn't null
			_ = _clientSecretCredential ??
				throw new NullReferenceException("Graph has not been initialized for app-only auth");

			// Request token with given scopes
			return (await _clientSecretCredential.GetTokenAsync(new TokenRequestContext(["https://graph.microsoft.com/.default"]))).Token;
		}
	}
}
