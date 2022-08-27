using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace SharePointUploader
{
	internal static class Shared
	{
		internal static Microsoft.Identity.Client.IConfidentialClientApplication? app { get; set; }

		/// <summary>
		/// An example of how to authenticate the Microsoft Graph SDK using the MSAL library
		/// </summary>
		/// <returns></returns>
		internal static GraphServiceClient GetAuthenticatedGraphClient(IConfidentialClientApplication app, string[] scopes)
		{

			GraphServiceClient graphServiceClient =
					new GraphServiceClient("https://graph.microsoft.com/V1.0/", new DelegateAuthenticationProvider(async (requestMessage) =>
					{
						// Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
						AuthenticationResult result = await app.AcquireTokenForClient(scopes)
							.ExecuteAsync();

						// Add the access token in the Authorization header of the API request.
						requestMessage.Headers.Authorization =
							new AuthenticationHeaderValue("Bearer", result.AccessToken);
					}));

			return graphServiceClient;
		}

		/// <summary>
		/// Checks if the sample is configured for using ClientSecret or Certificate. This method is just for the sake of this sample.
		/// You won't need this verification in your production application since you will be authenticating in AAD using one mechanism only.
		/// </summary>
		/// <param name="config">Configuration from appsettings.json</param>
		/// <returns></returns>
		internal static bool IsAppUsingClientSecret(AuthenticationConfig config)
		{
			string clientSecretPlaceholderValue = "[Enter here a client secret for your application]";

			if (!String.IsNullOrWhiteSpace(config.ClientSecret) && config.ClientSecret != clientSecretPlaceholderValue)
			{
				return true;
			}

			else if (config.Certificate != null)
			{
				return false;
			}

			else
				throw new Exception("You must choose between using client secret or certificate. Please update appsettings.json file.");
		}
	}
}
