using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Identity.Web;

namespace SharePointUploader
{
	internal static class Program
	{
		static void Main(string[] args)
		{
			if (args.Length != 4)
			{
				Console.WriteLine(@"Usage: SharePointUploader ""source folder name"" ""file name pattern"" ""SharePoint site name"" ""SharePoint destination folder path"" ");
				Environment.Exit(0);
			}

			try
			{
				RunAsync(args[0], args[1], args[2], args[3]).GetAwaiter().GetResult();
			}
			catch (Exception ex)
			{
				Console.ForegroundColor = ConsoleColor.Red;
				Console.WriteLine(ex.Message);
				Console.ResetColor();
			}
		}

		private static async Task RunAsync(string sourceFolder, string filePattern, string siteName, string siteFolderPath)
		{
			AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");

			// You can run this sample using ClientSecret or Certificate. The code will differ only when instantiating the IConfidentialClientApplication
			bool isUsingClientSecret = Shared.IsAppUsingClientSecret(config);

			// Even if this is a console application here, a daemon application is a confidential client application

			if (isUsingClientSecret)
			{
				// Even if this is a console application here, a daemon application is a confidential client application
				Shared.app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
					.WithClientSecret(config.ClientSecret)
					.WithAuthority(new Uri(config.Authority))
					.Build();
			}

			else
			{
				ICertificateLoader certificateLoader = new DefaultCertificateLoader();
				certificateLoader.LoadIfNeeded(config.Certificate);

				Shared.app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
					.WithCertificate(config.Certificate.Certificate)
					.WithAuthority(new Uri(config.Authority))
					.Build();
			}

			Shared.app.AddInMemoryTokenCache();

			// With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
			// application permissions need to be set statically (in the portal or by PowerShell), and then granted by
			// a tenant administrator. 
			// Generates a scope -> "https://graph.microsoft.com/.default"

			// Call MS graph using the Graph SDK
			GraphServiceClient graphServiceClient = Shared.GetAuthenticatedGraphClient(Shared.app, new string[] { $"{config.ApiUrl}.default" });

			var sites = await graphServiceClient.Sites.Request(new Option[] { new QueryOption("search", siteName) }).GetAsync();
			if (sites.Count > 1)
			{
				Logger.Log($"Non-unique site name. Search query = {siteName}");
				Environment.Exit(1);
			}
			var siteId = sites[0].Id;

			await uploadFiles(graphServiceClient, sourceFolder, filePattern, siteId, siteFolderPath);
		}

		private static async Task uploadFiles(GraphServiceClient client, string sourceFolder, string filePattern, string siteId, string siteFolderPath)
		{

			foreach (var fi in new DirectoryInfo(sourceFolder).GetFiles(filePattern))
			{
				Logger.Log($"Upload file {fi.FullName} to site {siteId} path {siteFolderPath + "/" + fi.Name}");
				using var fileStream = fi.OpenRead();
				var uploadProps = new DriveItemUploadableProperties
				{
					AdditionalData = new Dictionary<string, object>
					{
						 { "@microsoft.graph.conflictBehavior", "replace" }
					}
				};

				var uploadSession = await client.Sites[siteId].Drive.Root.ItemWithPath(siteFolderPath + "/" + fi.Name)
				.CreateUploadSession(uploadProps).Request().PostAsync();

				try
				{
					var uploadResult = await new LargeFileUploadTask<DriveItem>(uploadSession, fileStream).UploadAsync();
					Logger.Log(uploadResult.UploadSucceeded ?
									  $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
									  "Upload failed");
				}
				catch (ServiceException ex)
				{
					Logger.Log($"Error uploading: {ex.ToString()}");
				}
			}
		}
	}
}
