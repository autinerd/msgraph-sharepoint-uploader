﻿using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace SharePointInterface
{
	internal static class Program
	{
		static void Main(string[] args)
		{
			if (!new string[] { "--upload", "--download", "--move", "--delete" }.Contains(args[0]))
			{
				Console.Write(@$"
Usage: SharePointInterface <action> <parameters>

SharePointInterface --upload ""source folder name"" ""file name pattern"" ""SharePoint site name"" ""SharePoint destination folder path""
SharePointInterface --download ""SharePoint site name"" ""SharePoint source file path"" ""Destination file path""
SharePointInterface --move ""SharePoint site name"" ""SharePoint source file path"" ""SharePoint destination folder path""
SharePointInterface --delete ""SharePoint site name"" ""SharePoint file path""
");
				Environment.Exit(0);
			}

			try
			{
				RunAsync(args).GetAwaiter().GetResult();
			}
			catch (Exception ex)
			{
				Console.ForegroundColor = ConsoleColor.Red;
				Console.WriteLine(ex.Message);
				Console.WriteLine(ex.StackTrace);
				Console.ResetColor();
			}
		}

		private static GraphServiceClient InitializeGraph(Settings settings) => GraphHelper.InitializeGraphForAppOnlyAuth(settings);

		private static async Task RunAsync(string[] args)
		{
			var settings = Settings.LoadSettings();

			var graphServiceClient = InitializeGraph(settings);

			switch (args[0])
			{
				case "--upload":
					{
						var sourceFolder = args[1];
						var filePattern = args[2];
						var siteName = args[3];
						var siteFolderPath = args[4];
						IEnumerable<Site>? sites = (await graphServiceClient.Sites.GetAsync((config) => config.QueryParameters.Search = $"\"{siteName}\""))?.Value?.Where((site) => site.Name == siteName);
						if (sites == null)
						{
							Logger.Log($"ERROR: No site with name \"{siteName}\" found.");
							Environment.Exit(1);
						}
						var siteId = sites.First().Id;
						if (siteId == null)
						{
							Logger.Log($"ERROR: Site ID is null! Search query = {siteName}");
							Environment.Exit(1);
						}

						await uploadFiles(graphServiceClient, sourceFolder, filePattern, siteId, siteFolderPath);
						break;
					}
				case "--download":
					{
						var siteName = args[1];
						var siteFilePath = args[2];
						var destinationPath = args[3];
						IEnumerable<Site>? sites = (await graphServiceClient.Sites.GetAsync((config) => config.QueryParameters.Search = $"\"{siteName}\""))?.Value?.Where((site) => site.Name == siteName);
						if (sites == null)
						{
							Logger.Log($"ERROR: No site with name \"{siteName}\" found.");
							Environment.Exit(1);
						}
						var siteId = sites.First().Id;
						if (siteId == null)
						{
							Logger.Log($"ERROR: Site ID is null! Search query = {siteName}");
							Environment.Exit(1);
						}

						await downloadFile(graphServiceClient, siteId, siteFilePath, destinationPath);
						break;
					}
				case "--move":
					{
						var siteName = args[1];
						var siteFilePath = args[2];
						var destinationPath = args[3];
						IEnumerable<Site>? sites = (await graphServiceClient.Sites.GetAsync((config) => config.QueryParameters.Search = $"\"{siteName}\""))?.Value?.Where((site) => site.Name == siteName);
						if (sites == null)
						{
							Logger.Log($"ERROR: No site with name \"{siteName}\" found.");
							Environment.Exit(1);
						}
						var siteId = sites.First().Id;
						if (siteId == null)
						{
							Logger.Log($"ERROR: Site ID is null! Search query = {siteName}");
							Environment.Exit(1);
						}

						await moveFile(graphServiceClient, siteId, siteFilePath, destinationPath);
						break;
					}
				case "--delete":
					{
						var siteName = args[1];
						var siteFilePath = args[2];
						IEnumerable<Site>? sites = (await graphServiceClient.Sites.GetAsync((config) => config.QueryParameters.Search = $"\"{siteName}\""))?.Value?.Where((site) => site.Name == siteName);
						if (sites == null)
						{
							Logger.Log($"ERROR: No site with name \"{siteName}\" found.");
							Environment.Exit(1);
						}
						var siteId = sites.First().Id;
						if (siteId == null)
						{
							Logger.Log($"ERROR: Site ID is null! Search query = {siteName}");
							Environment.Exit(1);
						}

						await deleteFile(graphServiceClient, siteId, siteFilePath);
						break;
					}
				default:
					break;
			}
		}

		private static async Task uploadFiles(GraphServiceClient client, string sourceFolder, string filePattern, string siteId, string siteFolderPath)
		{

			foreach (var fi in new DirectoryInfo(sourceFolder).GetFiles(filePattern))
			{
				using var fileStream = fi.OpenRead();

				var drive = await client.Sites[siteId].Drive.GetAsync();
				if (drive == null)
				{
					Logger.Log($"ERROR: Drive of site {siteId} could not be found!");
					Environment.Exit(1);
				}

				var uploadSession = await client.Drives[drive.Id].Items["root"].ItemWithPath(siteFolderPath + "/" + fi.Name).CreateUploadSession.PostAsync(new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody
				{
					Item = new Microsoft.Graph.Models.DriveItemUploadableProperties
					{
						AdditionalData = new Dictionary<string, object> {
							{ "@microsoft.graph.conflictBehavior", "replace" }
						}
					}
				});

				try
				{
					var uploadResult = await new LargeFileUploadTask<DriveItem>(uploadSession, fileStream).UploadAsync();
					Logger.Log(uploadResult.UploadSucceeded ?
									  $"SUCCESS: Upload file {fi.FullName} to site {siteId} path {siteFolderPath + "/" + fi.Name} complete" :
									  $"ERROR: Upload file {fi.FullName} to site {siteId} path {siteFolderPath + "/" + fi.Name} failed");
				}
				catch (ODataError ex)
				{
					Logger.Log($"ERROR: Error uploading: {ex.ToString()}");
				}
			}
		}

		private static async Task downloadFile(GraphServiceClient client, string siteId, string siteFilePath, string destinationPath)
		{
			var fileInfo = new FileInfo(destinationPath);
			using var writeStream = fileInfo.Create();

			var drive = await client.Sites[siteId].Drive.GetAsync();
			if (drive == null)
			{
				Logger.Log($"ERROR: Drive of site {siteId} could not be found!");
				Environment.Exit(1);
			}

			var content = await client.Drives[drive.Id].Items["root"].ItemWithPath(siteFilePath).Content.GetAsync();

			if (content == null)
			{
				Logger.Log($"ERROR: File {siteFilePath} of site {siteId} could not be found!");
				Environment.Exit(1);
			}

			content.CopyTo(writeStream);
			Logger.Log(@$"SUCCESS: Downloaded file ""{siteFilePath}"" to ""{destinationPath}""");
		}

		private static async Task moveFile(GraphServiceClient client, string siteId, string siteFilePath, string destinationFolder)
		{
			var drive = await client.Sites[siteId].Drive.GetAsync();
			if (drive == null)
			{
				Logger.Log($"ERROR: Drive of site {siteId} could not be found!");
				Environment.Exit(1);
			}

			var file = await client.Drives[drive.Id].Items["root"].ItemWithPath(siteFilePath).GetAsync();

			if (file == null)
			{
				Logger.Log($"ERROR: File {siteFilePath} of site {siteId} could not be found!");
				Environment.Exit(1);
			}
			DriveItem? destfolder;
			try
			{
				destfolder = await client.Drives[drive.Id].Items["root"].ItemWithPath(destinationFolder).GetAsync();
			}
			catch (ODataError e)
			{
				if (e.Error?.Code == "itemNotFound")
				{
					destfolder = null;
				}
				else
				{
					throw;
				}
			}


			if (destfolder == null)
			{
				var path = destinationFolder.Split('/');
				var parent = await client.Drives[drive.Id].Items["root"].GetAsync();

				for (int i = 0; i < path.Length; i++)
				{
					var data = await client.Drives[drive.Id].Items[parent?.Id].Children.GetAsync();
					var items = data?.Value?.Where((item) => item.Name == path[i]);
					if (items == null)
					{
						Logger.Log($"ERROR: Error on enumeration in subfolders!");
						Environment.Exit(1);
					}
					switch (items.Count())
					{
						case 0:
							parent = await client.Drives[drive.Id].Items[parent?.Id].Children.PostAsync(new DriveItem
							{
								Name = path[i],
								Folder = new Folder()
							});
							break;
						case 1:
							if (items.First().Folder == null)
							{
								Logger.Log($"ERROR: Item {string.Join('/', path[0..i])} on site {siteId} is not a folder!");
								Environment.Exit(1);
							}
							parent = await client.Drives[drive.Id].Items[items.First().Id].GetAsync();
							break;
						default:
							Logger.Log($"ERROR: Multiple items with name {path[i]} on site {siteId}!");
							Environment.Exit(1);
							break;
					}
				}
				destfolder = parent;
			}

			if (destfolder == null)
			{
				Logger.Log($"ERROR: Folder {destinationFolder} doesn't exist on site {siteId}!");
				Environment.Exit(1);
			}

			var result = await client.Drives[drive.Id].Items["root"].ItemWithPath(siteFilePath).PatchAsync(new DriveItem
			{
				ParentReference = new ItemReference
				{
					Id = destfolder.Id
				}
			});

			if (result == null)
			{
				Logger.Log($"ERROR: Error in moving the file on site {siteId}!");
				Environment.Exit(1);
			}
			else
			{
				Logger.Log($"SUCCESS: File {siteFilePath} moved to folder {destinationFolder} on site {siteId}!");
			}
		}

		private static async Task deleteFile(GraphServiceClient client, string siteId, string filePath)
		{
			var drive = await client.Sites[siteId].Drive.GetAsync();
			if (drive == null)
			{
				Logger.Log($"ERROR: Drive of site {siteId} could not be found!");
				Environment.Exit(1);
			}

			var file = await client.Drives[drive.Id].Items["root"].ItemWithPath(filePath).GetAsync();

			if (file == null)
			{
				Logger.Log($"INFO: File {filePath} of site {siteId} is already deleted!");
				Environment.Exit(0);
			}
			else
			{
				await client.Drives[drive.Id].Items["root"].ItemWithPath(filePath).DeleteAsync();
				Logger.Log($"SUCCESS: File {filePath} of site {siteId} successfully deleted!");
			}
		}
	}
}
