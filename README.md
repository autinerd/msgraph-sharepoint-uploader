# Simple SharePoint Uploader for MS Graph

With this simple console application, you can automate uploading files to SharePoint.

## Preparation

 Register your app in Azure AD, as described [here](https://github.com/Azure-Samples/active-directory-dotnetcore-daemon-v2/tree/master/1-Call-MSGraph#step-2--register-the-sample-with-your-azure-active-directory-tenant)

Instead of the **User.Read.All** permission, you need these:

- **Files.ReadWrite.All**
- **Sites.ReadWrite.All**

 Fill out the Tenant ID, Client ID and Client secret/certificate in the `appsettings.sample.json` and rename it to `appsettings.json`

## Usage

`SharePointUploader "source folder name" "file name pattern" "SharePoint site name" "SharePoint destination folder path"`

Example: Uploading a bunch of PDF files to a folder on the "Test" site.

`SharePointUploader "C:\Files\" "*.pdf" "Test" "\folder"`

Logs will be generated in the `client.log` file.
