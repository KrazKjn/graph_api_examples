// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me.SendMail;

class GraphHelper
{
    // <UserAuthConfigSnippet>
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public enum ItemType
    {
        Audio,
        Bundle,
        File,
        Folder,
        Image,
        Photo,
        Video
    }

    public class WIN32_FIND_DATA
    {
        public Int32 FileAttributes;
        public DateTimeOffset CreationTime;
        public DateTimeOffset LastAccessTime;
        public DateTimeOffset LastWriteTime;
        public Int64 FileSize;
        public String FileName = "";
        public String AlternateFileName = "";
    };

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        var options = new DeviceCodeCredentialOptions
        {
            ClientId = settings.ClientId,
            TenantId = settings.TenantId,
            DeviceCodeCallback = deviceCodePrompt,
        };

        _deviceCodeCredential = new DeviceCodeCredential(options);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }
    // </UserAuthConfigSnippet>

    // <GetUserTokenSnippet>
    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }
    // </GetUserTokenSnippet>

    // <GetUserSnippet>
    public static Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] {"displayName", "mail", "userPrincipalName" };
        });
    }
    // </GetUserSnippet>

    // <GetInboxSnippet>
    public static Task<MessageCollectionResponse?> GetInboxAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            // Only messages from Inbox folder
            .MailFolders["Inbox"]
            .Messages
            .GetAsync((config) =>
            {
                // Only request specific properties
                config.QueryParameters.Select = new[] { "from", "isRead", "receivedDateTime", "subject" };
                // Get at most 25 results
                config.QueryParameters.Top = 25;
                // Sort by received time, newest first
                config.QueryParameters.Orderby = new[] { "receivedDateTime DESC" };
            });
    }
    // </GetInboxSnippet>

    // <SendMailSnippet>
    public static async Task SendMailAsync(string subject, string body, string recipient)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Create a new message
        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = BodyType.Text
            },
            ToRecipients = new List<Recipient>
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient
                    }
                }
            }
        };

        // Send the message
        await _userClient.Me
            .SendMail
            .PostAsync(new SendMailPostRequestBody
            {
                Message = message
            });
    }
    // </SendMailSnippet>

    #pragma warning disable CS1998
    // <GetDrivesAsync>
    public async static Task<DriveCollectionResponse?> GetDrivesAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return await _userClient.Me.Drives.GetAsync();
    }
    // </GetDrivesAsync>

    // <GetDriveItemsAsync>
    public async static Task<DriveItem?> GetDriveItemsAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        //return await _userClient.Me.Drive.GetAsync();

        var rootDrive = await _userClient.Me.Drives.GetAsync();
        _ = rootDrive ?? throw new System.NullReferenceException("No drives found");
        _ = rootDrive.Value ?? throw new System.NullReferenceException("No drives found");
        var rootItems = await _userClient.Drives[rootDrive.Value[0].Id].Root.GetAsync((conf) =>
        {
            conf.QueryParameters.Expand = new string[] { "children" };
        });
        return rootItems;
    }
    // </GetDriveItemsAsync>

    public static async Task<IList<(String name, String id, Int64 size, String description, ItemType itemType)>> GetAllFilesFromOne()
    {
        IList<(String name, String id, Int64 size, String description, ItemType itemType)> results = new List<(String name, String id, Int64 size, String description, ItemType itemType)>();
        try
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var drives = await _userClient.Me.Drives.GetAsync();
            _ = drives ?? throw new System.NullReferenceException("No drives found");
            _ = drives.Value ?? throw new System.NullReferenceException("No drives found");
            await AddFilesFromRoot(drives.Value[0], results);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error on List Files: {ex.Message}");
        }
        return results;
    }

    // Get the Folders and Files on 
    private static async Task AddFilesFromRoot(Microsoft.Graph.Models.Drive rootDrive, IList<(string name, string id, Int64 size, String description, ItemType itemType)> results)
    {
        // On this level is is required to use a filter to avoid error 'The 'filter' query option must be provided.'
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        var rootItems = await _userClient.Drives[rootDrive.Id].Root.GetAsync((conf) =>
        {
            conf.QueryParameters.Expand = new string[] { "children" };
        });
        _ = rootItems ?? throw new System.NullReferenceException("No items found");
        _ = rootItems.Children ?? throw new System.NullReferenceException("No items found");

        foreach (var child in rootItems.Children)
        {
            if (child.Folder != null)
            {
                // Add recursive the files from a subfolder
                //results.Add(($"{child.Name}", child.Id, (child.Size.HasValue ? child.Size.Value : 0), child.Description, ItemType.Folder));
                Console.WriteLine($"Item: {child.Name} (ID: {child.Id}): 0, {child.Description}, {ItemType.Folder}");
                await AddFolderFromOne(rootDrive, results, $"{child.Name}");
            }
            else
            {
                ItemType it = ItemType.File;
                if (child.Audio != null) it = ItemType.Audio;
                else if (child.Bundle != null) it = ItemType.Bundle;
                else if (child.Image != null) it = ItemType.Image;
                else if (child.Photo != null) it = ItemType.Photo;
                else if (child.Video != null) it = ItemType.Video;  
                //results.Add(($"{child.Name}", child.Id, (child.Size.HasValue ? child.Size.Value : 0), child.Description, it));
                Console.WriteLine($"Item: {child.Name} (ID: {child.Id}): {child.Size}, {child.Description}, {it}");
            }
        }
    }

    // Add recursive the files from a subfolder 
    private static async Task AddFolderFromOne(Microsoft.Graph.Models.Drive rootDrive, IList<(string name, string id, Int64 size, String description, ItemType itemType)> results, String itemPath)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        var children = _userClient.Drives[rootDrive.Id].Root.ItemWithPath(itemPath);
        var childs = await children.Children.GetAsync();
        _ = childs ?? throw new System.NullReferenceException("No items found");
        _ = childs.Value ?? throw new System.NullReferenceException("No items found");

        if (childs != null)
        {
            foreach (var child in childs.Value)
            {
                if (child.Folder != null)
                {
                    //results.Add(($"{itemPath}/{child.Name}", child.Id, (child.Size.HasValue ? child.Size.Value : 0), child.Description, ItemType.Folder));
                    Console.WriteLine($"Item: {itemPath}{child.Name} (ID: {child.Id}): 0, {child.Description}, {ItemType.Folder}");
                    await AddFolderFromOne(rootDrive, results, $"{itemPath}/{child.Name}");
                }
                else
                {
                    ItemType it = ItemType.File;
                    if (child.Audio != null) it = ItemType.Audio;
                    else if (child.Bundle != null) it = ItemType.Bundle;
                    else if (child.Image != null) it = ItemType.Image;
                    else if (child.Photo != null) it = ItemType.Photo;
                    else if (child.Video != null) it = ItemType.Video;  
                    //results.Add(($"{itemPath}/{child.Name}", child.Id, (child.Size.HasValue ? child.Size.Value : 0), child.Description, it));
                    Console.WriteLine($"Item: {itemPath}{child.Name} (ID: {child.Id}): {child.Size}, {child.Description}, {it}");
                }
            }
        }
    }

    public static async Task<WIN32_FIND_DATA> FindFirstFile(string fileName)
    {
        WIN32_FIND_DATA lpFindFileData = new();

        try
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var drives = await _userClient.Me.Drives.GetAsync();
            _ = drives ?? throw new System.NullReferenceException("No drives found");
            _ = drives.Value ?? throw new System.NullReferenceException("No drives found");

            // Get the file metadata
            var fileItem = await _userClient.Drives[drives.Value[0].Id].Root.ItemWithPath(fileName).GetAsync();
            _ = fileItem ?? throw new System.NullReferenceException("Item not found");

            Console.WriteLine($"File Name: {fileItem.Name}");
            Console.WriteLine($"File ID: {fileItem.Id}");
            lpFindFileData = new() 
            {
                FileName = fileItem.Name ?? "",
                FileSize = fileItem.Size ?? 0,
                AlternateFileName = fileItem.Name ?? "",
                //FileAttributes = fileItem.File.  .FileSystemInfo?.Attributes ?? 0,
                CreationTime = fileItem.FileSystemInfo?.CreatedDateTime ?? fileItem.CreatedDateTime ?? DateTime.MinValue,
                LastAccessTime = fileItem.FileSystemInfo?.LastAccessedDateTime ?? fileItem.LastModifiedDateTime ?? DateTime.MinValue,
                LastWriteTime = fileItem.FileSystemInfo?.LastModifiedDateTime ?? fileItem.LastModifiedDateTime ?? DateTime.MinValue
            };
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        return lpFindFileData;
    }

    public static async Task<String> GetFile(string fileName, string directory, bool overWrite = false)
    {
        string localFileName = "";
        try
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var drives = await _userClient.Me.Drives.GetAsync();
            _ = drives ?? throw new System.NullReferenceException("No drives found");
            _ = drives.Value ?? throw new System.NullReferenceException("No drives found");

            // Get the file metadata
            var fileItem = await _userClient.Drives[drives.Value[0].Id].Root.ItemWithPath(fileName).GetAsync();
            _ = fileItem ?? throw new System.NullReferenceException("Item not found");
            Console.WriteLine($"File Name: {fileItem.Name}");
            Console.WriteLine($"File ID: {fileItem.Id}");

            // Download the file content
            localFileName = $"{directory}\\{fileItem.Name}";
            var fileStream = await _userClient.Drives[drives.Value[0].Id].Root.ItemWithPath(fileName).Content.GetAsync();
            _ = fileStream ?? throw new System.NullReferenceException("Item not found");
            FileInfo fi = new(localFileName);
            if (fi.Exists)
            {
                if (overWrite)
                {
                    fi.Delete();
                }
                else
                {
                    Console.WriteLine($"File {localFileName} already exists.");
                    return localFileName;
                }
            }
            using (var file = new FileStream(localFileName, FileMode.Create, FileAccess.Write))
            {
                await fileStream.CopyToAsync(file);
            }

            FileInfo file1 = new(localFileName)
            {
                CreationTime = (fileItem.FileSystemInfo?.CreatedDateTime ?? fileItem.CreatedDateTime ?? DateTime.MinValue).UtcDateTime,
                LastAccessTime = (fileItem.FileSystemInfo?.LastAccessedDateTime ?? fileItem.LastModifiedDateTime ?? DateTime.MinValue).UtcDateTime,
                LastWriteTime = (fileItem.FileSystemInfo?.LastModifiedDateTime ?? fileItem.LastModifiedDateTime ?? DateTime.MinValue).UtcDateTime
            };

            Console.WriteLine("File downloaded successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        return localFileName;
    }

    public static async Task<String> SaveFile(string localFileName, string directory)
    {
        string remoteFileName = "";
        try
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var drives = await _userClient.Me.Drives.GetAsync();
            _ = drives ?? throw new System.NullReferenceException("No drives found");
            _ = drives.Value ?? throw new System.NullReferenceException("No drives found");

            FileInfo fi = new(localFileName);
            remoteFileName = $"{directory}\\{fi.Name}";

            // Upload a file
            using (var uploadStream = new FileStream(localFileName, FileMode.Open))
            {
                var uploadedItem = await _userClient.Drives[drives.Value[0].Id].Root.ItemWithPath(remoteFileName).Content.PutAsync(uploadStream);
                _ = uploadedItem ?? throw new System.NullReferenceException("Item not found");
                Console.WriteLine($"Uploaded File ID: {uploadedItem.Id}");
                uploadedItem.CreatedDateTime = fi.CreationTime;
                uploadedItem.LastModifiedDateTime = fi.LastWriteTime;
                uploadedItem.FileSystemInfo ??= new();
                uploadedItem.FileSystemInfo.CreatedDateTime = fi.CreationTime;
                uploadedItem.FileSystemInfo.LastAccessedDateTime = fi.LastAccessTime;
            }                        

            Console.WriteLine("File Saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        return remoteFileName;
    }

    public static async Task<bool> FileExists(string fileName)
    {
        try
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var drives = await _userClient.Me.Drives.GetAsync();
            _ = drives ?? throw new System.NullReferenceException("No drives found");
            _ = drives.Value ?? throw new System.NullReferenceException("No drives found");

            // Get the file metadata
            var driveItem = await _userClient.Drives[drives.Value[0].Id].Root.ItemWithPath(fileName).GetAsync();

            return driveItem != null;
        }
        //catch (ServiceException ex) when (ex.GetBaseException.StatusCode == System.Net.HttpStatusCode.NotFound)
        catch (Exception)
        {
            return false;
        }
    }
    public static async Task<bool> PromptUser(string question) 
    {
        Console.WriteLine($"{question} (Yes/No)");
        string? response = Console.ReadLine();
        if (response == null)
            return false;

        response = response.Trim().ToLower();
        return (response == "yes" || response == "y");
    }
}
