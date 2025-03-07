// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// <ProgramSnippet>
Console.WriteLine(".NET Graph Tutorial\n");

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

// Greet the user by name
await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. List my inbox");
    Console.WriteLine("3. Send mail");
    Console.WriteLine("4. List OneDrive Drives");
    Console.WriteLine("5. List OneDrive Contents");
    Console.WriteLine("6. Get OneDrive File using FindFirstFileAsync ('OneDriveTest.txt')");
    Console.WriteLine("7. Get OneDrive Contents using GetFile ('OneDriveTest.txt')");
    Console.WriteLine("8. Save a file to OneDrive ...");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch(choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // List emails from user's inbox
            await ListInboxAsync();
            break;
        case 3:
            // Send an email message
            await SendMailAsync();
            break;
        case 4:
            // List OneDrive drives
            await ListDrivesAsync();
            break;
        case 5:
            // List OneDrive contents
            await ListDriveContentsAsync();
            await ListDriveContentsAsync2();
            break;
        case 6:
            // FindFirstFileAsync
            await FindFirstFileAsync();
            break;
        case 7:
            // GetFileAsync
            await GetFileAsync();
            break;
        case 8:
            // SaveFileAsync
            await SaveFileAsync();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}
// </ProgramSnippet>

// <InitializeGraphSnippet>
void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForUserAuth(settings,
        (info, cancel) =>
        {
            // Display the device code message to
            // the user. This tells them
            // where to go to sign in and provides the
            // code to use.
            Console.WriteLine(info.Message);
            return Task.FromResult(0);
        });
}
// </InitializeGraphSnippet>

// <GreetUserSnippet>
async Task GreetUserAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}
// </GreetUserSnippet>

// <DisplayAccessTokenSnippet>
async Task DisplayAccessTokenAsync()
{
    try
    {
        var userToken = await GraphHelper.GetUserTokenAsync();
        Console.WriteLine($"User token: {userToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user access token: {ex.Message}");
    }
}
// </DisplayAccessTokenSnippet>

// <ListInboxSnippet>
async Task ListInboxAsync()
{
    try
    {
        var messagePage = await GraphHelper.GetInboxAsync();

        if (messagePage?.Value == null)
        {
            Console.WriteLine("No results returned.");
            return;
        }

        // Output each message's details
        foreach (var message in messagePage.Value)
        {
            Console.WriteLine($"Message: {message.Subject ?? "NO SUBJECT"}");
            Console.WriteLine($"  From: {message.From?.EmailAddress?.Name}");
            Console.WriteLine($"  Status: {(message.IsRead!.Value ? "Read" : "Unread")}");
            Console.WriteLine($"  Received: {message.ReceivedDateTime?.ToLocalTime().ToString()}");
        }

        // If NextPageRequest is not null, there are more messages
        // available on the server
        // Access the next page like:
        // var nextPageRequest = new MessagesRequestBuilder(messagePage.OdataNextLink, _userClient.RequestAdapter);
        // var nextPage = await nextPageRequest.GetAsync();
        var moreAvailable = !string.IsNullOrEmpty(messagePage.OdataNextLink);

        Console.WriteLine($"\nMore messages available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user's inbox: {ex.Message}");
    }
}
// </ListInboxSnippet>

// <SendMailSnippet>
async Task SendMailAsync()
{
    try
    {
        // Send mail to the signed-in user
        // Get the user for their email address
        var user = await GraphHelper.GetUserAsync();

        var userEmail = user?.Mail ?? user?.UserPrincipalName;

        if (string.IsNullOrEmpty(userEmail))
        {
            Console.WriteLine("Couldn't get your email address, canceling...");
            return;
        }

        await GraphHelper.SendMailAsync("Testing Microsoft Graph",
            "Hello world!", userEmail);

        Console.WriteLine("Mail sent.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error sending mail: {ex.Message}");
    }
}
// </SendMailSnippet>

// <ListDrivesAsync>
async Task ListDrivesAsync()
{
    try
    {
        var drives = await GraphHelper.GetDrivesAsync();

        if (drives == null || drives.Value == null || drives.Value.Count == 0)
        {
            Console.WriteLine("No Drives found in your OneDrive.");
            return;
        }

        foreach (var drive in drives.Value)
        {
            Console.WriteLine($"Item: {drive.Name} (ID: {drive.Id})");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error listing OneDrive derives: {ex.Message}");
    }
}
// </ListDrivesAsync>

// <ListDriveContentsSnippet>
async Task ListDriveContentsAsync()
{
    try
    {
        var driveItems = await GraphHelper.GetDriveItemsAsync();

        // if (driveItems == null || driveItems.Items == null || driveItems.Items.Count == 0)
        // {
        //     Console.WriteLine("No items found in your OneDrive.");
        //     return;
        // }

        // foreach (var item in driveItems.Items)
        // {
        //     Console.WriteLine($"Item: {item.Name} (ID: {item.Id})");
        // }
        if (driveItems == null || driveItems.Children == null || driveItems.Children.Count == 0)
        {
            Console.WriteLine("No items found in your OneDrive.");
            return;
        }

        foreach (var item in driveItems.Children)
        {
            Console.WriteLine($"Item: {item.Name} (ID: {item.Id})");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error listing OneDrive contents: {ex.Message}");
    }
}
// </ListDriveContentsSnippet>

// <ListDriveContentsSnippet>
async Task ListDriveContentsAsync2()
{
    try
    {
        var driveItems = await GraphHelper.GetAllFilesFromOne();

        if (driveItems == null || driveItems.Count == 0)
        {
            Console.WriteLine("No items found in your OneDrive.");
            return;
        }

        foreach (var item in driveItems)
        {
            Console.WriteLine($"Item: {item.name} (ID: {item.id}): {item.size}, {item.description}, {item.itemType}");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error listing OneDrive contents: {ex.Message}");
    }
}
// </ListDriveContentsSnippet>

// <FindFirstFileSnippet>
async Task FindFirstFileAsync()
{
    try
    {
        var driveItem = await GraphHelper.FindFirstFile("OneDriveTest.txt");

        if (driveItem == null)
        {
            Console.WriteLine("FindFirstFile: OneDrive item Not found.");
            return;
        }

        Console.WriteLine($"Item: {driveItem.FileName} (ID: {driveItem.FileAttributes}): {driveItem.FileSize}, {driveItem.CreationTime}, {driveItem.LastWriteTime}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"FindFirstFile Error: OneDrive contents: {ex.Message}");
    }
}
// </ListDriveContentsSnippet>

// <GetFileSnippet>
async Task GetFileAsync()
{
    try
    {
        var localFileName = await GraphHelper.GetFile("OneDriveTest.txt", "C:\\Temp");

        Console.WriteLine($"Downloaded to: {localFileName}.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"FindFirstFile Error: OneDrive contents: {ex.Message}");
    }
}
// </GetFileSnippet>

// <SaveFileSnippet>
async Task SaveFileAsync()
{
    try
    {
        if (await GraphHelper.FileExists("\\Temp\\OneDriveTest.txt") && !await GraphHelper.PromptUser("File already exists in OneDrive, overwrite?"))
        {
            Console.WriteLine("File already exists in OneDrive, skipping...");
            return;
        }
        var remoteFileName = await GraphHelper.SaveFile("C:\\Temp\\OneDriveTest.txt", "\\Temp");

        Console.WriteLine($"Uploaded/Saved to: {remoteFileName}.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"SaveFile Error: {ex.Message}");
    }
}
// </SaveFileSnippet>