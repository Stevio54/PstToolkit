# PstToolkit

A C# library for reading, creating, and transferring email messages between PST files without using Outlook Interop or paid libraries.

## Features

- Open and read existing PST files
- Create new PST files from scratch
- Extract messages to EML format
- Copy messages and folders between PST files
- Proper support for folder hierarchies and message metadata

## Usage Examples

### Reading a PST file

```csharp
using PstToolkit;

// Open an existing PST file in read-only mode
using var pst = PstFile.Open("example.pst", readOnly: true);

// Access the root folder
var rootFolder = pst.RootFolder;

// List folders and messages
Console.WriteLine($"Root folder name: {rootFolder.Name}");
foreach (var folder in rootFolder.SubFolders)
{
    Console.WriteLine($"Folder: {folder.Name} ({folder.MessageCount} messages)");
}
```

### Creating a new PST file

```csharp
using PstToolkit;

// Create a new PST file
using var pst = PstFile.Create("new.pst");

// Create folders
var inboxFolder = pst.RootFolder.CreateSubFolder("Inbox");
var testFolder = inboxFolder.CreateSubFolder("Test");

// Create a message
var message = PstMessage.Create(pst, 
    "Test Subject", 
    "This is the message body", 
    "sender@example.com", 
    "Sender Name");

// Add the message to a folder
testFolder.AddMessage(message);
```

### Copying between PST files

```csharp
using PstToolkit;

// Open source and destination PST files
using var sourcePst = PstFile.Open("source.pst", readOnly: true);
using var destPst = PstFile.Create("destination.pst");

// Copy all content from source to destination
destPst.CopyFrom(sourcePst);
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.