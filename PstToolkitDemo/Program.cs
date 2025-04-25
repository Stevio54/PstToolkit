using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Linq;
using PstToolkit;
using PstToolkit.Exceptions;
using PstToolkit.Formats;

namespace PstToolkitDemo
{
    /// <summary>
    /// Demonstrates the usage of the PstToolkit library.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("PstToolkit Demo");
            Console.WriteLine("===============");
            
            if (args.Length < 1)
            {
                ShowUsage();
                return;
            }
            
            string command = args[0].ToLowerInvariant();
            
            try
            {
                switch (command)
                {
                    case "info":
                        if (args.Length < 2)
                        {
                            Console.WriteLine("Error: PST file path is required for info command.");
                            ShowUsage();
                            return;
                        }
                        ShowPstInfo(args[1]);
                        break;
                        
                    case "list":
                        if (args.Length < 2)
                        {
                            Console.WriteLine("Error: PST file path is required for list command.");
                            ShowUsage();
                            return;
                        }
                        ListFolders(args[1]);
                        break;
                        
                    case "copy":
                        if (args.Length < 3)
                        {
                            Console.WriteLine("Error: Source and destination PST file paths are required for copy command.");
                            ShowUsage();
                            return;
                        }
                        CopyPst(args[1], args[2]);
                        break;
                        
                    case "create":
                        if (args.Length < 2)
                        {
                            Console.WriteLine("Error: PST file path is required for create command.");
                            ShowUsage();
                            return;
                        }
                        CreatePst(args[1]);
                        break;
                        
                    case "message":
                        if (args.Length < 3)
                        {
                            Console.WriteLine("Error: PST file path and folder name are required for message command.");
                            ShowUsage();
                            return;
                        }
                        AddSampleMessage(args[1], args[2]);
                        break;
                        
                    case "extract":
                        if (args.Length < 4)
                        {
                            Console.WriteLine("Error: PST file path, folder name, and output directory are required for extract command.");
                            ShowUsage();
                            return;
                        }
                        ExtractMessages(args[1], args[2], args[3]);
                        break;
                        
                    case "createfolder":
                        if (args.Length < 3)
                        {
                            Console.WriteLine("Error: PST file path and folder path are required for createfolder command.");
                            ShowUsage();
                            return;
                        }
                        CreateFolder(args[1], args[2]);
                        break;
                        
                    default:
                        Console.WriteLine($"Error: Unknown command '{command}'.");
                        ShowUsage();
                        break;
                }
            }
            catch (PstException ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Details: {ex.InnerException.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
                Debug.WriteLine(ex.ToString());
            }
        }

        static void ShowUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  PstToolkitDemo info <pst_file>                 - Show information about a PST file");
            Console.WriteLine("  PstToolkitDemo list <pst_file>                 - List all folders in a PST file");
            Console.WriteLine("  PstToolkitDemo copy <src> <dst>                - Copy messages from source PST to destination PST");
            Console.WriteLine("  PstToolkitDemo create <pst_file>               - Create a new PST file");
            Console.WriteLine("  PstToolkitDemo message <pst_file> <folder>     - Add a sample message to the specified folder");
            Console.WriteLine("  PstToolkitDemo extract <pst_file> <folder> <output_dir> - Extract messages from a folder to .eml files");
            Console.WriteLine("  PstToolkitDemo createfolder <pst_file> <path>  - Create a folder path (e.g. 'Inbox/Subfolder')");
        }

        static void ShowPstInfo(string pstFilePath)
        {
            Console.WriteLine($"Analyzing PST file: {pstFilePath}");
            
            using var pst = PstFile.Open(pstFilePath, readOnly: true);
            
            Console.WriteLine($"Format: {(pst.IsAnsi ? "ANSI" : "Unicode")}");
            Console.WriteLine($"Root folder: {pst.RootFolder.Name}");
            
            int folderCount = CountFolders(pst.RootFolder);
            int messageCount = CountMessages(pst.RootFolder);
            
            Console.WriteLine($"Total folders: {folderCount}");
            Console.WriteLine($"Total messages: {messageCount}");
        }

        static void ListFolders(string pstFilePath)
        {
            Console.WriteLine($"Listing folders in PST file: {pstFilePath}");
            
            using var pst = PstFile.Open(pstFilePath, readOnly: true);
            
            ListFoldersRecursive(pst.RootFolder, 0);
        }

        static void ListFoldersRecursive(PstFolder folder, int level)
        {
            string indent = new string(' ', level * 2);
            Console.WriteLine($"{indent}- {folder.Name} ({folder.MessageCount} messages)");
            
            foreach (var subFolder in folder.SubFolders)
            {
                ListFoldersRecursive(subFolder, level + 1);
            }
        }

        static void CopyPst(string sourcePath, string destPath)
        {
            Console.WriteLine($"Copying PST file: {sourcePath} -> {destPath}");
            
            // Open source PST in read-only mode
            using var sourcePst = PstFile.Open(sourcePath, readOnly: true);
            
            // Create a new destination PST
            using var destPst = PstFile.Create(destPath);
            
            Console.WriteLine("Copying folders and messages...");
            
            // Set up progress reporting
            int lastPercent = -1;
            destPst.CopyFrom(sourcePst, progress => {
                int percent = (int)(progress * 100);
                if (percent > lastPercent)
                {
                    Console.Write($"\rProgress: {percent}%");
                    lastPercent = percent;
                }
            });
            
            Console.WriteLine("\rProgress: 100%");
            Console.WriteLine("Copy completed successfully!");
        }

        static void CreatePst(string pstFilePath)
        {
            Console.WriteLine($"Creating new PST file: {pstFilePath}");
            
            using var pst = PstFile.Create(pstFilePath, PstFormatType.Unicode);
            
            Console.WriteLine("New PST file created successfully!");
            Console.WriteLine($"Format: {(pst.IsAnsi ? "ANSI" : "Unicode")}");
            Console.WriteLine($"Root folder: {pst.RootFolder.Name}");
        }

        static int CountFolders(PstFolder folder)
        {
            int count = 1; // Count this folder
            
            foreach (var subFolder in folder.SubFolders)
            {
                count += CountFolders(subFolder);
            }
            
            return count;
        }

        static int CountMessages(PstFolder folder)
        {
            int count = folder.MessageCount;
            
            foreach (var subFolder in folder.SubFolders)
            {
                count += CountMessages(subFolder);
            }
            
            return count;
        }
        
        static void AddSampleMessage(string pstFilePath, string folderName)
        {
            Console.WriteLine($"Adding sample message to folder '{folderName}' in PST file: {pstFilePath}");
            
            try
            {
                // Open the PST file in read-write mode
                using var pst = PstFile.Open(pstFilePath, readOnly: false);
                
                // Find the folder
                var folder = FindFolderByPath(pst.RootFolder, folderName);
                if (folder == null)
                {
                    Console.WriteLine($"Error: Folder '{folderName}' not found");
                    return;
                }
                
                // Create a sample message
                var now = DateTime.Now;
                var subject = $"Sample Message {now:yyyy-MM-dd HH:mm:ss}";
                var body = "This is a sample message created by PstToolkitDemo.\r\n\r\n";
                body += $"Created on: {now}\r\n";
                body += "This message demonstrates the PstToolkit's ability to create and add messages to PST files.";
                
                var message = PstMessage.Create(pst, subject, body, 
                    "sender@example.com", "Sample Sender");
                
                // Add a recipient
                message.AddRecipient("Sample Recipient", "recipient@example.com");
                
                // Add an attachment
                var attachmentContent = Encoding.UTF8.GetBytes(
                    "This is a sample attachment created by PstToolkitDemo.");
                message.AddAttachment("sample.txt", attachmentContent, "text/plain");
                
                // Add the message to the folder
                folder.AddMessage(message);
                
                Console.WriteLine("Sample message created and added successfully");
                Console.WriteLine($"Subject: {subject}");
                Console.WriteLine($"Size: {message.Size} bytes");
                Console.WriteLine($"Has attachment: {message.HasAttachments}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        
        static void ExtractMessages(string pstFilePath, string folderName, string outputDir)
        {
            Console.WriteLine($"Extracting messages from folder '{folderName}' in PST file: {pstFilePath}");
            Console.WriteLine($"Output directory: {outputDir}");
            
            try
            {
                // Ensure the output directory exists
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                
                // Open the PST file
                using var pst = PstFile.Open(pstFilePath, readOnly: true);
                
                // Find the folder
                var folder = FindFolderByPath(pst.RootFolder, folderName);
                if (folder == null)
                {
                    Console.WriteLine($"Error: Folder '{folderName}' not found");
                    return;
                }
                
                // Get messages in the folder
                var messages = folder.Messages;
                if (messages.Count == 0)
                {
                    Console.WriteLine("No messages found in the specified folder");
                    return;
                }
                
                Console.WriteLine($"Found {messages.Count} messages to extract");
                
                // Extract each message to an EML file
                int successCount = 0;
                foreach (var message in messages)
                {
                    try
                    {
                        // Create a safe filename from the subject
                        var safeSubject = string.Join("_", message.Subject.Split(Path.GetInvalidFileNameChars()));
                        if (string.IsNullOrWhiteSpace(safeSubject))
                        {
                            safeSubject = $"Message_{message.MessageId:X}";
                        }
                        
                        // Create output filename
                        var outputFile = Path.Combine(outputDir, $"{safeSubject}.eml");
                        
                        // Get the raw content
                        var content = message.GetRawContent();
                        
                        // Write to file
                        File.WriteAllBytes(outputFile, content);
                        
                        Console.WriteLine($"Extracted: {Path.GetFileName(outputFile)}");
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error extracting message {message.MessageId:X}: {ex.Message}");
                    }
                }
                
                Console.WriteLine($"Successfully extracted {successCount} of {messages.Count} messages");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        
        static void CreateFolder(string pstFilePath, string folderPath)
        {
            Console.WriteLine($"Creating folder path '{folderPath}' in PST file: {pstFilePath}");
            
            try
            {
                // Open the PST file in read-write mode
                using var pst = PstFile.Open(pstFilePath, readOnly: false);
                
                // Split the path into segments
                var pathSegments = folderPath.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
                if (pathSegments.Length == 0)
                {
                    Console.WriteLine("Error: Invalid folder path");
                    return;
                }
                
                // Start at the root folder
                PstFolder currentFolder = pst.RootFolder;
                
                // Navigate/create each segment
                foreach (var segment in pathSegments)
                {
                    // Try to find an existing folder with this name
                    var nextFolder = currentFolder.FindFolder(segment);
                    
                    if (nextFolder == null)
                    {
                        // Create folder if it doesn't exist
                        Console.WriteLine($"Creating subfolder: {segment}");
                        nextFolder = currentFolder.CreateSubFolder(segment);
                    }
                    else
                    {
                        Console.WriteLine($"Found existing subfolder: {segment}");
                    }
                    
                    currentFolder = nextFolder;
                }
                
                Console.WriteLine($"Folder path '{folderPath}' created/verified successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        
        static PstFolder FindFolderByPath(PstFolder rootFolder, string path)
        {
            // Handle special case for root folder
            if (path == "/" || path == "\\" || string.IsNullOrWhiteSpace(path))
            {
                return rootFolder;
            }
            
            // Split the path into segments
            var pathSegments = path.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (pathSegments.Length == 0)
            {
                return rootFolder;
            }
            
            // Navigate through the path segments
            PstFolder currentFolder = rootFolder;
            foreach (var segment in pathSegments)
            {
                var nextFolder = currentFolder.FindFolder(segment);
                if (nextFolder == null)
                {
                    // Folder not found
                    return null;
                }
                
                currentFolder = nextFolder;
            }
            
            return currentFolder; // This is guaranteed to be non-null here
        }
    }
}
