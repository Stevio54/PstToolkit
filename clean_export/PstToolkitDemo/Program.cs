using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Linq;
using PstToolkit;
using PstToolkit.Exceptions;
using PstToolkit.Formats;
using static PstToolkit.FilterOperator;

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
                        
                    case "listmessages":
                        if (args.Length < 3)
                        {
                            Console.WriteLine("Error: PST file path and folder name are required for listmessages command.");
                            ShowUsage();
                            return;
                        }
                        ListMessages(args[1], args[2]);
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
                        
                    case "filteredcopy":
                        if (args.Length < 4)
                        {
                            Console.WriteLine("Error: Source PST, destination PST, and filter parameters are required for filteredcopy command.");
                            ShowUsage();
                            return;
                        }
                        CopyPstWithFilter(args[1], args[2], args[3]);
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
            Console.WriteLine("  PstToolkitDemo listmessages <pst_file> <folder> - List all messages in a folder");
            Console.WriteLine("  PstToolkitDemo copy <src> <dst>                - Copy messages from source PST to destination PST");
            Console.WriteLine("  PstToolkitDemo filteredcopy <src> <dst> <filter> - Copy filtered messages from source PST to destination PST");
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
            
            try
            {
                // Open source PST in read-only mode
                using var sourcePst = PstFile.Open(sourcePath, readOnly: true);
                
                // Handle existing destination file
                if (File.Exists(destPath))
                {
                    // If file exists, delete it (for simplicity)
                    // In a real application, you might want to ask for confirmation
                    Console.WriteLine($"Destination file already exists, removing: {destPath}");
                    File.Delete(destPath);
                }
                
                // Create a new destination PST
                Console.WriteLine($"Creating new destination file: {destPath}");
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
            catch (Exception ex)
            {
                Console.WriteLine($"Error during copy: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Details: {ex.InnerException.Message}");
                }
            }
        }
        
        static void CopyPstWithFilter(string sourcePath, string destPath, string filterSpec)
        {
            Console.WriteLine($"Copying PST file with filter: {sourcePath} -> {destPath}");
            Console.WriteLine($"Filter criteria: {filterSpec}");
            
            try
            {
                // Open source PST in read-only mode
                using var sourcePst = PstFile.Open(sourcePath, readOnly: true);
                
                // Handle existing destination file
                if (File.Exists(destPath))
                {
                    Console.WriteLine($"Destination file already exists, removing: {destPath}");
                    File.Delete(destPath);
                }
                
                // Create a new destination PST
                Console.WriteLine($"Creating new destination file: {destPath}");
                using var destPst = PstFile.Create(destPath);
                
                // Create a message filter based on the filter specification
                var filter = ParseFilterSpec(filterSpec);
                if (filter == null)
                {
                    Console.WriteLine("Error: Invalid filter specification. Using default (copy all).");
                    destPst.CopyFrom(sourcePst);
                }
                else
                {
                    Console.WriteLine("Copying folders and filtered messages...");
                    
                    // Count messages matching the filter criteria
                    int totalMessages = CountMessages(sourcePst.RootFolder);
                    int matchingMessages = 0;
                    CountMatchingMessages(sourcePst.RootFolder, filter, ref matchingMessages);
                    
                    Console.WriteLine($"Total messages in source: {totalMessages}");
                    Console.WriteLine($"Messages matching filter: {matchingMessages}");
                    
                    if (matchingMessages == 0)
                    {
                        Console.WriteLine("Warning: No messages match the filter criteria.");
                        Console.Write("Do you want to continue with the copy operation? (y/n): ");
                        var response = Console.ReadLine()?.ToLower();
                        if (response != "y" && response != "yes")
                        {
                            Console.WriteLine("Operation canceled by user.");
                            return;
                        }
                    }
                    
                    // Using stopwatch to track performance
                    var stopwatch = new System.Diagnostics.Stopwatch();
                    stopwatch.Start();
                    
                    // Set up progress reporting
                    int lastPercent = -1;
                    destPst.CopyFrom(sourcePst, filter, progress => {
                        int percent = (int)(progress * 100);
                        if (percent > lastPercent)
                        {
                            Console.Write($"\rProgress: {percent}%");
                            lastPercent = percent;
                        }
                    });
                    
                    stopwatch.Stop();
                    Console.WriteLine("\rProgress: 100%");
                    
                    // Count messages in destination
                    int copiedMessages = CountMessages(destPst.RootFolder);
                    
                    // Calculate statistics
                    double timeInSeconds = stopwatch.ElapsedMilliseconds / 1000.0;
                    double messagesPerSecond = copiedMessages / timeInSeconds;
                    
                    Console.WriteLine($"Filtered copy completed successfully in {timeInSeconds:F2} seconds!");
                    Console.WriteLine($"Copied {copiedMessages} of {matchingMessages} matching messages ({messagesPerSecond:F2} messages/sec)");
                    
                    // If there's a difference, explain why
                    if (copiedMessages != matchingMessages)
                    {
                        Console.WriteLine($"Note: {matchingMessages - copiedMessages} messages were skipped due to errors or incompatibility.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during filtered copy: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Details: {ex.InnerException.Message}");
                }
            }
        }
        
        static MessageFilter? ParseFilterSpec(string filterSpec)
        {
            // Format examples:
            // subject:contains:Project Update - Filter messages with subject containing "Project Update" 
            // sender:equals:john@example.com - Filter messages with sender email exactly matching
            // date:after:2025-04-25 - Filter messages sent after 2025-04-25
            // Multiple filter formats:
            // AND - subject:contains:Important AND date:after:2025-04-25
            // OR - subject:contains:Report OR subject:contains:Summary
            
            try
            {
                // Check for multiple conditions with AND/OR
                if (filterSpec.Contains(" AND "))
                {
                    // Split by AND and create a filter with All (AND) logic
                    var conditions = filterSpec.Split(new[] { " AND " }, StringSplitOptions.None);
                    var filter = new MessageFilter().SetLogic(FilterLogic.All);
                    
                    // Add each condition directly to the filter
                    foreach (var condition in conditions)
                    {
                        var subFilter = ParseSingleFilterCondition(condition);
                        if (subFilter == null)
                        {
                            Console.WriteLine($"Warning: Invalid filter condition: {condition}");
                            return null;
                        }
                        
                        // Add the single filter condition directly
                        var parts = condition.Split(':');
                        if (parts.Length < 2)
                        {
                            Console.WriteLine($"Error: Filter must be in format 'property:operator:value': {condition}");
                            return null;
                        }
                        
                        string property = parts[0].ToLowerInvariant();
                        string op = parts.Length > 1 ? parts[1].ToLowerInvariant() : "contains";
                        string value = parts.Length > 2 ? string.Join(":", parts.Skip(2)) : "";
                        
                        // Map the operator string to FilterOperator enum
                        FilterOperator filterOp = GetOperator(op);
                        
                        // Special handling for dates
                        if (property == "date" || property == "sentdate" || property == "receiveddate")
                        {
                            if (DateTime.TryParse(value, out DateTime dateValue))
                            {
                                filter.AddCondition(property, filterOp, dateValue);
                            }
                            else
                            {
                                filter.AddCondition(property, filterOp, value);
                            }
                        }
                        else
                        {
                            filter.AddCondition(property, filterOp, value);
                        }
                    }
                    
                    return filter;
                }
                else if (filterSpec.Contains(" OR "))
                {
                    // Split by OR and create a filter with Any (OR) logic
                    var conditions = filterSpec.Split(new[] { " OR " }, StringSplitOptions.None);
                    var filter = new MessageFilter().SetLogic(FilterLogic.Any);
                    
                    // Add each condition directly to the filter
                    foreach (var condition in conditions)
                    {
                        var subFilter = ParseSingleFilterCondition(condition);
                        if (subFilter == null)
                        {
                            Console.WriteLine($"Warning: Invalid filter condition: {condition}");
                            return null;
                        }
                        
                        // Add the single filter condition directly
                        var parts = condition.Split(':');
                        if (parts.Length < 2)
                        {
                            Console.WriteLine($"Error: Filter must be in format 'property:operator:value': {condition}");
                            return null;
                        }
                        
                        string property = parts[0].ToLowerInvariant();
                        string op = parts.Length > 1 ? parts[1].ToLowerInvariant() : "contains";
                        string value = parts.Length > 2 ? string.Join(":", parts.Skip(2)) : "";
                        
                        // Map the operator string to FilterOperator enum
                        FilterOperator filterOp = GetOperator(op);
                        
                        // Special handling for dates
                        if (property == "date" || property == "sentdate" || property == "receiveddate")
                        {
                            if (DateTime.TryParse(value, out DateTime dateValue))
                            {
                                filter.AddCondition(property, filterOp, dateValue);
                            }
                            else
                            {
                                filter.AddCondition(property, filterOp, value);
                            }
                        }
                        else
                        {
                            filter.AddCondition(property, filterOp, value);
                        }
                    }
                    
                    return filter;
                }
                else
                {
                    // Single condition
                    return ParseSingleFilterCondition(filterSpec);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing filter: {ex.Message}");
                return null;
            }
        }
        
        // Helper method to get FilterOperator from string
        private static FilterOperator GetOperator(string op)
        {
            switch (op.ToLowerInvariant())
            {
                case "contains":
                    return FilterOperator.Contains;
                case "equals":
                case "is":
                    return FilterOperator.Equals;
                case "startswith":
                case "starts":
                    return FilterOperator.StartsWith;
                case "endswith":
                case "ends":
                    return FilterOperator.EndsWith;
                case "after":
                case "greaterthan":
                case "newer":
                    return FilterOperator.GreaterThan;
                case "before":
                case "lessthan":
                case "older":
                    return FilterOperator.LessThan;
                case "between":
                    return FilterOperator.Between;
                case "regex":
                case "matches":
                    return FilterOperator.RegexMatch;
                default:
                    return FilterOperator.Contains; // Default
            }
        }
        
        static MessageFilter? ParseSingleFilterCondition(string filterSpec)
        {
            try
            {
                var parts = filterSpec.Split(':');
                if (parts.Length < 2)
                {
                    Console.WriteLine("Error: Filter must be in format 'property:operator:value'");
                    return null;
                }
                
                string property = parts[0].ToLowerInvariant();
                string op = parts.Length > 1 ? parts[1].ToLowerInvariant() : "contains";
                string value = parts.Length > 2 ? string.Join(":", parts.Skip(2)) : "";
                
                var filter = new MessageFilter();
                
                // Map the operator string to FilterOperator enum
                FilterOperator filterOp = FilterOperator.Contains; // Default
                switch (op)
                {
                    case "contains":
                        filterOp = FilterOperator.Contains;
                        break;
                    case "equals":
                    case "is":
                        filterOp = FilterOperator.Equals;
                        break;
                    case "startswith":
                    case "starts":
                        filterOp = FilterOperator.StartsWith;
                        break;
                    case "endswith":
                    case "ends":
                        filterOp = FilterOperator.EndsWith;
                        break;
                    case "after":
                    case "greaterthan":
                    case "newer":
                        filterOp = FilterOperator.GreaterThan;
                        // For dates, try to parse the value
                        if (property == "date" || property == "sentdate" || property == "receiveddate")
                        {
                            if (DateTime.TryParse(value, out DateTime dateValue))
                            {
                                value = dateValue.ToString("o"); // Use ISO 8601 format
                            }
                        }
                        break;
                    case "before":
                    case "lessthan":
                    case "older":
                        filterOp = FilterOperator.LessThan;
                        // For dates, try to parse the value
                        if (property == "date" || property == "sentdate" || property == "receiveddate")
                        {
                            if (DateTime.TryParse(value, out DateTime dateValue))
                            {
                                value = dateValue.ToString("o"); // Use ISO 8601 format
                            }
                        }
                        break;
                    case "between":
                        filterOp = FilterOperator.Between;
                        // Extract range values
                        var rangeValues = value.Split('-');
                        if (rangeValues.Length >= 2)
                        {
                            value = rangeValues[0].Trim();
                            string value2 = rangeValues[1].Trim();
                            
                            // For dates, try to parse the values
                            if (property == "date" || property == "sentdate" || property == "receiveddate")
                            {
                                if (DateTime.TryParse(value, out DateTime dateValue1) && 
                                    DateTime.TryParse(value2, out DateTime dateValue2))
                                {
                                    filter.AddCondition(property, filterOp, dateValue1, dateValue2);
                                    return filter;
                                }
                            }
                            else
                            {
                                filter.AddCondition(property, filterOp, value, value2);
                                return filter;
                            }
                        }
                        break;
                    case "regex":
                    case "matches":
                        filterOp = FilterOperator.RegexMatch;
                        break;
                }
                
                // Add the condition to the filter
                filter.AddCondition(property, filterOp, value);
                return filter;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing filter: {ex.Message}");
                return null;
            }
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
        
        /// <summary>
        /// Counts messages in a folder and its subfolders that match the given filter.
        /// </summary>
        /// <param name="folder">The folder to check.</param>
        /// <param name="filter">The message filter to apply.</param>
        /// <param name="matchCount">Reference to a counter for matching messages.</param>
        static void CountMatchingMessages(PstFolder folder, MessageFilter? filter, ref int matchCount)
        {
            if (filter == null) 
            {
                // If no filter, all messages match
                matchCount += folder.MessageCount;
            }
            else
            {
                // Apply filter to messages in this folder
                foreach (var message in folder.Messages)
                {
                    if (filter.Matches(message))
                    {
                        matchCount++;
                    }
                }
            }
            
            // Recursively count in subfolders
            foreach (var subFolder in folder.SubFolders)
            {
                CountMatchingMessages(subFolder, filter, ref matchCount);
            }
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
                Console.WriteLine($"Sender: {message.SenderName} <{message.SenderEmail}>");
                Console.WriteLine($"Date: {message.SentDate}");
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
        
        static void ListMessages(string pstFilePath, string folderName)
        {
            Console.WriteLine($"Listing messages in folder '{folderName}' of PST file: {pstFilePath}");
            
            try
            {
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
                
                Console.WriteLine($"Found {messages.Count} messages in folder '{folder.Name}'");
                Console.WriteLine();
                
                // Display a table header for messages
                Console.WriteLine("ID".PadRight(10) + "Subject".PadRight(50) + "Sender".PadRight(30) + "Date".PadRight(25) + "Size");
                Console.WriteLine("".PadRight(120, '-'));
                
                // List each message with its metadata
                foreach (var message in messages)
                {
                    string id = $"{message.MessageId:X}".PadRight(10);
                    string subject = (message.Subject.Length > 47 ? message.Subject.Substring(0, 47) + "..." : message.Subject).PadRight(50);
                    string sender = (message.SenderName.Length > 27 ? message.SenderName.Substring(0, 27) + "..." : message.SenderName).PadRight(30);
                    string date = message.SentDate.ToString("g").PadRight(25);
                    string size = $"{message.Size:N0} bytes";
                    
                    Console.WriteLine($"{id}{subject}{sender}{date}{size}");
                    
                    // Display additional details for the message
                    if (message.HasAttachments)
                    {
                        Console.WriteLine($"  Attachments: {string.Join(", ", message.AttachmentNames)}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        
        static PstFolder? FindFolderByPath(PstFolder rootFolder, string path)
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
                    Console.WriteLine($"Folder segment not found: {segment}");
                    return null;
                }
                
                currentFolder = nextFolder;
            }
            
            return currentFolder; // This is guaranteed to be non-null here
        }
    }
}
