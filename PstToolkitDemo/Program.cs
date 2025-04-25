using System;
using System.Diagnostics;
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
            Console.WriteLine("  PstToolkitDemo info <pst_file>      - Show information about a PST file");
            Console.WriteLine("  PstToolkitDemo list <pst_file>      - List all folders in a PST file");
            Console.WriteLine("  PstToolkitDemo copy <src> <dst>     - Copy messages from source PST to destination PST");
            Console.WriteLine("  PstToolkitDemo create <pst_file>    - Create a new PST file");
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
    }
}
