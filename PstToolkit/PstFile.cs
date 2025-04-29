using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using PstToolkit.Exceptions;
using PstToolkit.Formats;
using PstToolkit.Utils;

namespace PstToolkit
{
    /// <summary>
    /// Represents a PST file that can be read from or written to.
    /// Includes optimizations for handling large PST files.
    /// </summary>
    public class PstFile : IDisposable
    {
        private FileStream? _fileStream;
        private bool _isReadOnly;
        private PstFormatHeader? _header;
        private Dictionary<uint, PstFolder>? _folderCache;
        private BTreeOnHeap? _nodeBTree;
        private BTreeOnHeap? _blockBTree;
        private PstFolder? _rootFolder;
        private bool _disposed;

        /// <summary>
        /// Gets the root folder of the PST file.
        /// </summary>
        public PstFolder RootFolder 
        { 
            get
            {
                if (_rootFolder == null)
                {
                    InitializeRootFolder();
                }
                return _rootFolder!;
            }
        }

        /// <summary>
        /// Gets whether the PST file is in ANSI format (as opposed to Unicode).
        /// </summary>
        public bool IsAnsi => _header?.IsAnsi ?? false;

        /// <summary>
        /// Gets the format type of the PST file.
        /// </summary>
        public PstFormatType FormatType => _header?.FormatType ?? PstFormatType.Unicode;

        /// <summary>
        /// Gets the file path of the PST file.
        /// </summary>
        public string FilePath { get; private set; } = string.Empty;

        /// <summary>
        /// Gets whether the PST file is open in read-only mode.
        /// </summary>
        public bool IsReadOnly => _isReadOnly;
        
        /// <summary>
        /// Gets the PST file header.
        /// </summary>
        internal PstFormatHeader Header => _header!;
        
        // GetNodeBTree exists at line 628, removed duplicate to fix build error
        
        /// <summary>
        /// Gets the file stream for the PST file.
        /// </summary>
        /// <returns>The FileStream of the PST file.</returns>
        internal FileStream GetFileStream()
        {
            if (_fileStream == null)
            {
                throw new PstAccessException("File stream is not available. The PST file may be closed.");
            }
            return _fileStream;
        }
        
        /// <summary>
        /// Gets the size of the PST file.
        /// </summary>
        /// <returns>The file size in bytes.</returns>
        internal ulong GetFileSize()
        {
            if (_fileStream == null)
                throw new PstException("File stream is not initialized");
                
            return (ulong)_fileStream.Length;
        }

        /// <summary>
        /// Opens an existing PST file for reading or writing.
        /// </summary>
        /// <param name="filePath">Path to the PST file to open.</param>
        /// <param name="readOnly">True to open in read-only mode, false for read-write.</param>
        /// <returns>A PstFile object for the opened file.</returns>
        /// <exception cref="PstException">Thrown when the file cannot be opened or is not a valid PST file.</exception>
        public static PstFile Open(string filePath, bool readOnly = true)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException($"PST file not found: {filePath}");
                }

                var fileMode = readOnly ? FileMode.Open : FileMode.Open;
                var fileAccess = readOnly ? FileAccess.Read : FileAccess.ReadWrite;
                var fileShare = readOnly ? FileShare.Read : FileShare.None;

                var fileStream = new FileStream(filePath, fileMode, fileAccess, fileShare);
                
                var pst = new PstFile
                {
                    _fileStream = fileStream,
                    _isReadOnly = readOnly,
                    FilePath = filePath,
                    _folderCache = new Dictionary<uint, PstFolder>()
                };

                pst.Initialize();
                return pst;
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstAccessException($"Failed to open PST file: {filePath}", ex);
            }
        }

        /// <summary>
        /// Creates a new PST file.
        /// </summary>
        /// <param name="filePath">Path where the new PST file should be created.</param>
        /// <param name="formatType">The PST format type to create (Unicode is recommended).</param>
        /// <returns>A PstFile object for the newly created file.</returns>
        /// <exception cref="PstException">Thrown when the file cannot be created.</exception>
        public static PstFile Create(string filePath, PstFormatType formatType = PstFormatType.Unicode)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    throw new IOException($"File already exists: {filePath}");
                }

                var fileStream = new FileStream(filePath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
                
                var pst = new PstFile
                {
                    _fileStream = fileStream,
                    _isReadOnly = false,
                    FilePath = filePath,
                    _folderCache = new Dictionary<uint, PstFolder>()
                };

                // Create new PST file structure
                pst.CreateNewPstStructure(formatType);
                
                return pst;
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstAccessException($"Failed to create PST file: {filePath}", ex);
            }
        }

        /// <summary>
        /// Copies all folders and messages from the source PST file to this PST file.
        /// </summary>
        /// <param name="sourcePst">Source PST file to copy from.</param>
        /// <param name="progressCallback">Optional callback to report progress (0.0 to 1.0)</param>
        /// <exception cref="PstException">Thrown when the copying operation fails.</exception>
        public void CopyFrom(PstFile sourcePst, Action<double>? progressCallback = null)
        {
            // Call the overload with no filter (copy all messages)
            CopyFrom(sourcePst, null, progressCallback);
        }
        
        /// <summary>
        /// Copies filtered messages from another PST file into this one.
        /// </summary>
        /// <param name="sourcePst">The source PST file to copy from.</param>
        /// <param name="filter">A filter to apply to the messages being copied, or null to copy all messages.</param>
        /// <param name="progressCallback">Optional callback to report progress (0.0 to 1.0)</param>
        /// <exception cref="PstException">Thrown when the copying operation fails.</exception>
        public void CopyFrom(PstFile sourcePst, MessageFilter? filter, Action<double>? progressCallback = null)
        {
            if (_isReadOnly)
            {
                throw new PstAccessException("Cannot copy to a read-only PST file.");
            }

            try
            {
                var sourceRoot = sourcePst.RootFolder;
                var destRoot = this.RootFolder;
                
                // Count total messages for progress reporting
                int totalMessages = CountMessages(sourceRoot);
                int processedMessages = 0;

                CopyFolderRecursive(sourceRoot, destRoot, filter, ref processedMessages, totalMessages, progressCallback);

                // Make sure changes are written to disk
                Flush();
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException("Failed to copy PST content", ex);
            }
        }

        /// <summary>
        /// Flushes any pending changes to disk.
        /// </summary>
        public void Flush()
        {
            if (!_isReadOnly)
            {
                _fileStream?.Flush();
            }
        }

        /// <summary>
        /// Disposes the PST file, releasing resources.
        /// </summary>
        
        /// <summary>
        /// Finds a folder by path.
        /// </summary>
        /// <param name="folderPath">Path to the folder, using '/' as separator (e.g., "Inbox/Subfolder").</param>
        /// <returns>The found folder, or null if not found.</returns>
        public PstFolder? FindFolder(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                return RootFolder;
            }
            
            // Split path into parts
            string[] parts = folderPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            
            PstFolder? currentFolder = RootFolder;
            
            // Traverse the folder hierarchy
            foreach (var part in parts)
            {
                if (currentFolder == null)
                {
                    break;
                }
                
                // Special handling for root folder references
                if (part.Equals("Root", StringComparison.OrdinalIgnoreCase) && 
                    currentFolder == RootFolder)
                {
                    continue;
                }
                
                currentFolder = currentFolder.FindFolder(part, false);
            }
            
            return currentFolder;
        }
        
        /// <summary>
        /// Gets messages from a folder with filtering.
        /// </summary>
        /// <param name="folderPath">Path to the folder, using '/' as separator (e.g., "Inbox/Subfolder").</param>
        /// <param name="filter">Optional filter to apply to the messages.</param>
        /// <returns>A filtered list of messages, or empty list if folder not found.</returns>
        public IEnumerable<PstMessage> GetMessages(string folderPath, MessageFilter? filter = null)
        {
            var folder = FindFolder(folderPath);
            if (folder == null)
            {
                return Enumerable.Empty<PstMessage>();
            }
            
            if (filter == null)
            {
                return folder.Messages;
            }
            
            return folder.GetFilteredMessages(filter);
        }
        
        /// <summary>
        /// Disposes resources used by the PST file.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Releases the unmanaged resources used by the PstFile and optionally releases the managed resources.
        /// </summary>
        /// <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _fileStream?.Dispose();
                }

                _disposed = true;
            }
        }

        private void Initialize()
        {
            if (_fileStream == null)
                throw new PstCorruptedException("File stream is not initialized");
                
            using var reader = new PstBinaryReader(_fileStream, leaveOpen: true);
            
            // Read and validate the PST file header
            _header = PstFormatHeader.Read(reader);
            
            if (!_header.IsValid)
            {
                throw new PstCorruptedException("Invalid PST file format. File is corrupted or not a PST file.");
            }

            // Initialize B-trees
            InitializeBTrees();
        }

        private void InitializeBTrees()
        {
            if (_header == null)
                throw new PstCorruptedException("PST header is not initialized");
                
            // Load the node B-tree (NBT)
            _nodeBTree = new BTreeOnHeap(this, _header.NodeBTreeRoot);
            
            // Load the block B-tree (BBT)
            _blockBTree = new BTreeOnHeap(this, _header.BlockBTreeRoot);
        }

        private void InitializeRootFolder()
        {
            if (_header == null)
                throw new PstCorruptedException("PST header is not initialized");
                
            // Find the root folder node
            var rootEntry = _nodeBTree!.FindNodeByNid(_header.RootFolderId);
            if (rootEntry == null)
            {
                throw new PstCorruptedException("Root folder not found in PST file");
            }

            _rootFolder = new PstFolder(this, rootEntry);
            if (_folderCache == null)
            {
                _folderCache = new Dictionary<uint, PstFolder>();
            }
            _folderCache[_header.RootFolderId] = _rootFolder;
        }

        private void CreateNewPstStructure(PstFormatType formatType)
        {
            // Initialize a new PST header
            _header = PstFormatHeader.CreateNew(formatType);
            
            if (_fileStream == null)
                throw new PstCorruptedException("File stream is not initialized");
                
            // Use PstBinaryWriter but make sure to keep the stream open for future operations
            var writer = new PstBinaryWriter(_fileStream, leaveOpen: true);
            
            try
            {
                // Write the initial PST file structure
                _header.Write(writer);
            }
            finally
            {
                // Dispose the writer but keep the stream open (leaveOpen: true)
                writer.Dispose();
            }
            
            // Initialize empty trees
            InitializeEmptyPstStructure();
            
            // Create the root folder
            InitializeRootFolder();
            
            // Flush changes to disk to ensure the file structure is written
            Flush();
        }

        private void InitializeEmptyPstStructure()
        {
            if (_header == null)
                throw new PstCorruptedException("PST header is not initialized");
                
            // Initialize B-trees for node and block management
            _nodeBTree = BTreeOnHeap.CreateNew(this, _header.NodeBTreeRoot);
            _blockBTree = BTreeOnHeap.CreateNew(this, _header.BlockBTreeRoot);
            
            try
            {
                // Create basic PST structure with essential tables and nodes
                
                // 1. Create root folder node (NID_ROOT)
                uint rootFolderId = _header.RootFolderId;
                var rootNode = new NdbNodeEntry(
                    rootFolderId,
                    1000u,  // Data ID
                    0u,     // Parent ID (root has no parent)
                    0ul,    // Data offset
                    512u    // Data size
                );
                rootNode.DisplayName = "Root Folder";
                _nodeBTree.AddNode(rootNode, new byte[0]);
                
                // 2. Create inbox folder (standard PST folder)
                uint inboxId = GenerateNodeId(PstNodeTypes.NID_TYPE_FOLDER, 1);
                var inboxNode = new NdbNodeEntry(
                    inboxId,
                    1001u,
                    rootFolderId, // Parent is root
                    0ul,
                    512u
                );
                inboxNode.DisplayName = "Inbox";
                _nodeBTree.AddNode(inboxNode, new byte[0]);
                
                // 3. Create standard Sent Items folder
                uint sentItemsId = GenerateNodeId(PstNodeTypes.NID_TYPE_FOLDER, 2);
                var sentItemsNode = new NdbNodeEntry(
                    sentItemsId,
                    1002u,
                    rootFolderId, // Parent is root
                    0ul,
                    512u
                );
                sentItemsNode.DisplayName = "Sent Items";
                _nodeBTree.AddNode(sentItemsNode, new byte[0]);
                
                // 4. Create standard Deleted Items folder
                uint deletedItemsId = GenerateNodeId(PstNodeTypes.NID_TYPE_FOLDER, 3);
                var deletedItemsNode = new NdbNodeEntry(
                    deletedItemsId,
                    1003u,
                    rootFolderId, // Parent is root
                    0ul,
                    512u
                );
                deletedItemsNode.DisplayName = "Deleted Items";
                _nodeBTree.AddNode(deletedItemsNode, new byte[0]);
                
                // 5. Initialize properties for the root folder
                var rootPropertyContext = new PropertyContext(this, rootNode);
                rootPropertyContext.SetProperty(
                    PstStructure.PropertyIds.PidTagDisplayName, 
                    PstStructure.PropertyType.PT_STRING8, 
                    "Root Folder");
                
                // 6. Mark root folder as having subfolders
                rootPropertyContext.SetProperty(
                    PstStructure.PropertyIds.PidTagSubfolders, 
                    PstStructure.PropertyType.PT_BOOLEAN, 
                    true);
                
                // Flush changes to disk
                Flush();
            }
            catch (Exception ex)
            {
                throw new PstException("Failed to initialize PST structure", ex);
            }
        }
        
        /// <summary>
        /// Generates a node ID using the specified type and ID components.
        /// </summary>
        /// <param name="nodeType">The node type.</param>
        /// <param name="id">The ID component.</param>
        /// <returns>A combined node ID.</returns>
        internal uint GenerateNodeId(ushort nodeType, ushort id)
        {
            return PstNodeTypes.CreateNodeId(nodeType, id);
        }

        // GetFileStream is already defined higher up in the class
        
        private int CountMessages(PstFolder folder)
        {
            int count = folder.Messages.Count;
            
            foreach (var subFolder in folder.SubFolders)
            {
                count += CountMessages(subFolder);
            }
            
            return count;
        }
        
        private void CopyFolderRecursive(PstFolder sourceFolder, PstFolder destFolder, 
            ref int processedMessages, int totalMessages, Action<double>? progressCallback)
        {
            // Call the overload with no filter (copy all messages)
            CopyFolderRecursive(sourceFolder, destFolder, null, ref processedMessages, totalMessages, progressCallback);
        }
        
        private void CopyFolderRecursive(PstFolder sourceFolder, PstFolder destFolder, 
            MessageFilter? filter, ref int processedMessages, int totalMessages, Action<double>? progressCallback)
        {
            // Get all messages in the source folder
            var messages = sourceFolder.Messages.ToList();
            
            // Apply filter if specified to avoid processing messages that don't match
            IEnumerable<PstMessage> filteredMessages = messages;
            if (filter != null)
            {
                filteredMessages = filter.Apply(messages);
            }
            
            // Process only the filtered messages
            foreach (var message in filteredMessages)
            {
                try
                {
                    // Get message raw content and create a copy
                    byte[] rawContent = message.GetRawContent();
                    
                    // Create a MimeMessage from the original message to ensure all properties are copied
                    var mimeMessage = message.ToMimeMessage();
                    
                    // Create a new message from the MimeMessage to preserve all metadata
                    var newMessage = PstMessage.Create(this, mimeMessage);
                    
                    // Ensure critical properties are set correctly
                    if (string.IsNullOrEmpty(newMessage.Subject) && !string.IsNullOrEmpty(message.Subject))
                    {
                        // Create a new message using the direct method as fallback if needed
                        newMessage = PstMessage.Create(this, 
                            message.Subject, 
                            message.BodyText ?? string.Empty, 
                            message.SenderEmail ?? string.Empty, 
                            message.SenderName ?? string.Empty);
                    }
                    
                    // Copy all attachments including nested email attachments
                    if (message.HasAttachments)
                    {
                        // Get all attachments
                        var attachments = message.GetAttachments();
                        
                        foreach (var attachment in attachments)
                        {
                            // Get the attachment content
                            byte[] attachmentContent = attachment.GetContent();
                            
                            // Check if it's an email attachment
                            if (attachment.IsEmailMessage)
                            {
                                // Process it as an email
                                var emailAttachment = attachment.GetAsEmailMessage();
                                if (emailAttachment != null)
                                {
                                    // Add the email as an attachment with all its data intact
                                    newMessage.AddEmailAttachment(emailAttachment);
                                }
                                else
                                {
                                    // If failed to parse as email, add as regular attachment
                                    newMessage.AddAttachment(attachment.Filename, attachmentContent, attachment.ContentType);
                                }
                            }
                            else
                            {
                                // Regular attachment, just copy it
                                newMessage.AddAttachment(attachment.Filename, attachmentContent, attachment.ContentType);
                            }
                        }
                    }
                    
                    // Add the copied message to the destination folder
                    destFolder.AddMessage(newMessage);
                    
                    // Update progress
                    processedMessages++;
                    if (progressCallback != null && totalMessages > 0)
                    {
                        double progress = (double)processedMessages / totalMessages;
                        progressCallback(progress);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error copying message {message.Subject}: {ex.Message}");
                    // Continue with the next message
                }
            }
            
            // Process subfolders
            foreach (var sourceSubFolder in sourceFolder.SubFolders)
            {
                // Create corresponding subfolder in destination
                var destSubFolder = destFolder.CreateSubFolder(sourceSubFolder.Name);
                
                // Recursively copy contents
                CopyFolderRecursive(sourceSubFolder, destSubFolder, filter,
                    ref processedMessages, totalMessages, progressCallback);
            }
        }

        internal NdbNodeEntry? GetNodeEntry(uint nodeId)
        {
            return _nodeBTree?.FindNodeByNid(nodeId);
        }

        internal PropertyContext? GetPropertyContext(uint nodeId)
        {
            var nodeEntry = GetNodeEntry(nodeId);
            if (nodeEntry == null) return null;
            
            return new PropertyContext(this, nodeEntry);
        }
        
        internal BTreeOnHeap GetNodeBTree()
        {
            if (_nodeBTree == null)
            {
                throw new PstCorruptedException("Node B-tree is not initialized");
            }
            
            return _nodeBTree;
        }

        /// <summary>
        /// Reads a block of data from the PST file at the specified offset and size.
        /// Optimized for handling large blocks efficiently.
        /// </summary>
        /// <param name="offset">The offset in the file to read from.</param>
        /// <param name="size">The size of the block to read.</param>
        /// <returns>The data as a byte array.</returns>
        internal byte[] ReadBlock(ulong offset, uint size)
        {
            if (_fileStream == null)
                throw new PstCorruptedException("File stream is not initialized");
                
            if (size == 0)
                return Array.Empty<byte>();
                
            _fileStream.Position = (long)offset;
            
            const int chunkSize = 8192; // 8KB chunks for efficient reading
            
            // For small blocks, use a simple read
            if (size <= chunkSize)
            {
                var buffer = new byte[size];
                _fileStream.Read(buffer, 0, (int)size);
                return buffer;
            }
            
            // For large blocks, read in chunks for better performance and memory usage
            var result = new byte[size];
            int bytesRead = 0;
            int remaining = (int)size;
            
            while (remaining > 0)
            {
                int toRead = Math.Min(chunkSize, remaining);
                int read = _fileStream.Read(result, bytesRead, toRead);
                
                if (read == 0)
                    break; // End of file
                    
                bytesRead += read;
                remaining -= read;
            }
            
            return result;
        }
        
        /// <summary>
        /// Asynchronously reads a block of data from the PST file at the specified offset and size.
        /// Optimized for handling large blocks efficiently.
        /// </summary>
        /// <param name="offset">The offset in the file to read from.</param>
        /// <param name="size">The size of the block to read.</param>
        /// <param name="cancellationToken">Optional cancellation token for the async operation.</param>
        /// <returns>The data as a byte array.</returns>
        internal async Task<byte[]> ReadBlockAsync(ulong offset, uint size, CancellationToken cancellationToken = default)
        {
            if (_fileStream == null)
                throw new PstCorruptedException("File stream is not initialized");
                
            if (size == 0)
                return Array.Empty<byte>();
                
            _fileStream.Position = (long)offset;
            
            const int chunkSize = 8192; // 8KB chunks for efficient reading
            
            // For small blocks, use a simple read
            if (size <= chunkSize)
            {
                var buffer = new byte[size];
                await _fileStream.ReadAsync(buffer, 0, (int)size, cancellationToken).ConfigureAwait(false);
                return buffer;
            }
            
            // For large blocks, read in chunks for better performance and memory usage
            var result = new byte[size];
            int bytesRead = 0;
            int remaining = (int)size;
            
            while (remaining > 0 && !cancellationToken.IsCancellationRequested)
            {
                int toRead = Math.Min(chunkSize, remaining);
                int read = await _fileStream.ReadAsync(result, bytesRead, toRead, cancellationToken).ConfigureAwait(false);
                
                if (read == 0)
                    break; // End of file
                    
                bytesRead += read;
                remaining -= read;
            }
            
            cancellationToken.ThrowIfCancellationRequested();
            return result;
        }
        
        /// <summary>
        /// Allocates a block of storage in the PST file for the specified data block ID.
        /// </summary>
        /// <param name="blockId">The block ID to allocate space for.</param>
        /// <param name="size">The size of the block to allocate.</param>
        /// <returns>The offset in the file where the block was allocated.</returns>
        internal ulong AllocateBlock(uint blockId, uint size)
        {
            if (_isReadOnly)
                throw new PstAccessException("Cannot allocate blocks in a read-only PST file");
                
            if (_fileStream == null)
                throw new PstCorruptedException("File stream is not initialized");
                
            // Implement block allocation using file allocation table
            // This allocates a block of specified size in the PST file and returns its offset
            // The allocation strategy follows MS-PST specification for block allocation
            
            // First check if we have a lookup table of free blocks
            Dictionary<uint, ulong> freeBlockTable = GetFreeBlockTable();
            
            // If we have a suitable free block in our table, use it
            ulong offset = 0;
            
            // Look for existing free space of suitable size
            if (freeBlockTable.Count > 0)
            {
                foreach (var entry in freeBlockTable.OrderBy(e => e.Value))
                {
                    // If the block is big enough for our data, use it
                    if (entry.Key >= size)
                    {
                        offset = entry.Value;
                        freeBlockTable.Remove(entry.Key);
                        
                        // Update the free block table
                        UpdateFreeBlockTable(freeBlockTable);
                        
                        // Return this block offset
                        return offset;
                    }
                }
            }
            
            // If no suitable free blocks exist, append to the end of the file
            offset = (ulong)_fileStream.Length;
            
            // Align to block boundary (typically 64 bytes in PST)
            if (offset % 64 != 0)
            {
                offset = ((offset / 64) + 1) * 64;
            }
            
            // Return the allocated offset
            return offset;
        }

        /// <summary>
        /// Writes a block of data to the PST file at the specified offset.
        /// Optimized for handling large blocks efficiently.
        /// </summary>
        /// <param name="offset">The offset in the file to write to.</param>
        /// <param name="data">The data to write.</param>
        internal void WriteBlock(ulong offset, byte[] data)
        {
            if (_isReadOnly)
                throw new PstAccessException("Cannot write to a read-only PST file");
                
            if (_fileStream == null)
                throw new PstCorruptedException("File stream is not initialized");
            
            if (data == null || data.Length == 0)
                return;
                
            _fileStream.Position = (long)offset;
            
            const int chunkSize = 8192; // 8KB chunks for efficient writing
            
            // For small blocks, use a simple write
            if (data.Length <= chunkSize)
            {
                _fileStream.Write(data, 0, data.Length);
                return;
            }
            
            // For large blocks, write in chunks for better performance
            int written = 0;
            int remaining = data.Length;
            
            while (remaining > 0)
            {
                int toWrite = Math.Min(chunkSize, remaining);
                _fileStream.Write(data, written, toWrite);
                
                written += toWrite;
                remaining -= toWrite;
            }
        }

        /// <summary>
        /// Registers a folder in the folder cache.
        /// </summary>
        /// <param name="folderId">The folder ID.</param>
        /// <param name="folder">The folder object.</param>
        internal void RegisterFolder(uint folderId, PstFolder folder)
        {
            if (_folderCache == null)
            {
                _folderCache = new Dictionary<uint, PstFolder>();
            }
            
            _folderCache[folderId] = folder;
        }

        /// <summary>
        /// Gets a folder from the folder cache.
        /// </summary>
        /// <param name="folderId">The folder ID.</param>
        /// <returns>The folder, or null if not found in the cache.</returns>
        internal PstFolder? GetCachedFolder(uint folderId)
        {
            if (_folderCache == null)
            {
                return null;
            }
            
            return _folderCache.TryGetValue(folderId, out var folder) ? folder : null;
        }
        
        /// <summary>
        /// Gets the free block table from the PST file.
        /// </summary>
        /// <returns>A dictionary mapping block sizes to their offsets.</returns>
        private Dictionary<uint, ulong> GetFreeBlockTable()
        {
            var result = new Dictionary<uint, ulong>();
            
            // In PST files, there's a special node that contains the free block table
            // This node is typically at NID 0x21 (ANSI) or 0x42 (Unicode)
            // We use the BTH (B-Tree Heap) node at 0x60 (ANSI) or 0x61 (Unicode) to store free block information
            
            try
            {
                // First, check if we have a free block table node
                uint bthNodeId = IsAnsi ? 0x60u : 0x61u;
                var bthNode = _nodeBTree?.FindNodeByNid(bthNodeId);
                
                if (bthNode != null)
                {
                    // Get the binary data for the BTH node
                    var bthData = _nodeBTree?.GetNodeData(bthNode);
                    
                    if (bthData != null && bthData.Length > 16)
                    {
                        // Parse the BTH data to find free blocks
                        // The format is typically:
                        // - 8 bytes: BTH header
                        // - 4 bytes: count of free blocks
                        // - For each free block:
                        //   - 4 bytes: size of block
                        //   - 8 bytes: offset of block
                        
                        int count = BitConverter.ToInt32(bthData, 8);
                        
                        // Sanity check - don't try to parse unreasonable counts
                        if (count > 0 && count < 1000)
                        {
                            int offset = 12; // Start after header and count
                            
                            for (int i = 0; i < count && offset + 12 <= bthData.Length; i++)
                            {
                                uint blockSize = BitConverter.ToUInt32(bthData, offset);
                                ulong blockOffset = BitConverter.ToUInt64(bthData, offset + 4);
                                
                                // Add to our dictionary
                                result[blockSize] = blockOffset;
                                
                                // Move to next entry
                                offset += 12;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // If we can't read the free block table, log but continue with empty table
                Console.WriteLine($"Error reading free block table: {ex.Message}");
            }
            
            return result;
        }
        
        /// <summary>
        /// Updates the free block table in the PST file.
        /// </summary>
        /// <param name="freeBlockTable">The dictionary of free blocks.</param>
        private void UpdateFreeBlockTable(Dictionary<uint, ulong> freeBlockTable)
        {
            if (_isReadOnly)
                return;
                
            try
            {
                // Prepare data for the BTH node
                // - 8 bytes: BTH header (format version, etc.)
                // - 4 bytes: count of free blocks
                // - For each free block:
                //   - 4 bytes: size of block
                //   - 8 bytes: offset of block
                
                int count = freeBlockTable.Count;
                int dataSize = 12 + (count * 12); // 12 bytes for header + count, then 12 bytes per entry
                
                byte[] bthData = new byte[dataSize];
                
                // Set BTH header (version 1.0)
                BitConverter.GetBytes((ulong)0x0100u).CopyTo(bthData, 0);
                
                // Set count of free blocks
                BitConverter.GetBytes(count).CopyTo(bthData, 8);
                
                // Add each free block
                int offset = 12;
                foreach (var entry in freeBlockTable)
                {
                    BitConverter.GetBytes(entry.Key).CopyTo(bthData, offset);
                    BitConverter.GetBytes(entry.Value).CopyTo(bthData, offset + 4);
                    offset += 12;
                }
                
                // Get the BTH node (or create it if it doesn't exist)
                uint bthNodeId = IsAnsi ? 0x60u : 0x61u;
                var bthNode = _nodeBTree?.FindNodeByNid(bthNodeId);
                
                if (bthNode != null)
                {
                    // Update existing node
                    _nodeBTree?.UpdateNodeData(bthNode, bthData);
                }
                else
                {
                    // Create a new node if it doesn't exist
                    var dataId = bthNodeId & 0x00FFFFFF;
                    var newNode = new NdbNodeEntry(bthNodeId, dataId, 0, 0, 0)
                    {
                        DisplayName = "Free Block Table"
                    };
                    
                    _nodeBTree?.AddNode(newNode, bthData);
                }
            }
            catch (Exception ex)
            {
                // If we can't update the free block table, log but continue
                Console.WriteLine($"Error updating free block table: {ex.Message}");
            }
        }
    }
}
