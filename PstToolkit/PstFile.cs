using System;
using System.Collections.Generic;
using System.IO;
using PstToolkit.Exceptions;
using PstToolkit.Formats;
using PstToolkit.Utils;

namespace PstToolkit
{
    /// <summary>
    /// Represents a PST file that can be read from or written to.
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

                CopyFolderRecursive(sourceRoot, destRoot, ref processedMessages, totalMessages, progressCallback);

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
                
            using var writer = new PstBinaryWriter(_fileStream, leaveOpen: true);
            
            // Write the initial PST file structure
            _header.Write(writer);
            
            // Initialize empty trees
            InitializeEmptyPstStructure();
            
            // Create the root folder
            InitializeRootFolder();
        }

        private void InitializeEmptyPstStructure()
        {
            // Implementation would create the basic structure of a new PST file
            // including node and block B-trees, and the necessary tables
            // This is a complex process that involves creating various metadata structures
            
            if (_header == null)
                throw new PstCorruptedException("PST header is not initialized");
                
            // For now, we'll just create placeholder structures
            _nodeBTree = BTreeOnHeap.CreateNew(this, _header.NodeBTreeRoot);
            _blockBTree = BTreeOnHeap.CreateNew(this, _header.BlockBTreeRoot);
        }

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
            // Copy messages from source folder to destination folder
            foreach (var message in sourceFolder.Messages)
            {
                destFolder.AddMessage(message);
                processedMessages++;
                
                if (progressCallback != null && totalMessages > 0)
                {
                    double progress = (double)processedMessages / totalMessages;
                    progressCallback(progress);
                }
            }
            
            // Process subfolders
            foreach (var sourceSubFolder in sourceFolder.SubFolders)
            {
                // Create corresponding subfolder in destination
                var destSubFolder = destFolder.CreateSubFolder(sourceSubFolder.Name);
                
                // Recursively copy contents
                CopyFolderRecursive(sourceSubFolder, destSubFolder, ref processedMessages, totalMessages, progressCallback);
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

        internal byte[] ReadBlock(ulong offset, uint size)
        {
            if (_fileStream == null)
                throw new PstCorruptedException("File stream is not initialized");
                
            _fileStream.Position = (long)offset;
            var buffer = new byte[size];
            _fileStream.Read(buffer, 0, (int)size);
            return buffer;
        }

        internal void WriteBlock(ulong offset, byte[] data)
        {
            if (_isReadOnly)
                throw new PstAccessException("Cannot write to a read-only PST file");
                
            if (_fileStream == null)
                throw new PstCorruptedException("File stream is not initialized");
                
            _fileStream.Position = (long)offset;
            _fileStream.Write(data, 0, data.Length);
        }

        internal void RegisterFolder(uint folderId, PstFolder folder)
        {
            if (_folderCache == null)
            {
                _folderCache = new Dictionary<uint, PstFolder>();
            }
            
            _folderCache[folderId] = folder;
        }

        internal PstFolder? GetCachedFolder(uint folderId)
        {
            if (_folderCache == null)
            {
                return null;
            }
            
            return _folderCache.TryGetValue(folderId, out var folder) ? folder : null;
        }
    }
}
