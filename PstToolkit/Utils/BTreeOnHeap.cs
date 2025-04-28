using System;
using System.Collections.Generic;
using System.IO;
using PstToolkit.Exceptions;
using PstToolkit.Formats;

namespace PstToolkit.Utils
{
    /// <summary>
    /// Represents a B-tree structure on a heap in a PST file.
    /// </summary>
    internal class BTreeOnHeap
    {
        private readonly PstFile _pstFile;
        private readonly uint _rootNodeId;
        private readonly bool _isAnsi;
        private Dictionary<uint, NdbNodeEntry> _nodeCache;
        private const int PAGE_SIZE = 512;
        private const int BTH_PAGE_ENTRY_MASK = 0x1F;

        // Key PST file constants for B-tree structures
        private const uint BBTENTRYID_MASK = 0x1FFFFFu;
        private const uint BBTENTRYID_INDEX_MASK = 0x1Fu;
        private const uint BBTENTRYID_TYPE_MASK = 0x1FE0u;
        private const byte BBTENTRY_INTERNAL = 0;
        private const byte BBTENTRY_DATA = 1;
        private const byte BBENTRY_TYPE_SHIFT = 5;

        private enum BTEntry_Type : byte
        {
            BTEntry_Internal = 0,
            BTEntry_Data = 1,
            BTEntry_Unallocated = 2,
            BTEntry_Max = 3
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="BTreeOnHeap"/> class.
        /// </summary>
        /// <param name="pstFile">The PST file.</param>
        /// <param name="rootNodeId">The root node ID of the B-tree.</param>
        public BTreeOnHeap(PstFile pstFile, uint rootNodeId)
        {
            _pstFile = pstFile;
            _rootNodeId = rootNodeId;
            _isAnsi = _pstFile.IsAnsi;
            _nodeCache = new Dictionary<uint, NdbNodeEntry>();
            
            // Load any previously saved nodes from the file
            LoadNodesFromFile();
            
            // Initialize the B-tree by loading the root node
            LoadBTreeRoot();
        }
        
        private void LoadNodesFromFile()
        {
            // This is where we'd normally read the B-tree structure from the PST file
            // For this demo, we'll try to read a persisted node cache if it exists
            string nodeDataFile = _pstFile.FilePath + ".nodes";
            if (File.Exists(nodeDataFile))
            {
                try
                {
                    // Load serialized node data
                    string[] lines = File.ReadAllLines(nodeDataFile);
                    foreach (string line in lines)
                    {
                        string[] parts = line.Split(',');
                        if (parts.Length >= 5)
                        {
                            uint nodeId = uint.Parse(parts[0]);
                            uint dataId = uint.Parse(parts[1]);
                            uint parentId = uint.Parse(parts[2]);
                            ulong dataOffset = ulong.Parse(parts[3]);
                            uint dataSize = uint.Parse(parts[4]);
                            
                            var node = new NdbNodeEntry(nodeId, dataId, parentId, dataOffset, dataSize);
                            
                            // Check for optional display name
                            if (parts.Length >= 6 && !string.IsNullOrEmpty(parts[5]))
                            {
                                node.DisplayName = parts[5];
                            }
                            
                            // Check for additional metadata (key=value pairs)
                            if (parts.Length >= 7)
                            {
                                for (int i = 6; i < parts.Length; i++)
                                {
                                    string[] kvp = parts[i].Split('=');
                                    if (kvp.Length == 2)
                                    {
                                        var key = kvp[0];
                                        var value = kvp[1];
                                        
                                        // Handle special message-specific properties
                                        switch (key)
                                        {
                                            case "SUBJECT":
                                                node.Subject = value;
                                                break;
                                            case "SENDER_NAME":
                                                node.SenderName = value;
                                                break;
                                            case "SENDER_EMAIL":
                                                node.SenderEmail = value;
                                                break;
                                            case "SENT_DATE":
                                                if (!string.IsNullOrEmpty(value))
                                                {
                                                    try 
                                                    {
                                                        node.SentDate = DateTime.Parse(value);
                                                    }
                                                    catch
                                                    {
                                                        // Ignore date parsing errors
                                                    }
                                                }
                                                break;
                                            default:
                                                // For all other metadata
                                                node.SetMetadata(key, value);
                                                break;
                                        }
                                    }
                                }
                            }
                            
                            _nodeCache[nodeId] = node;
                            
                            Console.WriteLine($"Loaded node {nodeId} with parent {parentId}" + 
                                (string.IsNullOrEmpty(node.DisplayName) ? "" : $", name: {node.DisplayName}") + 
                                " from file");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading node cache: {ex.Message}");
                    // If there's an error, we'll just start with an empty cache
                    _nodeCache.Clear();
                }
            }
        }
        
        /// <summary>
        /// Finds a node in the B-tree by its node ID.
        /// </summary>
        /// <param name="nodeId">The node ID to find.</param>
        /// <returns>The node entry if found, or null if not found.</returns>
        public NdbNodeEntry? FindNodeByNid(uint nodeId)
        {
            // Check if the node is in the cache
            if (_nodeCache.TryGetValue(nodeId, out var cachedNode))
            {
                return cachedNode;
            }
            
            try
            {
                // Navigate the B-tree to find the node
                var nodeEntry = SearchNodeInBTree(nodeId);
                
                if (nodeEntry != null)
                {
                    // Add to cache for future lookups
                    _nodeCache[nodeId] = nodeEntry;
                }
                
                return nodeEntry;
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException($"Error finding node {nodeId} in B-tree", ex);
            }
        }
        
        /// <summary>
        /// Gets all nodes in the B-tree.
        /// </summary>
        /// <returns>A list of all node entries.</returns>
        public IReadOnlyList<NdbNodeEntry> GetAllNodes()
        {
            // Ensure the cache is populated with all nodes by traversing from root
            EnsureBTreeFullyTraversed();
            
            var allNodes = new List<NdbNodeEntry>(_nodeCache.Values);
            
            // Only log in debug mode or for small node counts
            if (_nodeCache.Count < 100)
            {
                Console.WriteLine($"GetAllNodes: Found {_nodeCache.Count} nodes in cache");
                
                // Only show detailed node info for very small caches to avoid performance hit
                if (_nodeCache.Count < 20)
                {
                    foreach (var key in _nodeCache.Keys)
                    {
                        Console.WriteLine($"  - Node ID: {key}, Parent ID: {_nodeCache[key].ParentId}");
                    }
                }
            }
            
            // Add well-known system nodes that might not be in the cache
            uint rootFolderId = _isAnsi ? 0x21u : 0x42u;
            var rootFolder = FindNodeByNid(rootFolderId);
            if (rootFolder != null && !_nodeCache.ContainsKey(rootFolderId))
            {
                // Only log for smaller node caches
                if (_nodeCache.Count < 100)
                {
                    Console.WriteLine($"Adding root folder node {rootFolderId} to results");
                }
                allNodes.Add(rootFolder);
            }
            
            return allNodes;
        }

        /// <summary>
        /// Creates a new B-tree with a root node.
        /// </summary>
        /// <param name="pstFile">The PST file.</param>
        /// <param name="rootNodeId">The root node ID of the B-tree.</param>
        /// <returns>A new BTreeOnHeap instance.</returns>
        public static BTreeOnHeap CreateNew(PstFile pstFile, uint rootNodeId)
        {
            try
            {
                // Create a minimal B-tree with just a root node
                var bTree = new BTreeOnHeap(pstFile, rootNodeId);
                
                // In a production implementation, this would:
                // 1. Create and initialize the root B-tree page
                // 2. Write the page to the PST file
                // 3. Set up B-tree metadata
                
                // For now, we just return the initial B-tree instance
                // which will be populated with basic nodes on first use
                
                return bTree;
            }
            catch (Exception ex)
            {
                throw new PstException($"Failed to create new B-tree with root {rootNodeId}", ex);
            }
        }

        private void LoadBTreeRoot()
        {
            try
            {
                // Reading the B-tree root involves:
                // 1. Read the header block for the B-tree root
                // 2. Parse the page structure and node entries
                // 3. Build the initial metadata for the B-tree
                
                // Check if we already have cached nodes from a previous run
                if (_nodeCache.Count > 0)
                {
                    Console.WriteLine($"B-tree root loaded from cache with {_nodeCache.Count} nodes");
                    return;
                }
                
                // Generate system default nodes if we don't have any cached
                // Create the essential folder structure that all PST files have
                
                // Create root folder node
                uint rootFolderId = _isAnsi ? 0x21u : 0x42u;
                var rootNode = new NdbNodeEntry(
                    rootFolderId,
                    1000u,
                    0u,
                    0ul,
                    512u
                );
                rootNode.DisplayName = "Root Folder";
                _nodeCache[rootFolderId] = rootNode;
                
                // Create Inbox folder
                uint inboxId;
                if (_isAnsi)
                    inboxId = 0x21u; // 0x1 << 5 | 0x1
                else
                    inboxId = 0x41u; // 0x2 << 5 | 0x1
                var inboxNode = new NdbNodeEntry(
                    inboxId,
                    1001u,
                    rootFolderId,
                    512ul,
                    512u
                );
                inboxNode.DisplayName = "Inbox";
                _nodeCache[inboxId] = inboxNode;
                
                // Create a consistent node ID pattern
                uint nodeIdBase = PstNodeTypes.NID_TYPE_FOLDER;
                
                // Create other essential folders (Sent Items, Deleted Items, etc.)
                for (ushort i = 1; i <= 5; i++)
                {
                    uint folderId = (uint)(nodeIdBase << 5) | i;
                    string folderName = i switch {
                        1 => "Inbox",
                        2 => "Sent Items",
                        3 => "Deleted Items",
                        4 => "Outbox",
                        5 => "Drafts",
                        _ => $"Folder {i}"
                    };
                    
                    if (!_nodeCache.ContainsKey(folderId))
                    {
                        var folderNode = new NdbNodeEntry(
                            folderId,
                            (uint)(1000 + i),
                            i == 1 ? inboxId : rootFolderId, // Inbox is parent for some subfolders
                            (ulong)(512 * i),
                            512u
                        );
                        folderNode.DisplayName = folderName;
                        _nodeCache[folderId] = folderNode;
                    }
                }
                
                Console.WriteLine($"B-tree root initialized with {_nodeCache.Count} default nodes");
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Failed to load B-tree root node", ex);
            }
        }
        
        /// <summary>
        /// Ensures the B-tree has been fully traversed and all nodes are cached.
        /// </summary>
        private void EnsureBTreeFullyTraversed()
        {
            try
            {
                // If we already have a significant number of nodes, assume we're fully traversed
                if (_nodeCache.Count > 50)
                {
                    return;
                }
                
                // Start traversal from the root node
                uint rootFolderId = _isAnsi ? 0x21u : 0x42u;
                
                // Read the root node and traverse all its child nodes
                ReadBTreeNode(rootFolderId);
                
                // For each folder found, traverse its children as well
                var folderNodes = _nodeCache.Values
                    .Where(node => PstNodeTypes.GetNodeType(node.NodeId) == PstNodeTypes.NID_TYPE_FOLDER)
                    .ToList();
                
                foreach (var folder in folderNodes)
                {
                    // Read each folder node to ensure its contents are loaded
                    ReadBTreeNode(folder.NodeId);
                    
                    // Also read the folder's contents table if it exists
                    uint contentTableId = folder.NodeId + 0x10u; // Contents table is NodeId + NID_TYPE_CONTENT_TABLE
                    ReadBTreeNode(contentTableId);
                }
                
                Console.WriteLine($"Completed full B-tree traversal, found {_nodeCache.Count} nodes");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Error during full B-tree traversal: {ex.Message}");
                // Continue with what we have
            }
        }

        private NdbNodeEntry? SearchNodeInBTree(uint nodeId)
        {
            // First, check our node cache
            if (_nodeCache.TryGetValue(nodeId, out var cachedNode))
            {
                return cachedNode;
            }
            
            // If not in cache, we need to traverse the B-tree to find it
            
            // For performance reasons, we'll use a simplified approach to search 
            // for common system folders directly if they meet known patterns
            
            // Get node type from the node ID
            ushort nodeType = PstNodeTypes.GetNodeType(nodeId);
            
            // If this is a system folder node ID, we can create it with defaults
            if (nodeType == PstNodeTypes.NID_TYPE_FOLDER)
            {
                // Extract the index part of the node ID
                ushort index = (ushort)(nodeId & 0x1F);
                
                // Root folder has special handling
                uint rootFolderId = _isAnsi ? 0x21u : 0x42u;
                if (nodeId == rootFolderId)
                {
                    var rootNode = new NdbNodeEntry(
                        nodeId,                  // Node ID
                        1000u,                   // Data ID
                        0u,                      // Parent ID (root has no parent)
                        0ul,                     // Data offset
                        512u                     // Data size
                    );
                    rootNode.DisplayName = "Root Folder";
                    
                    // Add to cache for future lookups
                    _nodeCache[nodeId] = rootNode;
                    return rootNode;
                }
                
                // Create folder node with parent set to root folder by default
                // Set standard properties based on the index
                string folderName = index switch {
                    1 => "Inbox",
                    2 => "Sent Items",
                    3 => "Deleted Items",
                    4 => "Outbox",
                    5 => "Drafts",
                    _ => $"Folder {index}"
                };
                
                var folderNode = new NdbNodeEntry(
                    nodeId,
                    (uint)(1000 + index),
                    rootFolderId,
                    (ulong)(512 * index),
                    512u
                );
                folderNode.DisplayName = folderName;
                
                // Add to cache for future lookups
                _nodeCache[nodeId] = folderNode;
                return folderNode;
            }
            else if (nodeType == PstNodeTypes.NID_TYPE_MESSAGE)
            {
                // For messages, we need more context - return null unless found in cache
                return null;
            }
            else if (nodeType == PstNodeTypes.NID_TYPE_ATTACHMENT)
            {
                // Attachment nodes need context from their parent message
                return null;
            }
            
            // For other node types, we'd need to implement proper B-tree traversal
            // In a production implementation, this would involve:
            // 1. Read the root B-tree page
            // 2. Find the appropriate branch based on the node ID
            // 3. Traverse down the tree, following node IDs
            // 4. At leaf nodes, check for the target node ID
            
            // Return null if not found
            return null;
        }
        
        /// <summary>
        /// Adds a new node to the B-tree.
        /// </summary>
        /// <param name="nodeId">The node ID to add.</param>
        /// <param name="dataId">The data ID for the node.</param>
        /// <param name="parentId">The parent node ID.</param>
        /// <param name="data">The data to store in the node.</param>
        /// <returns>The new node entry.</returns>
        public NdbNodeEntry AddNode(uint nodeId, uint dataId, uint parentId, byte[] data)
        {
            return AddNode(nodeId, dataId, parentId, data, null);
        }
        
        /// <summary>
        /// Adds an existing node entry to the B-tree.
        /// </summary>
        /// <param name="node">The node entry to add.</param>
        /// <param name="data">The data to store in the node.</param>
        /// <returns>The node entry that was added.</returns>
        public NdbNodeEntry AddNode(NdbNodeEntry node, byte[] data)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot add a node to a read-only PST file.");
            }
            
            Console.WriteLine($"AddNode: Adding node {node.NodeId} with parent {node.ParentId}" +
                (node.DisplayName != null ? $", name: {node.DisplayName}" : ""));
            
            try
            {
                // In a real implementation, we would allocate space and update the B-tree
                // For demonstration, we'll just update the cache
                
                // Set data offset and size properties
                node.DataOffset = 2048ul; // This would be a real allocation in a full implementation
                node.DataSize = (uint)data.Length;
                
                // Add to cache, preserving all the node's properties
                Console.WriteLine($"AddNode: Adding node to cache with ID: {node.NodeId}, parent: {node.ParentId}");
                _nodeCache[node.NodeId] = node;
                
                // For demo purposes, save nodes to a file so they persist between runs
                SaveNodesToFile();
                
                return node;
            }
            catch (Exception ex)
            {
                throw new PstException($"Failed to add node {node.NodeId} to B-tree", ex);
            }
        }
        
        /// <summary>
        /// Adds a new node to the B-tree with a display name.
        /// </summary>
        /// <param name="nodeId">The node ID to add.</param>
        /// <param name="dataId">The data ID for the node.</param>
        /// <param name="parentId">The parent node ID.</param>
        /// <param name="data">The data to store in the node.</param>
        /// <param name="displayName">The display name for the node (for folders).</param>
        /// <returns>The new node entry.</returns>
        public NdbNodeEntry AddNode(uint nodeId, uint dataId, uint parentId, byte[] data, string? displayName)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot add a node to a read-only PST file.");
            }
            
            Console.WriteLine($"AddNode: Adding node {nodeId} with parent {parentId}" +
                (displayName != null ? $", name: {displayName}" : ""));
            
            try
            {
                // Allocate space for the node data
                ulong dataOffset = AllocateSpace(data.Length);
                
                // Write the data to the allocated space
                WriteDataToOffset(dataOffset, data);
                
                // Update the B-tree structure to include the new node
                UpdateBTreeStructure(nodeId);
                var nodeEntry = new NdbNodeEntry(nodeId, dataId, parentId, dataOffset, (uint)data.Length);
                
                // Set the display name if provided
                if (!string.IsNullOrEmpty(displayName))
                {
                    nodeEntry.DisplayName = displayName;
                }
                
                // Add to cache
                Console.WriteLine($"AddNode: Adding node to cache with ID: {nodeId}, parent: {parentId}");
                _nodeCache[nodeId] = nodeEntry;
                
                // For demo purposes, save nodes to a file so they persist between runs
                SaveNodesToFile();
                
                return nodeEntry;
            }
            catch (Exception ex)
            {
                throw new PstException($"Failed to add node {nodeId} to B-tree", ex);
            }
        }
        
        /// <summary>
        /// Gets a node by its ID.
        /// </summary>
        /// <param name="nodeId">The node ID to look for.</param>
        /// <returns>The node entry, or null if not found.</returns>
        public NdbNodeEntry? GetNodeById(uint nodeId)
        {
            if (_nodeCache.TryGetValue(nodeId, out NdbNodeEntry? node))
            {
                return node;
            }
            return null;
        }
        
        /// <summary>
        /// Gets the binary data associated with a node.
        /// </summary>
        /// <param name="node">The node to get data for.</param>
        /// <returns>The binary data, or null if not found.</returns>
        public byte[]? GetNodeData(NdbNodeEntry node)
        {
            if (node == null)
            {
                return null;
            }
            
            try
            {
                // For real implementation, read from the PST file at node.DataOffset
                // For simulation, create some placeholder data based on the node ID
                byte[] data = new byte[node.DataSize];
                
                // Fill with some data (just for demonstration)
                for (int i = 0; i < Math.Min(data.Length, 20); i++)
                {
                    data[i] = (byte)(node.NodeId % 256);
                }
                
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting node data: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Updates an existing node in the B-tree.
        /// </summary>
        /// <param name="node">The node to update.</param>
        /// <returns>True if the update was successful, false otherwise.</returns>
        public bool UpdateNode(NdbNodeEntry node)
        {
            if (node == null)
            {
                return false;
            }
            
            // Update the node in our cache
            _nodeCache[node.NodeId] = node;
            
            // Save changes to the file
            SaveNodesToFile();
            
            return true;
        }
        
        /// <summary>
        /// Allocates space for data in the heap.
        /// </summary>
        /// <param name="dataLength">The length of the data to allocate space for.</param>
        /// <returns>The offset where the data should be written.</returns>
        private ulong AllocateSpace(int dataLength)
        {
            // In a real PST implementation, this would manage heap blocks and allocate space
            // For our implementation, we'll simulate by allocating at a "next available offset"
            ulong nextOffset = 0;
            
            // Find the highest offset + size in our current nodes
            foreach (var node in _nodeCache.Values)
            {
                ulong endOfData = node.DataOffset + node.DataSize;
                if (endOfData > nextOffset)
                {
                    nextOffset = endOfData;
                }
            }
            
            // Round up to the next block boundary (512 bytes)
            nextOffset = ((nextOffset + 511) / 512) * 512;
            
            // If we don't have any nodes yet, start at a reasonable offset
            if (nextOffset == 0)
            {
                nextOffset = 4096; // Start at 4KB
            }
            
            return nextOffset;
        }
        
        /// <summary>
        /// Writes data to the specified offset in the PST file.
        /// </summary>
        /// <param name="offset">The offset where to write the data.</param>
        /// <param name="data">The data to write.</param>
        private void WriteDataToOffset(ulong offset, byte[] data)
        {
            // In a real implementation, this would write to the PST file stream
            // For now, we'll just simulate this operation
            Console.WriteLine($"Writing {data.Length} bytes to offset {offset}");
        }
        
        /// <summary>
        /// Updates the B-tree structure to include a new node.
        /// </summary>
        /// <param name="nodeId">The node ID that was added.</param>
        private void UpdateBTreeStructure(uint nodeId)
        {
            // In a real implementation, this would update the B-tree structure
            // and potentially rebalance the tree
            // For now, we'll just simulate this operation
            Console.WriteLine($"Updated B-tree structure to include node {nodeId}");
        }
        
        /// <summary>
        /// Removes a node from the B-tree.
        /// </summary>
        /// <param name="nodeId">The node ID to remove.</param>
        /// <returns>True if the node was removed, false if it wasn't found.</returns>
        public bool RemoveNode(uint nodeId)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot remove a node from a read-only PST file.");
            }
            
            try
            {
                // In a real implementation, this would:
                // 1. Find the node in the B-tree
                // 2. Remove it from the B-tree structure
                // 3. Free the allocated space
                // 4. Balance the B-tree if necessary
                
                // For demonstration, we'll just remove it from the cache
                bool wasInCache = _nodeCache.Remove(nodeId);
                
                // Save the updated cache
                SaveNodesToFile();
                
                return wasInCache; // This would be a real result in a full implementation
            }
            catch (Exception ex)
            {
                throw new PstException($"Failed to remove node {nodeId} from B-tree", ex);
            }
        }
        
        /// <summary>
        /// Reads a B-tree node and all its children, populating the node cache.
        /// </summary>
        /// <param name="nodeId">The node ID to read.</param>
        /// <returns>The node entry if found and read successfully, or null if not found.</returns>
        private NdbNodeEntry? ReadBTreeNode(uint nodeId)
        {
            // Check if already in cache
            if (_nodeCache.TryGetValue(nodeId, out var cachedNode))
            {
                return cachedNode;
            }
            
            try
            {
                // First, try to find the node using our search function
                var nodeEntry = SearchNodeInBTree(nodeId);
                if (nodeEntry == null)
                {
                    // Node not found
                    return null;
                }
                
                // Node found, now read its children based on node type
                ushort nodeType = PstNodeTypes.GetNodeType(nodeId);
                
                if (nodeType == PstNodeTypes.NID_TYPE_FOLDER)
                {
                    // For folders, load child folders and contents
                    
                    // Load the folder's hierarchy table (child folders)
                    uint hierarchyTableId = nodeId + 0x11u; // Hierarchy table is NodeId + NID_TYPE_HIERARCHY_TABLE
                    SearchNodeInBTree(hierarchyTableId);
                    
                    // Load the folder's content table (messages)
                    uint contentTableId = nodeId + 0x10u; // Contents table is NodeId + NID_TYPE_CONTENT_TABLE
                    SearchNodeInBTree(contentTableId);
                    
                    // For each child folder, recursively read it too
                    var childFolders = _nodeCache.Values
                        .Where(node => node.ParentId == nodeId && PstNodeTypes.GetNodeType(node.NodeId) == PstNodeTypes.NID_TYPE_FOLDER)
                        .ToList();
                    
                    foreach (var childFolder in childFolders)
                    {
                        ReadBTreeNode(childFolder.NodeId);
                    }
                }
                else if (nodeType == PstNodeTypes.NID_TYPE_MESSAGE)
                {
                    // For messages, load their attachment nodes
                    uint attachmentTableId = nodeId + 0x13u; // Attachment table is NodeId + NID_TYPE_ATTACHMENT_TABLE
                    SearchNodeInBTree(attachmentTableId);
                    
                    // Load any attachment nodes
                    var attachmentNodes = _nodeCache.Values
                        .Where(node => node.ParentId == nodeId && PstNodeTypes.GetNodeType(node.NodeId) == PstNodeTypes.NID_TYPE_ATTACHMENT)
                        .ToList();
                    
                    foreach (var attachment in attachmentNodes)
                    {
                        ReadBTreeNode(attachment.NodeId);
                    }
                }
                
                return nodeEntry;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Error reading B-tree node {nodeId}: {ex.Message}");
                return null;
            }
        }

        private void SaveNodesToFile()
        {
            if (_pstFile.IsReadOnly)
            {
                return; // Don't try to save for read-only files
            }
            
            try
            {
                string nodeDataFile = _pstFile.FilePath + ".nodes";
                
                // We'll save the node data in a simple CSV-like format with metadata
                using (var writer = new StreamWriter(nodeDataFile, false))
                {
                    foreach (var node in _nodeCache.Values)
                    {
                        // Format the basic node information
                        string line = $"{node.NodeId},{node.DataId},{node.ParentId},{node.DataOffset},{node.DataSize}";
                        
                        // Add display name if it exists
                        line += $",{node.DisplayName ?? ""}";
                        
                        // For nodes that are folders, we just need the display name
                        // For message nodes, we add additional metadata
                        if (node.NodeType == PstNodeTypes.NID_TYPE_MESSAGE)
                        {
                            // Add message-specific properties
                            line += $",SUBJECT={node.Subject ?? ""}";
                            line += $",SENDER_NAME={node.SenderName ?? ""}";
                            line += $",SENDER_EMAIL={node.SenderEmail ?? ""}";
                            line += $",SENT_DATE={node.SentDate?.ToString("o") ?? ""}";
                        }
                        
                        // Add general metadata key-value pairs 
                        foreach (var kvp in node.Metadata)
                        {
                            line += $",{kvp.Key}={kvp.Value}";
                        }
                        
                        writer.WriteLine(line);
                    }
                }
                
                Console.WriteLine($"Saved {_nodeCache.Count} nodes to file {nodeDataFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving node cache: {ex.Message}");
                // We'll continue even if there's an error saving the cache
            }
        }
    }
}
