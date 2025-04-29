using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PstToolkit.Exceptions;
using PstToolkit.Formats;
using PstToolkit.Utils;

namespace PstToolkit
{
    /// <summary>
    /// Represents a folder in a PST file.
    /// </summary>
    public class PstFolder
    {
        private readonly PstFile _pstFile;
        private readonly NdbNodeEntry _nodeEntry;
        private readonly PropertyContext _propertyContext;
        private readonly List<PstMessage> _messages;
        private readonly List<PstFolder> _subFolders;
        private bool _messagesLoaded;
        private bool _subFoldersLoaded;
        private uint _hierarchyTableNodeId;
        private uint _contentsTableNodeId;
        private uint _associatedContentsTableNodeId;

        /// <summary>
        /// Gets the unique identifier for this folder within the PST file.
        /// </summary>
        public uint FolderId => _nodeEntry.NodeId;

        /// <summary>
        /// Gets or sets the name of the folder.
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Gets the parent folder, or null if this is a root folder.
        /// </summary>
        public PstFolder? ParentFolder { get; private set; }

        /// <summary>
        /// Gets the subfolder count.
        /// </summary>
        public int SubFolderCount => SubFolders.Count;

        /// <summary>
        /// Gets the message count.
        /// </summary>
        public int MessageCount => Messages.Count;

        /// <summary>
        /// Gets the number of unread messages in this folder.
        /// </summary>
        public int UnreadCount { get; private set; }

        /// <summary>
        /// Gets the list of subfolders.
        /// </summary>
        public IReadOnlyList<PstFolder> SubFolders
        {
            get
            {
                if (!_subFoldersLoaded)
                {
                    LoadSubFolders();
                }
                return _subFolders;
            }
        }

        /// <summary>
        /// Gets the list of messages in this folder.
        /// </summary>
        public IReadOnlyList<PstMessage> Messages
        {
            get
            {
                if (!_messagesLoaded)
                {
                    LoadMessages();
                }
                return _messages;
            }
        }
        
        /// <summary>
        /// Gets a filtered list of messages in this folder.
        /// </summary>
        /// <param name="filter">The filter to apply to the messages.</param>
        /// <returns>A filtered list of messages.</returns>
        public IEnumerable<PstMessage> GetFilteredMessages(MessageFilter filter)
        {
            if (!_messagesLoaded)
            {
                LoadMessages();
            }
            
            return filter.Apply(_messages);
        }

        /// <summary>
        /// Gets the folder type (e.g., Inbox, Sent Items, etc.).
        /// </summary>
        public FolderType Type { get; private set; }

        /// <summary>
        /// Gets whether this folder has subfolders.
        /// </summary>
        public bool HasSubFolders { get; private set; }

        /// <summary>
        /// Enumeration of standard folder types.
        /// </summary>
        public enum FolderType
        {
            /// <summary>Regular user-created folder</summary>
            Normal = 0,
            
            /// <summary>Root folder</summary>
            Root = 1,
            
            /// <summary>Inbox folder</summary>
            Inbox = 2,
            
            /// <summary>Outbox folder</summary>
            Outbox = 3,
            
            /// <summary>Sent Items folder</summary>
            SentItems = 4,
            
            /// <summary>Deleted Items folder</summary>
            DeletedItems = 5,
            
            /// <summary>Calendar folder</summary>
            Calendar = 6,
            
            /// <summary>Contacts folder</summary>
            Contacts = 7,
            
            /// <summary>Drafts folder</summary>
            Drafts = 8,
            
            /// <summary>Journal folder</summary>
            Journal = 9,
            
            /// <summary>Notes folder</summary>
            Notes = 10,
            
            /// <summary>Tasks folder</summary>
            Tasks = 11
        }

        internal PstFolder(PstFile pstFile, NdbNodeEntry nodeEntry, PstFolder? parentFolder = null)
        {
            _pstFile = pstFile;
            _nodeEntry = nodeEntry;
            _propertyContext = new PropertyContext(pstFile, nodeEntry);
            _messages = new List<PstMessage>();
            _subFolders = new List<PstFolder>();
            ParentFolder = parentFolder;
            
            // Calculate table node IDs based on this folder's node ID
            // These are typically defined by the PST format
            ushort folderNid = (ushort)(_nodeEntry.NodeId & 0x1F);
            _hierarchyTableNodeId = (uint)(PstNodeTypes.NID_TYPE_HIERARCHY_TABLE << 5) | folderNid;
            _contentsTableNodeId = (uint)(PstNodeTypes.NID_TYPE_CONTENTS_TABLE << 5) | folderNid;
            _associatedContentsTableNodeId = (uint)(PstNodeTypes.NID_TYPE_ASSOCIATED_CONTENTS << 5) | folderNid;
            
            // Register this folder in the PST file's folder cache
            pstFile.RegisterFolder(nodeEntry.NodeId, this);
            
            // Load the folder properties
            LoadProperties();
        }

        /// <summary>
        /// Creates a subfolder with the specified name.
        /// </summary>
        /// <param name="folderName">The name of the new subfolder.</param>
        /// <returns>The newly created subfolder.</returns>
        /// <exception cref="PstAccessException">Thrown if the PST file is read-only.</exception>
        public PstFolder CreateSubFolder(string folderName)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot create a subfolder in a read-only PST file.");
            }

            try
            {
                // Generate a proper node ID for the folder
                // Get the node type for folders
                ushort folderNodeType = PstNodeTypes.NID_TYPE_FOLDER;
                
                // Find the next available folder ID by checking existing folders
                uint maxFolderId = 0;
                var bTree = _pstFile.GetNodeBTree();
                var allNodes = bTree.GetAllNodes();
                var existingFolders = allNodes
                    .Where(n => PstNodeTypes.GetNodeType(n.NodeId) == PstNodeTypes.NID_TYPE_FOLDER)
                    .ToList();
                
                if (existingFolders.Any())
                {
                    // Extract the index part from each folder ID and find the highest
                    maxFolderId = existingFolders.Max(f => f.NodeId & 0x1F);
                }
                
                ushort nextFolderId = (ushort)(maxFolderId + 1);
                if (nextFolderId > 0x1F) // Ensure we stay within the 5-bit index limit
                {
                    nextFolderId = 1; // Wrap around if needed
                }
                
                // Create the final node ID by combining the node type and index
                uint newFolderId = PstNodeTypes.CreateNodeId(folderNodeType, nextFolderId);
                
                // Create folder properties
                Dictionary<string, object> properties = new Dictionary<string, object>
                {
                    { "Name", folderName },
                    { "ContentCount", 0 },
                    { "UnreadCount", 0 },
                    { "HasSubfolders", false },
                    { "ContainerClass", "IPM.Note" }
                };
                
                // Create a new node entry for the folder
                byte[] folderData = SerializeFolderProperties(properties);
                var nodeEntry = bTree.AddNode(newFolderId, 0, FolderId, folderData, folderName);
                
                // Create a new folder object
                var newFolder = new PstFolder(_pstFile, nodeEntry, this);
                
                // Set the folder name and properties
                newFolder.Name = folderName;
                
                // Update the hierarchy table to include the new folder
                UpdateHierarchyTable(newFolder);
                
                // Add to the subfolders list if already loaded
                if (_subFoldersLoaded)
                {
                    _subFolders.Add(newFolder);
                }
                
                // Update the 'has subfolders' flag
                if (!HasSubFolders)
                {
                    HasSubFolders = true;
                    UpdateHasSubFoldersProperty();
                }
                
                return newFolder;
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException($"Failed to create subfolder '{folderName}'", ex);
            }
        }

        /// <summary>
        /// Finds a subfolder by name.
        /// </summary>
        /// <param name="folderName">The name of the folder to find.</param>
        /// <param name="recursive">Whether to search recursively through all subfolders.</param>
        /// <returns>The found folder, or null if not found.</returns>
        public PstFolder? FindFolder(string folderName, bool recursive = false)
        {
            // Search for the folder in the immediate subfolders
            var folder = SubFolders.FirstOrDefault(f => 
                string.Equals(f.Name, folderName, StringComparison.OrdinalIgnoreCase));
            
            if (folder != null || !recursive)
            {
                return folder;
            }
            
            // If recursive is true and folder wasn't found, search in all subfolders
            foreach (var subFolder in SubFolders)
            {
                folder = subFolder.FindFolder(folderName, true);
                if (folder != null)
                {
                    return folder;
                }
            }
            
            return null;
        }

        /// <summary>
        /// Adds a message to this folder.
        /// </summary>
        /// <param name="message">The message to add.</param>
        /// <exception cref="PstAccessException">Thrown if the PST file is read-only.</exception>
        public void AddMessage(PstMessage message)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot add a message to a folder in a read-only PST file.");
            }

            try
            {
                // Convert the message to raw data
                byte[] messageData = message.GetRawContent();
                
                // Generate a new node ID for the message
                uint newMessageId = GenerateNewMessageId();
                
                // Prepare the node with message metadata
                var bTree = _pstFile.GetNodeBTree();
                var messageNode = new NdbNodeEntry(newMessageId, 0, FolderId, 0, (uint)messageData.Length);
                
                // Set message metadata for persistence
                messageNode.Subject = message.Subject;
                messageNode.SenderName = message.SenderName;
                messageNode.SenderEmail = message.SenderEmail;
                messageNode.SentDate = message.SentDate;
                messageNode.DisplayName = message.Subject; // Use subject as display name for messages
                
                // Store additional metadata
                messageNode.SetMetadata("BODY_TEXT", message.BodyText ?? "");
                messageNode.SetMetadata("IS_READ", message.IsRead.ToString());
                messageNode.SetMetadata("HAS_ATTACHMENTS", message.HasAttachments.ToString());
                messageNode.SetMetadata("IMPORTANCE", ((int)message.Importance).ToString());
                
                if (!string.IsNullOrEmpty(message.BodyHtml))
                {
                    messageNode.SetMetadata("BODY_HTML", message.BodyHtml);
                }
                
                // Add the node to the PST file with all metadata
                bTree.AddNode(messageNode, messageData);
                
                // Create a new PstMessage from the node
                var newMessage = new PstMessage(_pstFile, messageNode);
                
                // Update the contents table to include the new message
                UpdateContentsTable(newMessage);
                
                // Add to the messages list if already loaded
                if (_messagesLoaded)
                {
                    _messages.Add(newMessage);
                }
                
                // Update folder message count
                IncrementMessageCount();
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException("Failed to add message to folder", ex);
            }
        }

        /// <summary>
        /// Deletes a message from this folder.
        /// </summary>
        /// <param name="message">The message to delete.</param>
        /// <exception cref="PstAccessException">Thrown if the PST file is read-only.</exception>
        public void DeleteMessage(PstMessage message)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot delete a message from a folder in a read-only PST file.");
            }

            try
            {
                // Remove the message node from the PST file
                var bTree = _pstFile.GetNodeBTree();
                bool removed = bTree.RemoveNode(message.MessageId);
                
                if (!removed)
                {
                    throw new PstException($"Failed to remove message node {message.MessageId}");
                }
                
                // Update the contents table to remove the message
                RemoveFromContentsTable(message);
                
                // Remove from the messages list if loaded
                if (_messagesLoaded)
                {
                    _messages.Remove(message);
                }
                
                // Update folder message count
                DecrementMessageCount(message.IsRead);
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException("Failed to delete message from folder", ex);
            }
        }

        /// <summary>
        /// Deletes this folder and all its contents.
        /// </summary>
        /// <exception cref="PstAccessException">Thrown if the PST file is read-only or if the folder is the root folder.</exception>
        public void Delete()
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot delete a folder in a read-only PST file.");
            }
            
            if (ParentFolder == null)
            {
                throw new PstAccessException("Cannot delete the root folder of a PST file.");
            }

            try
            {
                // First, ensure subfolders are loaded
                if (!_subFoldersLoaded)
                {
                    LoadSubFolders();
                }
                
                // Delete all subfolders recursively
                foreach (var subFolder in _subFolders.ToList()) // Use ToList to create a copy
                {
                    subFolder.Delete();
                }
                
                // Delete all messages in this folder
                if (!_messagesLoaded)
                {
                    LoadMessages();
                }
                
                foreach (var message in _messages.ToList()) // Use ToList to create a copy
                {
                    DeleteMessage(message);
                }
                
                // Remove from parent's hierarchy table
                if (ParentFolder != null)
                {
                    ParentFolder.RemoveFromHierarchyTable(this);
                }
                
                // Remove the folder node from the B-tree
                var bTree = _pstFile.GetNodeBTree();
                bool removed = bTree.RemoveNode(FolderId);
                
                if (!removed)
                {
                    throw new PstException($"Failed to remove folder node {FolderId}");
                }
                
                // Remove from parent's subfolder list if loaded
                if (ParentFolder != null && ParentFolder._subFoldersLoaded)
                {
                    ParentFolder._subFolders.Remove(this);
                }
                
                // Update parent's HasSubFolders property if needed
                if (ParentFolder != null && ParentFolder._subFolders.Count == 0)
                {
                    ParentFolder.HasSubFolders = false;
                    ParentFolder.UpdateHasSubFoldersProperty();
                }
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException($"Failed to delete folder: {Name}", ex);
            }
        }

        /// <summary>
        /// Moves this folder to become a subfolder of the specified target folder.
        /// </summary>
        /// <param name="targetFolder">The folder that will become this folder's new parent.</param>
        /// <exception cref="PstAccessException">Thrown if the PST file is read-only or if the folder is the root folder.</exception>
        /// <exception cref="ArgumentException">Thrown if the target folder is this folder or a subfolder of this folder.</exception>
        public void MoveTo(PstFolder targetFolder)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot move a folder in a read-only PST file.");
            }
            
            if (ParentFolder == null)
            {
                throw new PstAccessException("Cannot move the root folder of a PST file.");
            }
            
            if (targetFolder == this)
            {
                throw new ArgumentException("Cannot move a folder to itself.");
            }
            
            // Check if target is a subfolder of this folder (which would create a cycle)
            PstFolder? parent = targetFolder;
            while (parent != null)
            {
                if (parent == this)
                {
                    throw new ArgumentException("Cannot move a folder to one of its subfolders.");
                }
                parent = parent.ParentFolder;
            }

            try
            {
                // Remove from current parent's hierarchy table
                ParentFolder.RemoveFromHierarchyTable(this);
                
                // Update the node's parent ID in the PST file structure
                var bTree = _pstFile.GetNodeBTree();
                
                // Get the current node entry
                var nodeEntry = bTree.GetNodeById(FolderId);
                if (nodeEntry == null)
                {
                    throw new PstException($"Failed to find folder node {FolderId} in B-tree");
                }
                
                // Update the parent ID in the node entry
                nodeEntry.ParentId = targetFolder.FolderId;
                
                // Update the node in the B-tree
                bTree.UpdateNode(nodeEntry);
                
                // Remove from current parent's subfolder list if loaded
                if (ParentFolder._subFoldersLoaded)
                {
                    ParentFolder._subFolders.Remove(this);
                }
                
                // Update current parent's HasSubFolders property if needed
                if (ParentFolder._subFolders.Count == 0)
                {
                    ParentFolder.HasSubFolders = false;
                    ParentFolder.UpdateHasSubFoldersProperty();
                }
                
                // Add to new parent's hierarchy table
                targetFolder.UpdateHierarchyTable(this);
                
                // Add to new parent's subfolder list if loaded
                if (targetFolder._subFoldersLoaded)
                {
                    targetFolder._subFolders.Add(this);
                }
                
                // Update new parent's HasSubFolders property if needed
                if (!targetFolder.HasSubFolders)
                {
                    targetFolder.HasSubFolders = true;
                    targetFolder.UpdateHasSubFoldersProperty();
                }
                
                // Update this folder's parent reference
                ParentFolder = targetFolder;
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException($"Failed to move folder '{Name}'", ex);
            }
        }

        /// <summary>
        /// Gets special system folders (Inbox, Sent Items, etc.) if they exist.
        /// </summary>
        /// <param name="folderType">The type of system folder to find.</param>
        /// <returns>The system folder if found, or null if not found.</returns>
        public PstFolder? GetSystemFolder(FolderType folderType)
        {
            // This requires a root folder
            if (Type != FolderType.Root)
            {
                throw new InvalidOperationException("System folders can only be accessed from the root folder.");
            }
            
            // Load subfolders if not already loaded
            if (!_subFoldersLoaded)
            {
                LoadSubFolders();
            }
            
            // Look for folders with matching types
            return _subFolders.FirstOrDefault(f => f.Type == folderType);
        }

        private void LoadProperties()
        {
            try
            {
                // Folder name is stored in property 0x3001 (PidTagDisplayName)
                Name = _propertyContext.GetString(0x3001) ?? "Unnamed Folder";
                
                // Folder type is stored in property 0x3613 (PidTagContainerClass)
                var containerClass = _propertyContext.GetString(0x3613) ?? "";
                Type = DetermineType(containerClass);
                
                // Content count is stored in property 0x3602 (PidTagContentCount)
                var contentCount = _propertyContext.GetInt32(0x3602);
                if (contentCount.HasValue)
                {
                    // This is just a cached value. Real count is determined by loading messages.
                }
                
                // Unread count is stored in property 0x3603 (PidTagContentUnreadCount)
                var unreadCount = _propertyContext.GetInt32(0x3603);
                UnreadCount = unreadCount ?? 0;
                
                // Has subfolders is stored in property 0x360A (PidTagSubfolders)
                var hasSubfolders = _propertyContext.GetBoolean(0x360A);
                HasSubFolders = hasSubfolders ?? false;
                
                // If this is the root folder and no type is determined from the container class,
                // explicitly set it to Root
                if (ParentFolder == null && Type == FolderType.Normal)
                {
                    Type = FolderType.Root;
                }
                
                // For certain standard folders, set the type based on the name if not already set
                if (Type == FolderType.Normal && ParentFolder?.Type == FolderType.Root)
                {
                    Type = Name.ToLowerInvariant() switch
                    {
                        "inbox" => FolderType.Inbox,
                        "outbox" => FolderType.Outbox,
                        "sent items" => FolderType.SentItems,
                        "deleted items" => FolderType.DeletedItems,
                        "calendar" => FolderType.Calendar,
                        "contacts" => FolderType.Contacts,
                        "drafts" => FolderType.Drafts,
                        "journal" => FolderType.Journal,
                        "notes" => FolderType.Notes,
                        "tasks" => FolderType.Tasks,
                        _ => FolderType.Normal
                    };
                }
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error loading folder properties", ex);
            }
        }

        private FolderType DetermineType(string containerClass)
        {
            // Determine the folder type based on the container class
            if (string.IsNullOrEmpty(containerClass))
            {
                return FolderType.Normal;
            }
            
            return containerClass.ToLowerInvariant() switch
            {
                "ipm.note" => FolderType.Normal,
                "ipm.note.outlookspamfolder" => FolderType.Normal,
                "ipm.appointment" => FolderType.Calendar,
                "ipm.contact" => FolderType.Contacts,
                "ipm.stickynote" => FolderType.Notes,
                "ipm.task" => FolderType.Tasks,
                "ipm.journal" => FolderType.Journal,
                _ => FolderType.Normal
            };
        }

        private void LoadSubFolders()
        {
            try
            {
                Console.WriteLine($"Loading subfolders for folder: {Name} (ID: {FolderId})");
                _subFolders.Clear();
                
                // Find the hierarchy table node for this folder
                var bTree = _pstFile.GetNodeBTree();
                var hierarchyTableNode = bTree.FindNodeByNid(_hierarchyTableNodeId);
                
                // Get all nodes that have this folder as a parent
                var allNodes = bTree.GetAllNodes();
                Console.WriteLine($"Total nodes in tree: {allNodes.Count}");
                
                // Find all nodes that have this folder as parent
                var childFolderNodes = allNodes.Where(node => 
                    node.ParentId == FolderId && 
                    node.NodeId != FolderId).ToList();
                    
                Console.WriteLine($"Found {childFolderNodes.Count} potential child nodes for parent {FolderId}");
                
                // If the list is empty, try the legacy ID mask filter as a fallback
                if (childFolderNodes.Count == 0)
                {
                    childFolderNodes = allNodes.Where(node => 
                        node.ParentId == FolderId && 
                        node.NodeId != FolderId &&
                        (node.NodeId & 0x1Fu) == 0u).ToList();  // Folder nodes typically have specific IDs
                }
                
                Console.WriteLine($"Found {childFolderNodes.Count} child folder nodes for parent {FolderId}");
                
                foreach (var childNode in childFolderNodes)
                {
                    Console.WriteLine($"Processing child node {childNode.NodeId} with parent {childNode.ParentId}");
                    
                    // Check if we already have this folder cached
                    var cachedFolder = _pstFile.GetCachedFolder(childNode.NodeId);
                    if (cachedFolder != null)
                    {
                        Console.WriteLine($"Using cached folder: {cachedFolder.Name}");
                        _subFolders.Add(cachedFolder);
                    }
                    else
                    {
                        // Create and add the folder
                        Console.WriteLine($"Creating new folder from node {childNode.NodeId}");
                        var childFolder = new PstFolder(_pstFile, childNode, this);
                        
                        // Set the display name if it exists in the node
                        if (!string.IsNullOrEmpty(childNode.DisplayName))
                        {
                            childFolder.Name = childNode.DisplayName;
                            Console.WriteLine($"Setting folder name to: {childNode.DisplayName}");
                        }
                        
                        _subFolders.Add(childFolder);
                    }
                }
                
                // We used to create standard folders as fallback, but for production
                // we should only use what's actually in the PST file.
                // If root folder has no subfolders, that's valid in certain PST files.
                
                _subFoldersLoaded = true;
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error loading subfolders", ex);
            }
        }

        private void CreateStandardFolders()
        {
            // Create standard system folders that typically exist in PST files
            CreateSystemFolder(0x122u, "Inbox", FolderType.Inbox);
            CreateSystemFolder(0x222u, "Sent Items", FolderType.SentItems);
            CreateSystemFolder(0x322u, "Deleted Items", FolderType.DeletedItems);
            CreateSystemFolder(0x422u, "Outbox", FolderType.Outbox);
            CreateSystemFolder(0x522u, "Drafts", FolderType.Drafts);
            CreateSystemFolder(0x622u, "Calendar", FolderType.Calendar);
            CreateSystemFolder(0x722u, "Contacts", FolderType.Contacts);
            CreateSystemFolder(0x822u, "Tasks", FolderType.Tasks);
            CreateSystemFolder(0x922u, "Notes", FolderType.Notes);
        }

        private void CreateSystemFolder(uint nodeId, string name, FolderType type)
        {
            // Instead of using a mock node, we'll create a proper one and add it to the B-tree
            var dataId = nodeId + 1000u; // Create unique data ID
            var nodeEntry = new NdbNodeEntry(nodeId, dataId, FolderId, 0ul, 512u);
            nodeEntry.DisplayName = name;
            
            // Add the folder node to the node B-tree
            var bTree = _pstFile.GetNodeBTree();
            
            // Create folder properties
            Dictionary<string, object> properties = new Dictionary<string, object>
            {
                { "DisplayName", name },
                { "FolderType", (int)type }
            };
            
            // Serialize properties to a byte array (in real implementation)
            byte[] data = SerializeFolderProperties(properties);
            
            // Add the node to the B-tree
            bTree.AddNode(nodeEntry, data);
            
            // Create folder object and add to collection
            var folder = new PstFolder(_pstFile, nodeEntry, this);
            folder.Name = name;
            folder.Type = type;
            _subFolders.Add(folder);
        }

        private void LoadMessages()
        {
            try
            {
                _messages.Clear();
                
                // Find messages in the node cache that have this folder as parent
                var bTree = _pstFile.GetNodeBTree();
                var allNodes = bTree.GetAllNodes();
                
                // Only log detailed messages for small folders/files
                bool verboseLogging = allNodes.Count < 100;
                
                if (verboseLogging)
                {
                    Console.WriteLine($"Looking for messages in folder: {Name} (ID: {FolderId})");
                    Console.WriteLine($"Scanning {allNodes.Count} nodes for messages with parent ID {FolderId}");
                }
                
                // First, gather all message nodes that belong to this folder
                var messageNodes = allNodes
                    .Where(node => node.NodeType == PstNodeTypes.NID_TYPE_MESSAGE && node.ParentId == FolderId)
                    .ToList();
                
                // Process in batches for better performance with large message sets
                const int batchSize = 50;
                int realMessagesFound = 0;
                
                for (int i = 0; i < messageNodes.Count; i++)
                {
                    var node = messageNodes[i];
                    
                    try
                    {
                        // Only log details for the first few messages or small message sets
                        if (verboseLogging && (realMessagesFound < 10 || messageNodes.Count < 20))
                        {
                            Console.WriteLine($"Found message node {node.NodeId} with parent {node.ParentId}, name: {node.DisplayName ?? "(no name)"}");
                        }
                        
                        // Create a message object from the node
                        var message = new PstMessage(_pstFile, node);
                        
                        // Add it to the messages list
                        _messages.Add(message);
                        realMessagesFound++;
                        
                        // For large message sets, provide periodic status updates
                        if (messageNodes.Count > batchSize*2 && realMessagesFound % batchSize == 0)
                        {
                            Console.WriteLine($"Processed {realMessagesFound}/{messageNodes.Count} messages in folder '{Name}'...");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Only log first few errors to avoid console spam
                        if (realMessagesFound < 5)
                        {
                            Console.WriteLine($"Error loading message {node.NodeId}: {ex.Message}");
                        }
                    }
                }
                
                // If we found real messages, use those
                if (realMessagesFound > 0)
                {
                    Console.WriteLine($"Loaded {realMessagesFound} real messages for folder '{Name}' from node cache");
                    _messagesLoaded = true;
                    return;
                }
                
                // As a fallback, check the contents table 
                // (this is how it would work in a real PST file)
                var contentsTableNode = bTree.FindNodeByNid(_contentsTableNodeId);
                if (contentsTableNode != null)
                {
                    Console.WriteLine($"No messages found directly. Checking contents table for folder '{Name}'");
                    
                    // Read the contents table rows and load messages
                    var contentsData = bTree.GetNodeData(contentsTableNode);
                    if (contentsData != null && contentsData.Length > 0)
                    {
                        // Parse the content table data to extract message IDs
                        var messageIds = ParseContentsTable(contentsData);
                        
                        Console.WriteLine($"Found {messageIds.Count} messages in contents table for folder '{Name}'");
                        
                        // Load each message by its ID
                        foreach (var messageId in messageIds)
                        {
                            try 
                            {
                                var messageNode = bTree.GetNodeById(messageId);
                                if (messageNode != null)
                                {
                                    var message = new PstMessage(_pstFile, messageNode);
                                    _messages.Add(message);
                                    realMessagesFound++;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error loading message {messageId}: {ex.Message}");
                            }
                        }
                    }
                }
                
                _messagesLoaded = true;
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error loading messages", ex);
            }
        }
        
        #region Private Helper Methods
        
        // Counter to ensure unique IDs
        private static uint _nextFolderId = 0x1000u;
        
        private uint GenerateNewFolderId()
        {
            // This method allocates a unique folder ID in compliance with PST specification
            // Folder node IDs follow the PST storage format requirements:
            // - The high bits (27-31) contain the node type (NID_TYPE_FOLDER for folders)
            // - The low bits (0-26) contain a unique identifier
            
            // Get the PST node B-tree to check for existing nodes
            var bTree = _pstFile.GetNodeBTree();
            
            // Node ID counter for generating unique IDs (initialized from static field)
            uint nodeCounter = _nextFolderId;
            
            // This will hold our new unique node ID
            uint newId;
            
            // Make sure we find an ID that isn't already in use
            int maxAttempts = 1000;  // Prevent infinite loops
            int attempts = 0;
            
            do {
                // Create a new ID with the folder node type in the high bits (27-31)
                // and a unique counter in the remaining bits
                // NID format: TTTTTNNN NNNNNNNN NNNNNNNN NNNNNNNN
                // where T = Type bits, N = Node ID bits
                
                // We use a 5-bit field for the node type according to PST specs
                // For folder, we use NID_TYPE_FOLDER (0x1)
                newId = (uint)(PstNodeTypes.NID_TYPE_FOLDER << 5) | (nodeCounter & 0x1FFFFF);
                
                // Increment counter for next attempt
                nodeCounter++;
                attempts++;
                
                // Check if this ID already exists in the B-tree
                // We only create the folder if the ID is unique
            } while (bTree.FindNodeByNid(newId) != null && attempts < maxAttempts);
            
            // If we hit max attempts, fall back to using a time-based unique ID
            if (attempts >= maxAttempts)
            {
                // This is a fallback method that should rarely be needed
                // We create a mostly-unique ID based on ticks
                uint tickBasedId = (uint)(DateTime.Now.Ticks & 0x1FFFFF);
                newId = (uint)(PstNodeTypes.NID_TYPE_FOLDER << 5) | tickBasedId;
            }
            
            // Update the static field for the next folder
            _nextFolderId = nodeCounter;
            
            Console.WriteLine($"Generated new folder ID: {newId}");
            return newId;
        }
        
        private uint GenerateNewMessageId()
        {
            // Generate a unique message node ID
            // Message node IDs have the NID_TYPE_MESSAGE type (shifted left by 5 bits)
            // followed by a unique ID based on the current folder
            
            // Get all existing message IDs in this folder
            var bTree = _pstFile.GetNodeBTree();
            var allNodes = bTree.GetAllNodes();
            var existingMessageIds = new HashSet<uint>();
            
            // Find existing message IDs for this folder
            foreach (var node in allNodes)
            {
                if (PstNodeTypes.GetNodeType(node.NodeId) == PstNodeTypes.NID_TYPE_MESSAGE && 
                    node.ParentId == FolderId)
                {
                    existingMessageIds.Add(node.NodeId);
                }
            }
            
            // Start with a base ID using the current time to ensure uniqueness
            uint baseId = (uint)((DateTime.Now.Ticks & 0xFFFF) << 8);
            
            // Create a unique message ID by combining:
            // 1. The message node type (shifted left by 5 bits)
            // 2. A portion of the folder ID for context (lower 5 bits)
            // 3. A unique counter to avoid collisions
            uint counter = 0;
            uint newMessageId;
            
            do
            {
                // Create a unique ID pattern
                newMessageId = (uint)(PstNodeTypes.NID_TYPE_MESSAGE << 5) | 
                               ((baseId + counter) & 0x1Fu);
                counter++;
            } while (existingMessageIds.Contains(newMessageId) && counter < 0xFFFF);
            
            // If we've exhausted the standard range, use a time-based approach
            if (counter >= 0xFFFF)
            {
                // Use millisecond precision for uniqueness
                newMessageId = (uint)(PstNodeTypes.NID_TYPE_MESSAGE << 5) | 
                               (uint)(DateTime.Now.Millisecond & 0x1F);
            }
            
            Console.WriteLine($"Generated new message ID: {newMessageId} for folder {FolderId}");
            return newMessageId;
        }
        
        private byte[] SerializeFolderProperties(Dictionary<string, object> properties)
        {
            // Serialize properties dictionary to a byte array in a format that can be 
            // stored in the PST file and later reconstructed
            
            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var writer = new PstBinaryWriter(memoryStream, _pstFile.IsAnsi))
                    {
                        // Write the number of properties
                        writer.Write((int)properties.Count);
                        
                        // Write each property
                        foreach (var property in properties)
                        {
                            // Determine the encoding based on the PST file type
                            var encoding = _pstFile.IsAnsi ? System.Text.Encoding.ASCII : System.Text.Encoding.Unicode;
                            
                            // Write property name
                            writer.WriteString(property.Key, encoding);
                            
                            // Write property type marker based on value type
                            if (property.Value is string stringValue)
                            {
                                writer.Write((byte)1); // String type
                                writer.WriteString(stringValue, encoding);
                            }
                            else if (property.Value is int intValue)
                            {
                                writer.Write((byte)2); // Integer type
                                writer.Write(intValue);
                            }
                            else if (property.Value is bool boolValue)
                            {
                                writer.Write((byte)3); // Boolean type
                                writer.Write(boolValue);
                            }
                            else if (property.Value is DateTime dateValue)
                            {
                                writer.Write((byte)4); // DateTime type
                                writer.Write(dateValue.ToBinary());
                            }
                            else if (property.Value is byte[] byteArrayValue)
                            {
                                writer.Write((byte)5); // ByteArray type
                                writer.Write(byteArrayValue.Length);
                                writer.Write(byteArrayValue);
                            }
                            else
                            {
                                // For other types, convert to string
                                writer.Write((byte)1); // String type
                                writer.WriteString(property.Value?.ToString() ?? string.Empty, encoding);
                            }
                        }
                    }
                    
                    // Return the serialized data
                    return memoryStream.ToArray();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error serializing folder properties: {ex.Message}");
                
                // In case of error, return a minimal valid data structure
                using (var memoryStream = new MemoryStream())
                {
                    using (var writer = new BinaryWriter(memoryStream))
                    {
                        writer.Write(0); // No properties
                    }
                    return memoryStream.ToArray();
                }
            }
        }
        
        private void UpdateHierarchyTable(PstFolder folder)
        {
            // This method updates the PST hierarchy table with the folder information
            // The hierarchy table tracks parent-child relationships between folders
            
            if (folder == null)
            {
                return;
            }
            
            try
            {
                // Get the PST file's root hierarchy node
                var bTree = _pstFile.GetNodeBTree();
                // Use the RootFolder property of PstFile instead of a constant
                var rootNode = bTree.FindNodeByNid(_pstFile.RootFolder.FolderId);
                
                if (rootNode == null)
                {
                    Console.WriteLine("Warning: Could not find root folder node for hierarchy update");
                    return;
                }
                
                // Create hierarchy table entry properties - serialize to a string representation
                // that can be stored in the folder metadata
                var properties = new Dictionary<string, string>
                {
                    { "ParentEntryId", folder.ParentFolder?.FolderId.ToString() ?? "0" },
                    { "EntryId", folder.FolderId.ToString() },
                    { "DisplayName", folder.Name ?? "Unnamed Folder" },
                    { "HasSubfolders", folder.HasSubFolders.ToString() },
                    { "ContentCount", folder.MessageCount.ToString() },
                    { "ContainerClass", "IPM.Note" } // Default container class for folders
                };
                
                // Convert properties to a serialized representation that can be stored
                string propertiesString = string.Join(";", properties.Select(p => $"{p.Key}={p.Value}"));
                byte[] hierarchyData = System.Text.Encoding.UTF8.GetBytes(propertiesString);
                
                // Create a unique hierarchy table entry ID
                uint hierarchyEntryId = (uint)(PstNodeTypes.NID_TYPE_HIERARCHY_TABLE << 5) | 
                                        (folder.FolderId & 0x1Fu);
                
                // Add entry to hierarchy table
                var dataId = hierarchyEntryId + 1u;
                string displayName = $"Hierarchy: {folder.Name} ({folder.FolderId})";
                
                // Add or update the node in the B-tree - use the appropriate method signature
                bTree.AddNode(hierarchyEntryId, dataId, rootNode.NodeId, hierarchyData, displayName);
                
                Console.WriteLine($"Updated hierarchy table entry for folder: {folder.Name}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating hierarchy table: {ex.Message}");
            }
        }
        
        private void RemoveFromHierarchyTable(PstFolder folder)
        {
            // This method removes a folder entry from the PST hierarchy table
            
            if (folder == null)
            {
                return;
            }
            
            try
            {
                // Get the B-tree from the PST file
                var bTree = _pstFile.GetNodeBTree();
                
                // Calculate the hierarchy table entry ID for this folder
                uint hierarchyEntryId = (uint)(PstNodeTypes.NID_TYPE_HIERARCHY_TABLE << 5) | 
                                        (folder.FolderId & 0x1Fu);
                
                // Try to remove the node using its ID
                bool removed = bTree.RemoveNode(hierarchyEntryId);
                
                if (removed)
                {
                    // The node was successfully removed
                    Console.WriteLine($"Removed folder {folder.Name} (ID: {folder.FolderId}) from hierarchy table");
                }
                else
                {
                    Console.WriteLine($"Warning: Could not find hierarchy entry for folder {folder.Name} (ID: {folder.FolderId})");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error removing folder from hierarchy table: {ex.Message}");
            }
        }
        
        private void UpdateContentsTable(PstMessage message)
        {
            // This method updates the PST contents table with a message entry
            // The contents table tracks messages within a folder
            
            if (message == null)
            {
                return;
            }
            
            try
            {
                // Get the B-tree from the PST file
                var bTree = _pstFile.GetNodeBTree();
                
                // Create a serialized string representation of the key message properties
                var properties = new Dictionary<string, string>
                {
                    { "ParentEntryId", FolderId.ToString() },
                    { "EntryId", message.MessageId.ToString() },
                    { "MessageSize", message.Size.ToString() },
                    { "LastModifiedTime", message.SentDate.ToString("o") },  // Use SentDate as LastModifiedTime
                    { "ReceivedTime", message.ReceivedDate.ToString("o") },  // Use ReceivedDate
                    { "Subject", message.Subject ?? "" },
                    { "SenderName", message.SenderName ?? "" },
                    { "SenderEmail", message.SenderEmail ?? "" },  // Use SenderEmail property
                    { "IsRead", message.IsRead.ToString() },
                    { "HasAttachments", message.HasAttachments.ToString() }
                };
                
                // Convert properties to a serialized representation
                string propertiesString = string.Join(";", properties.Select(p => $"{p.Key}={p.Value}"));
                byte[] contentData = System.Text.Encoding.UTF8.GetBytes(propertiesString);
                
                // Create a unique contents table entry ID based on the folder and message
                uint contentsEntryId = (uint)(PstNodeTypes.NID_TYPE_CONTENTS_TABLE << 5) | 
                                       (message.MessageId & 0x1Fu);
                
                // Create data node ID (must be unique and related to the entry ID)
                var dataId = contentsEntryId + 1u;
                
                // Add the node using the method signature available in BTreeOnHeap
                bTree.AddNode(contentsEntryId, dataId, FolderId, contentData, message.Subject);
                
                // Update the message count property for the folder
                IncrementMessageCount();
                
                Console.WriteLine($"Added message {message.Subject} to contents table");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating contents table: {ex.Message}");
            }
        }
        
        private void RemoveFromContentsTable(PstMessage message)
        {
            // This method removes a message entry from the PST contents table
            
            if (message == null)
            {
                return;
            }
            
            try
            {
                // Get the B-tree from the PST file
                var bTree = _pstFile.GetNodeBTree();
                
                // Calculate the contents table entry ID for this message
                uint contentsEntryId = (uint)(PstNodeTypes.NID_TYPE_CONTENTS_TABLE << 5) | 
                                      (message.MessageId & 0x1Fu);
                
                // Use the RemoveNode method with the node ID
                bool removed = bTree.RemoveNode(contentsEntryId);
                
                if (removed)
                {
                    // Update folder message count
                    DecrementMessageCount(message.IsRead);
                    
                    Console.WriteLine($"Removed message {message.Subject} (ID: {message.MessageId}) from contents table");
                }
                else
                {
                    Console.WriteLine($"Warning: Could not find contents entry for message {message.Subject} (ID: {message.MessageId})");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error removing message from contents table: {ex.Message}");
            }
        }
        
        private void IncrementMessageCount()
        {
            // This method updates the message count properties when a message is added
            
            try
            {
                // Increment total message count
                int currentCount = MessageCount;
                currentCount++;
                
                // Update the property in the property context
                _propertyContext.SetProperty(
                    PstStructure.PropertyIds.PidTagContentCount,
                    PstStructure.PropertyType.PT_LONG,
                    currentCount
                );
                
                // Note: Cannot directly update MessageCount as it's a read-only property
                // Instead, we add the message to the internal list which will update the property
                // when it's accessed next time
                
                // Save the property context changes
                _propertyContext.Save();
                
                // Also update the folder node properties to reflect this change
                var bTree = _pstFile.GetNodeBTree();
                var folderNode = bTree.FindNodeByNid(FolderId);
                
                if (folderNode != null)
                {
                    // Instead of directly modifying the node, we'll update its metadata
                    folderNode.SetMetadata("ContentCount", currentCount.ToString());
                    
                    // Use the correct node adding/updating method
                    // Create serialized representation of the node's data
                    string propertiesString = string.Join(";", 
                        folderNode.Metadata.Select(p => $"{p.Key}={p.Value}"));
                    byte[] folderData = System.Text.Encoding.UTF8.GetBytes(propertiesString);
                    
                    // Re-add the node with updated data
                    bTree.AddNode(folderNode.NodeId, folderNode.DataId, folderNode.ParentId, 
                                  folderData, folderNode.DisplayName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error incrementing message count: {ex.Message}");
            }
        }
        
        private void DecrementMessageCount(bool wasRead)
        {
            // This method updates the message count properties when a message is removed
            
            try
            {
                // Get current counts
                int currentTotalCount = MessageCount;
                int currentUnreadCount = UnreadCount;
                
                // Only decrement if we have messages
                if (currentTotalCount > 0)
                {
                    currentTotalCount--;
                    
                    // Update total message count property
                    _propertyContext.SetProperty(
                        PstStructure.PropertyIds.PidTagContentCount,
                        PstStructure.PropertyType.PT_LONG,
                        currentTotalCount
                    );
                    
                    // Note: Cannot directly update MessageCount as it's a read-only property
                    // Instead, we remove the message from the internal list which will update the property
                    // when it's accessed next time
                    
                    // If the message was unread, also decrement unread count
                    if (!wasRead && currentUnreadCount > 0)
                    {
                        currentUnreadCount--;
                        
                        // Update unread count property
                        _propertyContext.SetProperty(
                            PstStructure.PropertyIds.PidTagContentUnreadCount,
                            PstStructure.PropertyType.PT_LONG,
                            currentUnreadCount
                        );
                        
                        // Note: Cannot directly update UnreadCount as it might be a read-only property
                    }
                    
                    // Save the property context changes
                    _propertyContext.Save();
                    
                    // Also update the folder node properties
                    var bTree = _pstFile.GetNodeBTree();
                    var folderNode = bTree.FindNodeByNid(FolderId);
                    
                    if (folderNode != null)
                    {
                        // Instead of directly modifying the node, we'll update its metadata
                        folderNode.SetMetadata("ContentCount", currentTotalCount.ToString());
                        
                        // Create serialized representation of the node's data
                        string propertiesString = string.Join(";", 
                            folderNode.Metadata.Select(p => $"{p.Key}={p.Value}"));
                        byte[] folderData = System.Text.Encoding.UTF8.GetBytes(propertiesString);
                        
                        // Re-add the node with updated data
                        bTree.AddNode(folderNode.NodeId, folderNode.DataId, folderNode.ParentId, 
                                      folderData, folderNode.DisplayName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error decrementing message count: {ex.Message}");
            }
        }
        
        private void UpdateHasSubFoldersProperty()
        {
            // Update the HasSubFolders property in the folder's property context and PST file
            
            // Set the property in the property context
            _propertyContext.SetProperty(
                PstStructure.PropertyIds.PidTagSubfolders,
                PstStructure.PropertyType.PT_BOOLEAN,
                HasSubFolders
            );
            
            // Also update the node metadata in the B-tree
            try
            {
                // Get the node for this folder
                var bTree = _pstFile.GetNodeBTree();
                var folderNode = bTree.GetNodeById(FolderId);
                
                if (folderNode != null)
                {
                    // Set the HasSubFolders metadata
                    folderNode.SetMetadata("HAS_SUBFOLDERS", HasSubFolders.ToString());
                    
                    // Update the node in the B-tree
                    bTree.UpdateNode(folderNode);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Error updating HasSubFolders property: {ex.Message}");
            }
            
            // Save the changes
            _propertyContext.Save();
        }
        
        /// <summary>
        /// Parses the contents table binary data to extract message IDs.
        /// </summary>
        /// <param name="contentsData">The binary contents table data.</param>
        /// <returns>A list of message node IDs.</returns>
        private List<uint> ParseContentsTable(byte[] contentsData)
        {
            List<uint> messageIds = new List<uint>();
            
            try
            {
                using (var memoryStream = new MemoryStream(contentsData))
                using (var reader = new BinaryReader(memoryStream))
                {
                    // Read the count of entries in the contents table
                    int entryCount = 0;
                    
                    // Attempt to read entry count
                    if (memoryStream.Length >= 4)
                    {
                        entryCount = reader.ReadInt32();
                    }
                    
                    // Validate entry count is reasonable
                    if (entryCount < 0 || entryCount > 100000) // Sanity check
                    {
                        Console.WriteLine($"Warning: Invalid content table entry count: {entryCount}, using safe value");
                        entryCount = Math.Min(Math.Max(0, (int)(memoryStream.Length - 4) / 8), 1000);
                    }
                    
                    // Read each entry
                    for (int i = 0; i < entryCount && memoryStream.Position + 4 <= memoryStream.Length; i++)
                    {
                        // Read a message ID
                        uint messageId = reader.ReadUInt32();
                        
                        // Skip additional data (if any) for this entry
                        if (memoryStream.Position + 4 <= memoryStream.Length)
                        {
                            int additionalDataSize = reader.ReadInt32();
                            
                            // Validate the size is reasonable
                            if (additionalDataSize >= 0 && 
                                additionalDataSize <= 10000 && 
                                memoryStream.Position + additionalDataSize <= memoryStream.Length)
                            {
                                // Skip the additional data
                                memoryStream.Position += additionalDataSize;
                            }
                            else
                            {
                                // Invalid data size, stop processing
                                break;
                            }
                        }
                        
                        // Store the message ID if it's valid
                        if (messageId != 0 && PstNodeTypes.IsMessage(messageId))
                        {
                            messageIds.Add(messageId);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing contents table: {ex.Message}");
            }
            
            return messageIds;
        }
        
        #endregion
    }
}
