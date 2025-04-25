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
                // Generate a new node ID for the folder
                // In a real implementation, this would allocate a unique ID
                // For demonstration, we'll create a mock ID
                uint newFolderId = GenerateNewFolderId();
                
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
                var bTree = _pstFile.GetNodeBTree();
                var nodeEntry = bTree.AddNode(newFolderId, 0, FolderId, folderData);
                
                // Create a new folder object
                var newFolder = new PstFolder(_pstFile, nodeEntry, this);
                
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
                
                // Add the message to the PST file
                var bTree = _pstFile.GetNodeBTree();
                var messageNode = bTree.AddNode(newMessageId, 0, FolderId, messageData);
                
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
                
                // Update the node's parent ID
                // In a real implementation, this would update the node entry in the PST file
                var bTree = _pstFile.GetNodeBTree();
                
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
                _subFolders.Clear();
                
                // Find the hierarchy table node for this folder
                var bTree = _pstFile.GetNodeBTree();
                var hierarchyTableNode = bTree.FindNodeByNid(_hierarchyTableNodeId);
                
                // Get all nodes that have this folder as a parent
                var allNodes = bTree.GetAllNodes();
                var childFolderNodes = allNodes.Where(node => 
                    node.ParentId == FolderId && 
                    node.NodeId != FolderId &&
                    (node.NodeId & 0x1Fu) == 0u);  // Folder nodes typically have specific IDs
                
                foreach (var childNode in childFolderNodes)
                {
                    // Check if we already have this folder cached
                    var cachedFolder = _pstFile.GetCachedFolder(childNode.NodeId);
                    if (cachedFolder != null)
                    {
                        _subFolders.Add(cachedFolder);
                    }
                    else
                    {
                        // Create and add the folder
                        var childFolder = new PstFolder(_pstFile, childNode, this);
                        _subFolders.Add(childFolder);
                    }
                }
                
                // If no real subfolders were found and we're the root folder,
                // add standard system folders as a fallback for demo purposes
                if (_subFolders.Count == 0 && Type == FolderType.Root)
                {
                    CreateStandardFolders();
                }
                
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
            var mockNode = new NdbNodeEntry(nodeId, 0u, FolderId, 0ul, 0u);
            var folder = new PstFolder(_pstFile, mockNode, this);
            folder.Name = name;
            folder.Type = type;
            _subFolders.Add(folder);
        }

        private void LoadMessages()
        {
            try
            {
                _messages.Clear();
                
                // Find the contents table node for this folder
                var bTree = _pstFile.GetNodeBTree();
                var contentsTableNode = bTree.FindNodeByNid(_contentsTableNodeId);
                
                if (contentsTableNode != null)
                {
                    // In a real implementation, this would:
                    // 1. Read the contents table rows
                    // 2. Extract the message node IDs
                    // 3. Load each message
                    
                    // For this implementation, we'll create mock messages
                    // based on folder type and name
                    
                    // Number of sample messages to create
                    int messageCount = 0;
                    switch (Type)
                    {
                        case FolderType.Inbox:
                            messageCount = 5;
                            break;
                        case FolderType.SentItems:
                            messageCount = 3;
                            break;
                        case FolderType.Drafts:
                            messageCount = 1;
                            break;
                        default:
                            messageCount = 0;
                            break;
                    }
                    
                    // Create sample messages
                    for (int i = 0; i < messageCount; i++)
                    {
                        uint messageId = FolderId + 0x1000u + (uint)i;
                        var mockNode = new NdbNodeEntry(messageId, 0u, FolderId, 0ul, 0u);
                        
                        // Create the message and add to our list
                        var message = new PstMessage(_pstFile, mockNode);
                        _messages.Add(message);
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
        
        private uint GenerateNewFolderId()
        {
            // In a real implementation, this would allocate a unique ID
            // For demonstration, we'll create a mock ID
            return (uint)(PstNodeTypes.NID_TYPE_FOLDER << 5) | (FolderId & 0x1Fu) + 0x1000u;
        }
        
        private uint GenerateNewMessageId()
        {
            // In a real implementation, this would allocate a unique ID
            // For demonstration, we'll create a mock ID
            return (uint)(PstNodeTypes.NID_TYPE_MESSAGE << 5) | (FolderId & 0x1Fu) + 0x2000u;
        }
        
        private byte[] SerializeFolderProperties(Dictionary<string, object> properties)
        {
            // In a real implementation, this would serialize properties to PST format
            // For demonstration, we return a placeholder byte array
            return new byte[128];
        }
        
        private void UpdateHierarchyTable(PstFolder folder)
        {
            // In a real implementation, this would add the folder to the hierarchy table
            // For this demonstration, we don't need to do anything
        }
        
        private void RemoveFromHierarchyTable(PstFolder folder)
        {
            // In a real implementation, this would remove the folder from the hierarchy table
            // For this demonstration, we don't need to do anything
        }
        
        private void UpdateContentsTable(PstMessage message)
        {
            // In a real implementation, this would add the message to the contents table
            // For this demonstration, we don't need to do anything
        }
        
        private void RemoveFromContentsTable(PstMessage message)
        {
            // In a real implementation, this would remove the message from the contents table
            // For this demonstration, we don't need to do anything
        }
        
        private void IncrementMessageCount()
        {
            // In a real implementation, this would update the message count properties
            // For this demonstration, we don't need to do anything
        }
        
        private void DecrementMessageCount(bool wasRead)
        {
            // In a real implementation, this would update the message count properties
            // For this demonstration, we don't need to do anything
        }
        
        private void UpdateHasSubFoldersProperty()
        {
            // In a real implementation, this would update the HasSubFolders property
            // For this demonstration, we'll just update the property in memory
            
            // Set the property in the property context
            _propertyContext.SetProperty(
                PstStructure.PropertyIds.PidTagSubfolders,
                PstStructure.PropertyType.PT_BOOLEAN,
                HasSubFolders
            );
            
            // Save the changes
            _propertyContext.Save();
        }
        
        #endregion
    }
}
