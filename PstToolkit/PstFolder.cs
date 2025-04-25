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

        /// <summary>
        /// Gets the unique identifier for this folder within the PST file.
        /// </summary>
        public uint FolderId => _nodeEntry.NodeId;

        /// <summary>
        /// Gets or sets the name of the folder.
        /// </summary>
        public string Name { get; set; }

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
                // In a full implementation, this would:
                // 1. Create a new node in the PST file
                // 2. Set up the folder properties
                // 3. Link it to this folder as a subfolder
                // 4. Return the new folder object
                
                // For now, we'll throw a NotImplementedException
                throw new NotImplementedException("Creating subfolders is not yet implemented");
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
                // In a full implementation, this would:
                // 1. Create a new message node in the PST file
                // 2. Copy the message content and properties
                // 3. Link it to this folder
                // 4. Add the message to the _messages list if already loaded
                
                // For now, we'll throw a NotImplementedException
                throw new NotImplementedException("Adding messages is not yet implemented");
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
                // In a full implementation, this would:
                // 1. Remove the message node from the PST file
                // 2. Update folder metadata
                // 3. Remove the message from the _messages list if loaded
                
                // For now, we'll throw a NotImplementedException
                throw new NotImplementedException("Deleting messages is not yet implemented");
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
                // In a full implementation, this would:
                // 1. Delete all messages in this folder
                // 2. Delete all subfolders recursively
                // 3. Remove this folder node from the PST file
                // 4. Update parent folder metadata
                
                // For now, we'll throw a NotImplementedException
                throw new NotImplementedException("Deleting folders is not yet implemented");
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException($"Failed to delete folder: {Name}", ex);
            }
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
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error loading folder properties", ex);
            }
        }

        private FolderType DetermineType(string containerClass)
        {
            // Determine the folder type based on the container class
            // This is a simplified implementation
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
                // In a complete implementation, this would read the hierarchy table
                // associated with this folder node to find all subfolders
                
                // For this demonstration, we'll just set _subFoldersLoaded to true
                _subFoldersLoaded = true;
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error loading subfolders", ex);
            }
        }

        private void LoadMessages()
        {
            try
            {
                // In a complete implementation, this would read the contents table
                // associated with this folder node to find all messages
                
                // For this demonstration, we'll just set _messagesLoaded to true
                _messagesLoaded = true;
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error loading messages", ex);
            }
        }
    }
}
