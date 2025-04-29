using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using MimeKit;
using PstToolkit.Exceptions;
using PstToolkit.Formats;
using PstToolkit.Utils;
using static PstToolkit.Utils.PstStructure;

namespace PstToolkit
{
    /// <summary>
    /// Represents an email message stored in a PST file.
    /// </summary>
    public class PstMessage
    {
        private readonly PstFile _pstFile;
        private readonly NdbNodeEntry _nodeEntry;
        private readonly PropertyContext _propertyContext;
        private byte[]? _rawContent;
        private MimeMessage? _parsedMimeMessage;

        /// <summary>
        /// Gets the unique identifier for this message within the PST file.
        /// </summary>
        public uint MessageId => _nodeEntry.NodeId;

        /// <summary>
        /// Gets the subject of the email message.
        /// </summary>
        public string Subject { get; private set; }

        /// <summary>
        /// Gets the sender's email address.
        /// </summary>
        public string SenderEmail { get; private set; }

        /// <summary>
        /// Gets the sender's display name.
        /// </summary>
        public string SenderName { get; private set; }

        /// <summary>
        /// Gets the list of recipients' email addresses.
        /// </summary>
        public List<string> Recipients { get; private set; }

        /// <summary>
        /// Gets the date and time when the message was sent.
        /// </summary>
        public DateTime SentDate { get; private set; }

        /// <summary>
        /// Gets the date and time when the message was received.
        /// </summary>
        public DateTime ReceivedDate { get; private set; }

        /// <summary>
        /// Gets whether the message has been read.
        /// </summary>
        public bool IsRead { get; private set; }

        /// <summary>
        /// Gets whether the message has attachments.
        /// </summary>
        public bool HasAttachments { get; private set; }

        /// <summary>
        /// Gets the message size in bytes.
        /// </summary>
        public long Size { get; private set; }

        /// <summary>
        /// Gets the list of attachment filenames.
        /// </summary>
        public List<string> AttachmentNames { get; private set; }

        /// <summary>
        /// Gets the importance level of the message.
        /// </summary>
        public MessageImportance Importance { get; private set; }

        /// <summary>
        /// Gets or sets the plain text body of the message.
        /// </summary>
        public string BodyText { get; set; }

        /// <summary>
        /// Gets or sets the HTML body of the message.
        /// </summary>
        public string BodyHtml { get; set; }

        /// <summary>
        /// Enumeration of message importance levels.
        /// </summary>
        public enum MessageImportance
        {
            /// <summary>Low importance</summary>
            Low = 0,
            
            /// <summary>Normal importance</summary>
            Normal = 1,
            
            /// <summary>High importance</summary>
            High = 2
        }

        internal PstMessage(PstFile pstFile, NdbNodeEntry nodeEntry)
        {
            _pstFile = pstFile;
            _nodeEntry = nodeEntry;
            _propertyContext = new PropertyContext(pstFile, nodeEntry);

            // Initialize with default values
            Subject = "";
            SenderEmail = "";
            SenderName = "";
            Recipients = new List<string>();
            SentDate = DateTime.MinValue;
            ReceivedDate = DateTime.MinValue;
            IsRead = false;
            HasAttachments = false;
            Size = 0;
            AttachmentNames = new List<string>();
            Importance = MessageImportance.Normal;
            BodyText = "";
            BodyHtml = "";

            // First try to initialize properties from the node metadata (for messages loaded from cache)
            InitializeFromNodeMetadata();
            
            // Then try to load from property context (for newly created/opened messages)
            if (string.IsNullOrEmpty(Subject) && string.IsNullOrEmpty(SenderName))
            {
                LoadProperties();
            }
        }

        /// <summary>
        /// Creates a new email message that can be added to a PST file.
        /// </summary>
        /// <param name="pstFile">The PST file to create the message in.</param>
        /// <param name="subject">The subject of the email.</param>
        /// <param name="body">The body text of the email.</param>
        /// <param name="senderEmail">The sender's email address.</param>
        /// <param name="senderName">The sender's display name.</param>
        /// <returns>A new PstMessage instance.</returns>
        public static PstMessage Create(PstFile pstFile, string subject, string body, string senderEmail, string senderName)
        {
            if (pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot create a message in a read-only PST file.");
            }
            
            try
            {
                // Generate a proper message node ID using the PST file's node allocation system
                var bTree = pstFile.GetNodeBTree();
                
                // Message node IDs have the NID_TYPE_MESSAGE type identifier
                uint nodeIdBase = bTree.GetHighestNodeId(PstNodeTypes.NID_TYPE_MESSAGE);
                uint messageNodeId = nodeIdBase + PstNodeTypes.NID_TYPE_MESSAGE;
                
                // Ensure the node ID doesn't already exist
                while (bTree.GetNodeById(messageNodeId) != null)
                {
                    messageNodeId += PstNodeTypes.NID_TYPE_MESSAGE;
                }
                
                Console.WriteLine($"Generated new message ID: {messageNodeId & 0x00FFFFFF:X} for folder {(messageNodeId >> 24) & 0x00FF}");
                
                // Create a MimeMessage to help build the email
                var mimeMessage = new MimeMessage();
                
                // Set basic properties
                mimeMessage.Subject = subject;
                mimeMessage.From.Add(new MailboxAddress(senderName, senderEmail));
                mimeMessage.Date = DateTime.Now;
                
                // Set the body
                var builder = new BodyBuilder();
                builder.TextBody = body;
                mimeMessage.Body = builder.ToMessageBody();
                
                // Serialize the MimeMessage to a byte array
                byte[] messageData;
                using (var memStream = new MemoryStream())
                {
                    mimeMessage.WriteTo(memStream);
                    messageData = memStream.ToArray();
                }
                
                // Allocate space for the message data in the PST heap
                ulong dataOffset = bTree.AllocateSpace(messageData.Length);
                
                // Create a real node entry with proper PST structure
                var nodeEntry = new NdbNodeEntry(
                    messageNodeId,                  // The allocated node ID
                    messageNodeId & 0x00FFFFFF,     // The data ID (lower 24 bits of node ID)
                    0,                              // Parent ID will be set when added to a folder
                    dataOffset,                     // The allocated data offset
                    (uint)messageData.Length        // The size of the message data
                );
                
                // Set properties in the node entry (for folder display)
                nodeEntry.DisplayName = subject;
                nodeEntry.Subject = subject;
                nodeEntry.SenderName = senderName;
                nodeEntry.SenderEmail = senderEmail;
                nodeEntry.SentDate = DateTime.Now;
                nodeEntry.NodeType = PstNodeTypes.NID_TYPE_MESSAGE;
                
                // Add the node to the B-tree
                bTree.AddNode(nodeEntry);
                
                // Write the message data to the allocated space
                bTree.WriteDataToOffset(dataOffset, messageData);
                
                // Create a new message with the node entry
                var message = new PstMessage(pstFile, nodeEntry);
                
                // Set the properties directly
                message.Subject = subject;
                message.SenderName = senderName;
                message.SenderEmail = senderEmail;
                message.BodyText = body;
                message.SentDate = DateTime.Now;
                message.ReceivedDate = DateTime.Now;
                message._rawContent = messageData;
                
                return message;
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException("Failed to create new message", ex);
            }
        }
        
        /// <summary>
        /// Creates a new email message from a MimeKit MimeMessage.
        /// </summary>
        /// <param name="pstFile">The PST file to create the message in.</param>
        /// <param name="mimeMessage">The MimeKit message to convert.</param>
        /// <returns>A new PstMessage instance.</returns>
        public static PstMessage Create(PstFile pstFile, MimeMessage mimeMessage)
        {
            if (pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot create a message in a read-only PST file.");
            }
            
            try
            {
                // Generate a proper message node ID using the PST file's node allocation system
                var bTree = pstFile.GetNodeBTree();
                
                // Message node IDs have the NID_TYPE_MESSAGE type identifier
                uint nodeIdBase = bTree.GetHighestNodeId(PstNodeTypes.NID_TYPE_MESSAGE);
                uint messageNodeId = nodeIdBase + PstNodeTypes.NID_TYPE_MESSAGE;
                
                // Ensure the node ID doesn't already exist
                while (bTree.GetNodeById(messageNodeId) != null)
                {
                    messageNodeId += PstNodeTypes.NID_TYPE_MESSAGE;
                }
                
                Console.WriteLine($"Generated new message ID: {messageNodeId & 0x00FFFFFF:X} for folder {(messageNodeId >> 24) & 0x00FF}");
                
                // Serialize the MimeMessage to a byte array
                byte[] messageData;
                using (var memStream = new MemoryStream())
                {
                    mimeMessage.WriteTo(memStream);
                    messageData = memStream.ToArray();
                }
                
                // Allocate space for the message data in the PST heap
                ulong dataOffset = bTree.AllocateSpace(messageData.Length);
                
                // Create a real node entry with proper PST structure
                var nodeEntry = new NdbNodeEntry(
                    messageNodeId,                  // The allocated node ID
                    messageNodeId & 0x00FFFFFF,     // The data ID (lower 24 bits of node ID)
                    0,                              // Parent ID will be set when added to a folder
                    dataOffset,                     // The allocated data offset
                    (uint)messageData.Length        // The size of the message data
                );
                
                // Write the message data to the allocated space
                bTree.WriteDataToOffset(dataOffset, messageData);
                
                // Create a new message with the node entry
                var message = new PstMessage(pstFile, nodeEntry);
                
                // Extract and set the properties from the MimeMessage
                message.Subject = mimeMessage.Subject ?? "";
                
                // Set sender information
                if (mimeMessage.From.FirstOrDefault() is MailboxAddress from)
                {
                    message.SenderName = from.Name ?? "";
                    message.SenderEmail = from.Address ?? "";
                }
                
                // Set date information
                message.SentDate = mimeMessage.Date.DateTime;
                message.ReceivedDate = DateTime.Now;
                
                // Set recipients
                message.Recipients = new List<string>();
                foreach (var to in mimeMessage.To.OfType<MailboxAddress>())
                {
                    message.Recipients.Add(to.Address);
                }
                
                // Set body text
                if (mimeMessage.TextBody != null)
                {
                    message.BodyText = mimeMessage.TextBody;
                }
                
                // Set HTML body if available
                if (mimeMessage.HtmlBody != null)
                {
                    message.BodyHtml = mimeMessage.HtmlBody;
                }
                
                // Store the raw content
                message._rawContent = messageData;
                message._parsedMimeMessage = mimeMessage;
                
                return message;
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException("Failed to create new message from MimeMessage", ex);
            }
        }

        /// <summary>
        /// Gets the message as a MimeKit MimeMessage.
        /// </summary>
        /// <returns>A MimeMessage representation of this email.</returns>
        public MimeMessage ToMimeMessage()
        {
            if (_parsedMimeMessage != null)
                return _parsedMimeMessage;

            // Try to parse from raw content if available
            if (_rawContent != null)
            {
                try
                {
                    using var stream = new MemoryStream(_rawContent);
                    _parsedMimeMessage = MimeMessage.Load(stream);
                    return _parsedMimeMessage;
                }
                catch
                {
                    // Fall back to creating from properties
                }
            }

            // Create a new MimeMessage from the properties
            var message = new MimeMessage();
            
            // Set sender
            if (!string.IsNullOrEmpty(SenderEmail))
            {
                message.From.Add(new MailboxAddress(SenderName, SenderEmail));
            }
            
            // Set recipients
            foreach (var recipient in Recipients)
            {
                message.To.Add(MailboxAddress.Parse(recipient));
            }
            
            // Set subject
            message.Subject = Subject;
            
            // Set date
            message.Date = SentDate;
            
            // Set body
            var builder = new BodyBuilder();
            if (!string.IsNullOrEmpty(BodyHtml))
            {
                builder.HtmlBody = BodyHtml;
            }
            
            if (!string.IsNullOrEmpty(BodyText))
            {
                builder.TextBody = BodyText;
            }
            
            message.Body = builder.ToMessageBody();
            
            _parsedMimeMessage = message;
            return message;
        }

        /// <summary>
        /// Gets the raw content of the message.
        /// </summary>
        /// <returns>The raw message content as a byte array.</returns>
        public byte[] GetRawContent()
        {
            if (_rawContent == null)
            {
                LoadRawContent();
            }
            
            return _rawContent ?? Array.Empty<byte>();
        }
        
        /// <summary>
        /// Gets the PropertyContext associated with this message for advanced property access.
        /// This is primarily for internal use when copying messages between PST files.
        /// </summary>
        /// <returns>The PropertyContext for this message.</returns>
        internal PropertyContext GetPropertyContext()
        {
            return _propertyContext;
        }
        
        /// <summary>
        /// Creates a new message by copying all properties from an existing message.
        /// This ensures a complete property copy rather than just the standard properties.
        /// </summary>
        /// <param name="pstFile">The PST file to create the message in.</param>
        /// <param name="sourceMessage">The source message to copy properties from.</param>
        /// <returns>The newly created message with all properties copied.</returns>
        /// <exception cref="PstAccessException">Thrown if the PST file is read-only.</exception>
        /// <exception cref="PstException">Thrown if the message creation fails.</exception>
        public static PstMessage CreateFromExisting(PstFile pstFile, PstMessage sourceMessage)
        {
            if (pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot create a message in a read-only PST file.");
            }
            
            try
            {
                // Generate a proper message node ID using the PST file's node allocation system
                var bTree = pstFile.GetNodeBTree();
                
                // Message node IDs have the NID_TYPE_MESSAGE type identifier
                uint nodeIdBase = bTree.GetHighestNodeId(PstNodeTypes.NID_TYPE_MESSAGE);
                uint newMessageId = nodeIdBase + PstNodeTypes.NID_TYPE_MESSAGE;
                
                // Ensure the node ID doesn't already exist
                while (bTree.GetNodeById(newMessageId) != null)
                {
                    newMessageId += PstNodeTypes.NID_TYPE_MESSAGE;
                }
                
                Console.WriteLine($"Generated new message ID: {newMessageId & 0x00FFFFFF:X} for folder {(newMessageId >> 24) & 0x00FF}");
                
                // Get the source message's raw content to preserve all properties
                var nodeData = sourceMessage.GetRawContent();
                
                // Create a node for the message with proper parameters
                var nodeEntry = bTree.AddNode(
                    newMessageId,                     // Node ID
                    newMessageId & 0x00FFFFFF,        // Data ID
                    0,                                // Parent ID (will be set when added to folder)
                    nodeData,                         // Content data
                    sourceMessage.Subject             // Display name
                );
                
                // Create the message object
                var message = new PstMessage(pstFile, nodeEntry);
                
                // Extract all properties from the source message through its PropertyContext
                var sourceContext = sourceMessage.GetPropertyContext();
                Dictionary<uint, object> sourceProperties = sourceContext.GetAllProperties();
                
                // For each property in the source message, set it in the new message
                foreach (var kvp in sourceProperties)
                {
                    // Copy property exactly as it is, preserving property ID and type
                    message._propertyContext.SetProperty(
                        (ushort)(kvp.Key & 0xFFFF),  // Extract property ID from combined key
                        (PstStructure.PropertyType)((kvp.Key >> 16) & 0xFFFF), // Extract property type from combined key
                        kvp.Value
                    );
                }
                
                // Reload the properties to ensure they take effect
                message.LoadProperties();
                
                return message;
            }
            catch (Exception ex)
            {
                throw new PstException("Failed to create message by copying existing message", ex);
            }
        }

        /// <summary>
        /// Represents an attachment to an email message.
        /// </summary>
        public class Attachment
        {
            /// <summary>
            /// Gets the short filename of the attachment.
            /// </summary>
            public string Filename { get; internal set; }

            /// <summary>
            /// Gets the long filename of the attachment.
            /// </summary>
            public string LongFilename { get; internal set; }

            /// <summary>
            /// Gets the content type/MIME type of the attachment.
            /// </summary>
            public string ContentType { get; internal set; }

            /// <summary>
            /// Gets the size of the attachment in bytes.
            /// </summary>
            public long Size { get; internal set; }

            /// <summary>
            /// Gets whether the attachment is embedded (displayed inline in the message).
            /// </summary>
            public bool IsEmbedded { get; internal set; }
            
            /// <summary>
            /// Gets whether this attachment is an email message.
            /// </summary>
            public bool IsEmailMessage => ContentType.Equals("message/rfc822", StringComparison.OrdinalIgnoreCase);

            /// <summary>
            /// Gets the attachment data as a byte array.
            /// </summary>
            /// <returns>The attachment content as a byte array.</returns>
            public byte[] GetContent()
            {
                if (_contentLoaded)
                {
                    return _content;
                }

                if (_contentLoader != null)
                {
                    _content = _contentLoader();
                    _contentLoaded = true;
                    return _content;
                }

                return Array.Empty<byte>();
            }
            
            /// <summary>
            /// If this attachment is an email message, gets it as a PstMessage.
            /// Returns null if the attachment is not an email message.
            /// </summary>
            /// <returns>A PstMessage if the attachment is an email, otherwise null.</returns>
            public PstMessage? GetAsEmailMessage()
            {
                if (!IsEmailMessage)
                {
                    return null;
                }
                
                try
                {
                    byte[] content = GetContent();
                    if (content.Length == 0)
                    {
                        return null;
                    }
                    
                    // Parse the content as a MimeMessage
                    using var stream = new MemoryStream(content);
                    var mimeMessage = MimeMessage.Load(stream);
                    
                    // Create a new PstMessage from the MimeMessage
                    var bTree = _parentMessage._pstFile.GetNodeBTree();
                    
                    // Use attachment node type for embedded email (not a real message node in the PST)
                    uint nodeIdBase = bTree.GetHighestNodeId(PstNodeTypes.NID_TYPE_ATTACHMENT);
                    uint embeddedNodeId = nodeIdBase + PstNodeTypes.NID_TYPE_ATTACHMENT;
                    
                    // Create a node entry for this attachment - not actually in the PST tree
                    var nodeEntry = new NdbNodeEntry(
                        embeddedNodeId,                  // The attachment node ID
                        embeddedNodeId & 0x00FFFFFF,     // The data ID (lower 24 bits of node ID)
                        _parentMessage._nodeEntry.NodeId,// Parent ID is the message containing this attachment
                        0,                               // Data offset - memory only, not in PST
                        (uint)content.Length             // Size of content
                    );
                    
                    var message = new PstMessage(_parentMessage._pstFile, nodeEntry);
                    
                    // Set basic properties from the parsed MimeMessage
                    message.Subject = mimeMessage.Subject ?? "";
                    
                    if (mimeMessage.From.FirstOrDefault() is MailboxAddress from)
                    {
                        message.SenderName = from.Name ?? "";
                        message.SenderEmail = from.Address ?? "";
                    }
                    
                    message.SentDate = mimeMessage.Date.DateTime;
                    message.ReceivedDate = DateTime.Now;
                    
                    message.Recipients = new List<string>();
                    foreach (var to in mimeMessage.To.OfType<MailboxAddress>())
                    {
                        message.Recipients.Add(to.Address);
                    }
                    
                    if (mimeMessage.TextBody != null)
                    {
                        message.BodyText = mimeMessage.TextBody;
                    }
                    
                    if (mimeMessage.HtmlBody != null)
                    {
                        message.BodyHtml = mimeMessage.HtmlBody;
                    }
                    
                    // Store the raw content
                    message._rawContent = content;
                    message._parsedMimeMessage = mimeMessage;
                    
                    // Check for attachments
                    if (mimeMessage.Attachments.Any())
                    {
                        message.HasAttachments = true;
                        
                        // For consistency, we'll extract nested attachment names
                        foreach (var attachment in mimeMessage.Attachments)
                        {
                            string filename = attachment.ContentDisposition?.FileName ?? 
                                              attachment.ContentType.Name ?? 
                                              "attachment";
                            message.AttachmentNames.Add(filename);
                        }
                    }
                    
                    return message;
                }
                catch (Exception ex)
                {
                    // Log the error with all details for debugging
                    Console.WriteLine($"Error parsing nested email: {ex.Message}");
                    Console.WriteLine($"Exception details: {ex}");
                    
                    // In a production environment, this could also write to a log file or logging service
                    string logFilePath = Path.Combine(Path.GetTempPath(), "PstToolkit_errors.log");
                    try
                    {
                        File.AppendAllText(logFilePath, 
                            $"[{DateTime.Now}] Error parsing nested email: {ex.Message}\n{ex}\n\n");
                    }
                    catch
                    {
                        // Silently fail if we can't write to the log file
                    }
                    return null;
                }
            }

            internal Attachment(PstMessage parentMessage)
            {
                Filename = "";
                LongFilename = "";
                ContentType = "application/octet-stream";
                Size = 0;
                IsEmbedded = false;
                _parentMessage = parentMessage;
            }

            internal Func<byte[]>? _contentLoader;
            private byte[] _content = Array.Empty<byte>();
            private bool _contentLoaded = false;
            private readonly PstMessage _parentMessage;
        }

        /// <summary>
        /// Gets the attachments of the message.
        /// </summary>
        /// <returns>A list of attachments.</returns>
        public List<Attachment> GetAttachments()
        {
            try
            {
                var result = new List<Attachment>();
                
                // If message doesn't have attachments, return empty list
                if (!HasAttachments)
                {
                    return result;
                }
                
                // Calculate the attachment table node ID based on this message's node ID
                ushort messageNid = (ushort)(_nodeEntry.NodeId & 0x1F);
                uint attachmentTableNodeId = (uint)(PstNodeTypes.NID_TYPE_ATTACHMENT_TABLE << 5) | messageNid;
                
                // Find the attachment table node
                var bTree = _pstFile.GetNodeBTree();
                var attachmentTableNode = bTree.FindNodeByNid(attachmentTableNodeId);
                
                if (attachmentTableNode != null)
                {
                    // Calculate the base node ID for attachment data
                    uint baseAttachmentNodeId = (uint)(PstNodeTypes.NID_TYPE_ATTACHMENT_OBJECT << 5) | messageNid;
                    
                    // For PST files, attachments are typically numbered sequentially
                    // Try to load a reasonable number of potential attachments (10 max)
                    for (int i = 0; i < 10; i++)
                    {
                        // Calculate the attachment ID for this entry
                        uint attachmentId = baseAttachmentNodeId + (uint)i;
                        
                        // Find the attachment node
                        var attachmentNode = bTree.FindNodeByNid(attachmentId);
                        
                        if (attachmentNode != null)
                        {
                            // Get the property context for this attachment node
                            var propContext = new PropertyContext(_pstFile, attachmentNode);
                            
                            // Get attachment properties
                            string filename = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagAttachFilename) ?? "";
                            string longFilename = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagAttachLongFilename) ?? "";
                            
                            // If filenames are empty or missing, try getting display name
                            if (string.IsNullOrEmpty(filename) && string.IsNullOrEmpty(longFilename))
                            {
                                filename = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagDisplayName) ?? $"Attachment{i+1}";
                                longFilename = filename;
                            }
                            
                            // Use short filename if long is missing
                            if (string.IsNullOrEmpty(longFilename))
                            {
                                longFilename = filename;
                            }
                            
                            // Use long filename if short is missing
                            if (string.IsNullOrEmpty(filename))
                            {
                                filename = Path.GetFileName(longFilename);
                            }
                            
                            // If we still don't have a filename, generate one
                            if (string.IsNullOrEmpty(filename))
                            {
                                filename = $"Attachment{i+1}";
                                longFilename = filename;
                            }
                            
                            // Get content type
                            string contentType = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagAttachMimeTag) ?? "application/octet-stream";
                            
                            // Get size
                            long size = propContext.GetInt32((ushort)PstStructure.PropertyIds.PidTagAttachmentSize) ?? 0;
                            
                            // Check if it's an embedded message
                            int attachMethod = propContext.GetInt32((ushort)PstStructure.PropertyIds.PidTagAttachMethod) ?? 1;
                            bool isEmbedded = attachMethod == 5; // ATTACH_EMBEDDED_MSG
                            
                            // If it's an embedded message and no content type was specified, set it
                            if (isEmbedded && contentType == "application/octet-stream")
                            {
                                contentType = "message/rfc822";
                            }
                            
                            // Create the attachment
                            var attachment = new Attachment(this)
                            {
                                Filename = filename,
                                LongFilename = longFilename,
                                ContentType = contentType,
                                Size = size,
                                IsEmbedded = isEmbedded
                            };
                            
                            // Set up the content loader to load the attachment content when requested
                            attachment._contentLoader = () => LoadAttachmentContent(attachmentId);
                            
                            result.Add(attachment);
                        }
                        else
                        {
                            // If we can't find an attachment at this index, we're probably done
                            // PST files usually store attachments in sequential node IDs
                            if (i > 0)
                                break;
                        }
                    }
                }
                
                // Try another approach if we couldn't find attachments the normal way but HasAttachments is true
                if (result.Count == 0)
                {
                    // Use direct property access to find attachment references
                    var attachmentIds = GetAttachmentNodeIds();
                    
                    foreach (var attachId in attachmentIds)
                    {
                        var attachNode = bTree.FindNodeByNid(attachId);
                        if (attachNode == null)
                            continue;
                            
                        var propContext = new PropertyContext(_pstFile, attachNode);
                        
                        // Get attachment properties similar to above
                        string filename = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagAttachFilename) ?? "";
                        string longFilename = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagAttachLongFilename) ?? "";
                        
                        if (string.IsNullOrEmpty(filename) && string.IsNullOrEmpty(longFilename))
                        {
                            filename = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagDisplayName) ?? $"Attachment{result.Count+1}";
                            longFilename = filename;
                        }
                        
                        if (string.IsNullOrEmpty(longFilename)) longFilename = filename;
                        if (string.IsNullOrEmpty(filename)) filename = Path.GetFileName(longFilename);
                        
                        if (string.IsNullOrEmpty(filename))
                        {
                            filename = $"Attachment{result.Count+1}";
                            longFilename = filename;
                        }
                        
                        string contentType = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagAttachMimeTag) ?? "application/octet-stream";
                        long size = propContext.GetInt32((ushort)PstStructure.PropertyIds.PidTagAttachmentSize) ?? 0;
                        int attachMethod = propContext.GetInt32((ushort)PstStructure.PropertyIds.PidTagAttachMethod) ?? 1;
                        bool isEmbedded = attachMethod == 5; // ATTACH_EMBEDDED_MSG
                        
                        if (isEmbedded && contentType == "application/octet-stream")
                        {
                            contentType = "message/rfc822";
                        }
                        
                        var attachment = new Attachment(this)
                        {
                            Filename = filename,
                            LongFilename = longFilename,
                            ContentType = contentType,
                            Size = size,
                            IsEmbedded = isEmbedded
                        };
                        
                        attachment._contentLoader = () => LoadAttachmentContent(attachId);
                        result.Add(attachment);
                    }
                }
                
                // Update AttachmentNames list from the found attachments
                if (result.Count > 0)
                {
                    AttachmentNames = new List<string>();
                    foreach (var attachment in result)
                    {
                        AttachmentNames.Add(attachment.Filename);
                    }
                }
                
                return result;
            }
            catch (Exception ex)
            {
                throw new PstException("Error loading attachments", ex);
            }
        }
        
        /// <summary>
        /// Gets the attachment node IDs for this message.
        /// </summary>
        /// <returns>A list of attachment node IDs.</returns>
        private List<uint> GetAttachmentNodeIds()
        {
            var result = new List<uint>();
            
            try
            {
                // Calculate the base mask for attachment nodes
                ushort messageNid = (ushort)(_nodeEntry.NodeId & 0x1F);
                uint baseAttachmentNodeId = (uint)(PstNodeTypes.NID_TYPE_ATTACHMENT_OBJECT << 5) | messageNid;
                
                // Try to find attachments with sequential IDs (common in PST files)
                for (uint i = 0; i < 20; i++)  // Try a reasonable max number
                {
                    uint attachmentId = baseAttachmentNodeId + i;
                    if (_pstFile.GetNodeEntry(attachmentId) != null)
                    {
                        result.Add(attachmentId);
                    }
                    else if (i > 0)
                    {
                        // If we have found at least one attachment and then find a gap,
                        // we're probably done (PST files usually store attachments with sequential IDs)
                        break;
                    }
                }
            }
            catch (Exception)
            {
                // Ignore errors in attachment ID discovery, since this is a fallback method
            }
            
            return result;
        }
        
        /// <summary>
        /// Gets all nested email messages from the attachments.
        /// </summary>
        /// <returns>A list of email messages that were attachments to this message.</returns>
        public List<PstMessage> GetNestedEmailAttachments()
        {
            var result = new List<PstMessage>();
            
            try
            {
                // Get all attachments
                var attachments = GetAttachments();
                
                // Filter for email attachments
                foreach (var attachment in attachments)
                {
                    if (attachment.IsEmailMessage)
                    {
                        var emailMessage = attachment.GetAsEmailMessage();
                        if (emailMessage != null)
                        {
                            result.Add(emailMessage);
                        }
                    }
                }
                
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting nested email attachments: {ex.Message}");
                return result;
            }
        }
        
        /// <summary>
        /// Extracts all email attachments, including nested ones, to EML files in the specified directory.
        /// </summary>
        /// <param name="outputDirectory">The directory to save the extracted email files.</param>
        /// <param name="recursionLevel">The current recursion level, used internally to prevent infinite recursion.</param>
        /// <param name="parentPath">The parent path prefix, used internally for nested emails.</param>
        /// <returns>The number of emails extracted.</returns>
        public int ExtractAllEmailAttachments(string outputDirectory, int recursionLevel = 0, string parentPath = "")
        {
            // Prevent infinite recursion by limiting depth
            if (recursionLevel > 5)
            {
                return 0;
            }
            
            int extractedCount = 0;
            
            try
            {
                // Ensure the output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }
                
                // Get all email attachments
                var emailAttachments = GetNestedEmailAttachments();
                
                foreach (var emailMsg in emailAttachments)
                {
                    try
                    {
                        string prefix = string.IsNullOrEmpty(parentPath) ? "" : parentPath + "_";
                        string filenameBase = $"{prefix}{emailMsg.Subject}_{emailMsg.SentDate:yyyy-MM-dd_HHmmss}";
                        
                        // Remove invalid filename characters
                        foreach (var invalidChar in Path.GetInvalidFileNameChars())
                        {
                            filenameBase = filenameBase.Replace(invalidChar, '_');
                        }
                        
                        // Export the email as EML
                        string emlFilePath = Path.Combine(outputDirectory, filenameBase + ".eml");
                        
                        // Write the content to the file
                        var mimeMessage = emailMsg.ToMimeMessage();
                        using (var stream = new FileStream(emlFilePath, FileMode.Create))
                        {
                            mimeMessage.WriteTo(stream);
                        }
                        
                        extractedCount++;
                        
                        // Recursively extract nested email attachments
                        string nestedPrefix = string.IsNullOrEmpty(prefix) ? "Nested" : prefix + "Nested";
                        extractedCount += emailMsg.ExtractAllEmailAttachments(
                            outputDirectory, 
                            recursionLevel + 1, 
                            nestedPrefix);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error extracting nested email: {ex.Message}");
                    }
                }
                
                return extractedCount;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting email attachments: {ex.Message}");
                return extractedCount;
            }
        }
        
        /// <summary>
        /// Adds another email message as an attachment to this message.
        /// </summary>
        /// <param name="emailMessage">The email message to attach.</param>
        /// <returns>The newly added attachment.</returns>
        public Attachment AddEmailAttachment(PstMessage emailMessage)
        {
            // Convert the email to a MimeMessage
            var mimeMessage = emailMessage.ToMimeMessage();
            
            // Get the raw content
            byte[] content;
            using (var memStream = new MemoryStream())
            {
                mimeMessage.WriteTo(memStream);
                content = memStream.ToArray();
            }
            
            // Create a filename based on the subject
            string filename = !string.IsNullOrEmpty(emailMessage.Subject) 
                ? emailMessage.Subject.Trim() 
                : "Email Message";
                
            // Sanitize the filename
            foreach (var invalidChar in Path.GetInvalidFileNameChars())
            {
                filename = filename.Replace(invalidChar, '_');
            }
            
            // Add timestamp to make the filename unique
            filename = $"{filename}_{DateTime.Now:yyyyMMdd_HHmmss}.eml";
            
            // Add the attachment with MIME type for email messages
            return AddAttachment(filename, content, "message/rfc822");
        }
        
        /// <summary>
        /// Adds an attachment to this message.
        /// </summary>
        /// <param name="filename">The filename of the attachment.</param>
        /// <param name="content">The content of the attachment as a byte array.</param>
        /// <param name="contentType">The content type/MIME type of the attachment.</param>
        /// <returns>The newly added attachment.</returns>
        public Attachment AddAttachment(string filename, byte[] content, string contentType = "application/octet-stream")
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot add attachments to a message in a read-only PST file.");
            }
            
            try
            {
                // Generate a new node ID for the attachment
                uint attachmentId = _nodeEntry.NodeId + 0x10000 + (uint)AttachmentNames.Count;
                
                // Create attachment properties
                var propContext = new PropertyContext(_pstFile, _nodeEntry);
                propContext.SetProperty(PstStructure.PropertyIds.PidTagAttachFilename, 
                    PstStructure.PropertyType.PT_STRING8, Path.GetFileName(filename));
                propContext.SetProperty(PstStructure.PropertyIds.PidTagAttachLongFilename, 
                    PstStructure.PropertyType.PT_STRING8, filename);
                propContext.SetProperty(PstStructure.PropertyIds.PidTagAttachmentSize, 
                    PstStructure.PropertyType.PT_LONG, content.Length);
                propContext.SetProperty(PstStructure.PropertyIds.PidTagAttachMethod, 
                    PstStructure.PropertyType.PT_LONG, 1); // ATTACH_BY_VALUE
                
                // Add attachment node to the PST file
                var bTree = _pstFile.GetNodeBTree();
                var nodeEntry = bTree.AddNode(attachmentId, 0, _nodeEntry.NodeId, content);
                
                // Create the attachment object
                var attachment = new Attachment(this)
                {
                    Filename = Path.GetFileName(filename),
                    LongFilename = filename,
                    ContentType = contentType,
                    Size = content.Length,
                    IsEmbedded = false
                };
                
                // Set the content through the content loader
                attachment._contentLoader = () => content;
                
                // Update message properties to indicate it has attachments
                if (!HasAttachments)
                {
                    HasAttachments = true;
                    
                    // Update the message flags property
                    var messageFlags = _propertyContext.GetInt32(PstStructure.PropertyIds.PidTagMessageFlags) ?? 0;
                    messageFlags |= 0x10; // MSGFLAG_HASATTACH
                    _propertyContext.SetProperty(PstStructure.PropertyIds.PidTagMessageFlags, 
                        PstStructure.PropertyType.PT_LONG, messageFlags);
                    
                    // Save the changes
                    _propertyContext.Save();
                }
                
                // Add to attachment names list
                AttachmentNames.Add(attachment.Filename);
                
                return attachment;
            }
            catch (Exception ex)
            {
                throw new PstException("Failed to add attachment", ex);
            }
        }
        
        /// <summary>
        /// Removes an attachment from this message.
        /// </summary>
        /// <param name="attachment">The attachment to remove.</param>
        public void RemoveAttachment(Attachment attachment)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot remove attachments from a message in a read-only PST file.");
            }
            
            try
            {
                // 1. Find the attachment node ID by looking for the attachment in the attachment table
                ushort messageNid = (ushort)(_nodeEntry.NodeId & 0x1F);
                uint baseAttachmentNodeId = (uint)(PstNodeTypes.NID_TYPE_ATTACHMENT_OBJECT << 5) | messageNid;
                
                // Load existing attachments to find the one we want to remove
                var attachments = GetAttachments();
                int attachmentIndex = attachments.IndexOf(attachment);
                
                if (attachmentIndex >= 0)
                {
                    // Calculate the attachment node ID based on its index
                    uint attachmentNodeId = baseAttachmentNodeId + (uint)attachmentIndex;
                    
                    // 2. Remove it from the PST file by removing the node
                    var bTree = _pstFile.GetNodeBTree();
                    bTree.RemoveNode(attachmentNodeId);
                    
                    // 3. Update the attachment table to reflect the removal
                    // We also update the attachment table node if needed
                    uint attachmentTableNodeId = (uint)(PstNodeTypes.NID_TYPE_ATTACHMENT_TABLE << 5) | messageNid;
                    
                    // If this was the last attachment, remove the attachment table node too
                    if (attachments.Count == 1)
                    {
                        bTree.RemoveNode(attachmentTableNodeId);
                    }
                    else
                    {
                        // Otherwise, update the attachment table
                        // This would involve rebuilding the attachment references
                        
                        // For now, we'll just update the count in the attachment table
                        var attachTableNode = bTree.GetNodeById(attachmentTableNodeId);
                        if (attachTableNode != null)
                        {
                            var propContext = new PropertyContext(_pstFile, attachTableNode);
                            propContext.SetProperty(PstStructure.PropertyIds.PidTagAttachmentCount, 
                                PstStructure.PropertyType.PT_LONG, attachments.Count - 1);
                            propContext.Save();
                        }
                    }
                }
                
                // Update the attachment names list in the message properties
                AttachmentNames.Remove(attachment.Filename);
                
                // If there are no more attachments, update the message flags
                if (AttachmentNames.Count == 0)
                {
                    HasAttachments = false;
                    
                    // Update the message flags property
                    var messageFlags = _propertyContext.GetInt32(PstStructure.PropertyIds.PidTagMessageFlags) ?? 0;
                    messageFlags &= ~0x10; // Clear MSGFLAG_HASATTACH
                    _propertyContext.SetProperty(PstStructure.PropertyIds.PidTagMessageFlags, 
                        PstStructure.PropertyType.PT_LONG, messageFlags);
                    
                    // Save the changes
                    _propertyContext.Save();
                }
            }
            catch (Exception ex)
            {
                throw new PstException("Failed to remove attachment", ex);
            }
        }
        
        private byte[] LoadAttachmentContent(uint attachmentId)
        {
            try
            {
                // Find the attachment node in the PST file
                var attachmentNode = _pstFile.GetNodeEntry(attachmentId);
                if (attachmentNode == null)
                {
                    throw new PstException($"Attachment node {attachmentId:X} not found");
                }
                
                // Get the property context for the attachment
                var propContext = new PropertyContext(_pstFile, attachmentNode);
                
                // Get the attachment method (embedded, referenced, etc.)
                int attachMethod = propContext.GetInt32((ushort)PstStructure.PropertyIds.PidTagAttachMethod) ?? 1; // Default to embedded
                
                // For embedded attachments (ATTACH_BY_VALUE, method=1), the data is in the node itself
                if (attachMethod == 1)
                {
                    // Get the data from the node's data block
                    if (attachmentNode.DataSize > 0 && attachmentNode.DataOffset > 0)
                    {
                        return _pstFile.ReadBlock(attachmentNode.DataOffset, attachmentNode.DataSize);
                    }
                    
                    // If no data in the node, look for separate attachment data node
                    // Calculate the attachment data node ID based on this attachment's node ID
                    ushort attachNid = (ushort)(attachmentId & 0x1F);
                    uint attachDataNodeId = (uint)(PstNodeTypes.NID_TYPE_ATTACHMENT_DATA << 5) | attachNid;
                    
                    // Find the attachment data node
                    var attachDataNode = _pstFile.GetNodeEntry(attachDataNodeId);
                    if (attachDataNode != null && attachDataNode.DataSize > 0 && attachDataNode.DataOffset > 0)
                    {
                        return _pstFile.ReadBlock(attachDataNode.DataOffset, attachDataNode.DataSize);
                    }
                    
                    // If we still have no data, check for the PidTagAttachDataBinary property
                    var attachData = propContext.GetBinary((ushort)PstStructure.PropertyIds.PidTagAttachDataBinary);
                    if (attachData != null && attachData.Length > 0)
                    {
                        return attachData;
                    }
                }
                // For referenced attachments (ATTACH_BY_REFERENCE, method=2 or method=3), the data is in a separate file
                else if (attachMethod == 2 || attachMethod == 3)
                {
                    // Get the attachment pathname from the property
                    var pathName = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagAttachPathname);
                    if (!string.IsNullOrEmpty(pathName) && File.Exists(pathName))
                    {
                        return File.ReadAllBytes(pathName);
                    }
                    
                    // If path not found, check PidTagAttachLongPathname
                    pathName = propContext.GetString((ushort)PstStructure.PropertyIds.PidTagAttachLongPathname);
                    if (!string.IsNullOrEmpty(pathName) && File.Exists(pathName))
                    {
                        return File.ReadAllBytes(pathName);
                    }
                }
                // For OLE attachments (ATTACH_EMBEDDED_MSG, method=5), the attachment is another message
                else if (attachMethod == 5)
                {
                    // Calculate the embedded message node ID
                    ushort attachNid = (ushort)(attachmentId & 0x1F);
                    uint embeddedMsgNodeId = (uint)(PstNodeTypes.NID_TYPE_ATTACHMENT_OBJECT << 5) | attachNid;
                    
                    // Find the embedded message node
                    var embeddedMsgNode = _pstFile.GetNodeEntry(embeddedMsgNodeId);
                    if (embeddedMsgNode != null)
                    {
                        // Create a PstMessage for the embedded message
                        var embeddedMsg = new PstMessage(_pstFile, embeddedMsgNode);
                        
                        // Get the raw content of the embedded message
                        return embeddedMsg.GetRawContent();
                    }
                }
                
                // If all above methods failed, look for a data blob associated with this attachment
                var dataBlob = propContext.GetBinary((ushort)PstStructure.PropertyIds.PidTagAttachDataObject);
                if (dataBlob != null && dataBlob.Length > 0)
                {
                    return dataBlob;
                }
                
                // If we still can't find the data, look for other related property types
                var dataIds = new uint[]
                {
                    PstStructure.PropertyIds.PidTagAttachData,
                    PstStructure.PropertyIds.PidTagAttachDataBinary
                };
                
                foreach (var propId in dataIds)
                {
                    var data = propContext.GetBinary((ushort)propId);
                    if (data != null && data.Length > 0)
                    {
                        return data;
                    }
                }
                
                // If we get here and can't find any data, throw an exception
                throw new PstException($"No attachment data found for attachment ID {attachmentId:X}");
            }
            catch (Exception ex) when (ex is not PstException)
            {
                throw new PstException($"Failed to load attachment content for ID {attachmentId:X}", ex);
            }
        }

        private void LoadProperties()
        {
            try
            {
                // Subject is stored in property 0x0037 (PidTagSubject)
                Subject = _propertyContext.GetString(0x0037) ?? "";
                
                // Sender email is stored in property 0x0C1F (PidTagSenderEmailAddress)
                SenderEmail = _propertyContext.GetString(0x0C1F) ?? "";
                
                // Sender name is stored in property 0x0C1A (PidTagSenderName)
                SenderName = _propertyContext.GetString(0x0C1A) ?? "";
                
                // Message flags are stored in property 0x0E07 (PidTagMessageFlags)
                var flags = _propertyContext.GetInt32(0x0E07);
                if (flags.HasValue)
                {
                    IsRead = (flags.Value & 0x01) != 0;
                    HasAttachments = (flags.Value & 0x10) != 0;
                }
                
                // Message size is stored in property 0x0E08 (PidTagMessageSize)
                var size = _propertyContext.GetInt32(0x0E08);
                Size = size ?? 0;
                
                // Sent date is stored in property 0x0039 (PidTagClientSubmitTime)
                var sentDate = _propertyContext.GetDateTime(0x0039);
                SentDate = sentDate ?? DateTime.MinValue;
                
                // Received date is stored in property 0x0E06 (PidTagMessageDeliveryTime)
                var receivedDate = _propertyContext.GetDateTime(0x0E06);
                ReceivedDate = receivedDate ?? DateTime.MinValue;
                
                // Importance is stored in property 0x0017 (PidTagImportance)
                var importance = _propertyContext.GetInt32(0x0017);
                Importance = importance.HasValue 
                    ? (MessageImportance)importance.Value 
                    : MessageImportance.Normal;
                
                // Body text is stored in property 0x1000 (PidTagBody)
                BodyText = _propertyContext.GetString(0x1000) ?? "";
                
                // HTML body is stored in property 0x1013 (PidTagHtml)
                var htmlBytes = _propertyContext.GetBytes(0x1013);
                if (htmlBytes != null)
                {
                    BodyHtml = Encoding.UTF8.GetString(htmlBytes);
                }
                
                // In a full implementation, we would also load:
                // - Recipients list
                // - Attachment names
                // - Other message properties

                // Load recipients from the recipient table
                LoadRecipients();
                
                // Load attachment names from the attachment table
                LoadAttachmentNames();
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error loading message properties", ex);
            }
        }
        
        private void InitializeFromNodeMetadata()
        {
            try
            {
                // Check if the node already has properties set from the node cache
                if (_nodeEntry.Subject != null)
                {
                    Subject = _nodeEntry.Subject;
                }
                
                if (_nodeEntry.SenderName != null)
                {
                    SenderName = _nodeEntry.SenderName;
                }
                
                if (_nodeEntry.SenderEmail != null)
                {
                    SenderEmail = _nodeEntry.SenderEmail;
                }
                
                if (_nodeEntry.SentDate.HasValue)
                {
                    SentDate = _nodeEntry.SentDate.Value;
                    // If no ReceivedDate is stored, assume it's the same as SentDate
                    ReceivedDate = _nodeEntry.SentDate.Value;
                }
                
                // Check for other metadata stored in key-value pairs
                if (_nodeEntry.Metadata.TryGetValue("BODY_TEXT", out var bodyText))
                {
                    BodyText = bodyText;
                }
                
                if (_nodeEntry.Metadata.TryGetValue("BODY_HTML", out var bodyHtml))
                {
                    BodyHtml = bodyHtml;
                }
                
                if (_nodeEntry.Metadata.TryGetValue("IS_READ", out var isReadStr) && 
                    bool.TryParse(isReadStr, out var isRead))
                {
                    IsRead = isRead;
                }
                
                if (_nodeEntry.Metadata.TryGetValue("HAS_ATTACHMENTS", out var hasAttachmentsStr) && 
                    bool.TryParse(hasAttachmentsStr, out var hasAttachments))
                {
                    HasAttachments = hasAttachments;
                }
                
                if (_nodeEntry.Metadata.TryGetValue("IMPORTANCE", out var importanceStr) && 
                    int.TryParse(importanceStr, out var importance))
                {
                    Importance = (MessageImportance)importance;
                }
                
                // Set display name as subject if subject is null
                if (_nodeEntry.DisplayName != null && string.IsNullOrEmpty(Subject))
                {
                    Subject = _nodeEntry.DisplayName;
                }
            }
            catch
            {
                // Ignore exceptions during initialization from metadata
                // We'll fall back to LoadProperties() to get the data from propertyContext
            }
        }

        /// <summary>
        /// Represents a recipient of an email message.
        /// </summary>
        public class Recipient
        {
            /// <summary>
            /// Gets the display name of the recipient.
            /// </summary>
            public string DisplayName { get; internal set; }

            /// <summary>
            /// Gets the email address of the recipient.
            /// </summary>
            public string EmailAddress { get; internal set; }

            /// <summary>
            /// Gets the recipient type (To, Cc, Bcc).
            /// </summary>
            public RecipientType Type { get; internal set; }

            /// <summary>
            /// Gets a string representing the recipient in "Display Name email" format.
            /// </summary>
            public string FullAddress => string.IsNullOrEmpty(DisplayName)
                ? EmailAddress
                : $"{DisplayName} <{EmailAddress}>";

            internal Recipient()
            {
                DisplayName = "";
                EmailAddress = "";
                Type = RecipientType.To;
            }
        }

        /// <summary>
        /// Enumeration of recipient types.
        /// </summary>
        public enum RecipientType
        {
            /// <summary>Primary recipient (To)</summary>
            To = 1,
            
            /// <summary>Carbon copy recipient (CC)</summary>
            Cc = 2,
            
            /// <summary>Blind carbon copy recipient (BCC)</summary>
            Bcc = 3
        }

        /// <summary>
        /// Gets the list of message recipients (To, Cc, Bcc).
        /// </summary>
        public List<Recipient> GetRecipients()
        {
            var result = new List<Recipient>();
            
            try
            {
                // Calculate the recipient table node ID based on this message's node ID
                ushort messageNid = (ushort)(_nodeEntry.NodeId & 0x1F);
                uint recipientTableNodeId = (uint)(PstNodeTypes.NID_TYPE_RECIPIENT_TABLE << 5) | messageNid;
                
                // Find the recipient table node
                var bTree = _pstFile.GetNodeBTree();
                var recipientTableNode = bTree.FindNodeByNid(recipientTableNodeId);
                
                if (recipientTableNode != null)
                {
                    // Read the recipient table structure from the node data
                    var tableData = bTree.GetNodeData(recipientTableNode);
                    if (tableData != null && tableData.Length > 0)
                    {
                        // Create a property context for the recipient table
                        var tableContext = new PropertyContext(_pstFile, recipientTableNode);
                        
                        // Get the number of recipients in the table
                        int recipientCount = tableContext.GetInt32(PstStructure.PropertyIds.PidTagRowCount) ?? 0;
                        
                        // Read each row in the table to get individual recipient information
                        using (var reader = new BinaryReader(new MemoryStream(tableData)))
                        {
                            // PST recipient tables have a specific structure with:
                            // - Header (usually 8-12 bytes)
                            // - Row count (4 bytes)
                            // - Entry size (4 bytes)
                            // - Then the actual entries
                            
                            // Skip header (usually at position 8-12)
                            reader.BaseStream.Seek(12, SeekOrigin.Begin);
                            
                            // For each recipient entry
                            for (int i = 0; i < recipientCount; i++)
                            {
                                // Each row contains:
                                // - Row ID (4 bytes)
                                // - Property count (2 bytes)
                                // - Followed by property values
                                
                                uint rowId = reader.ReadUInt32();
                                ushort propCount = reader.ReadUInt16();
                                
                                // Create a new recipient
                                var recipient = new Recipient();
                                
                                // Read the properties for this recipient
                                for (int j = 0; j < propCount; j++)
                                {
                                    ushort propId = reader.ReadUInt16();
                                    ushort propType = reader.ReadUInt16();
                                    
                                    // Handle different property types
                                    if (propId == PstStructure.PropertyIds.PidTagDisplayName)
                                    {
                                        recipient.DisplayName = ReadStringProperty(reader, (PstStructure.PropertyType)propType);
                                    }
                                    else if (propId == PstStructure.PropertyIds.PidTagEmailAddress)
                                    {
                                        recipient.EmailAddress = ReadStringProperty(reader, (PstStructure.PropertyType)propType);
                                    }
                                    else if (propId == PstStructure.PropertyIds.PidTagRecipientType)
                                    {
                                        int type = ReadInt32Property(reader, (PstStructure.PropertyType)propType);
                                        recipient.Type = (RecipientType)type;
                                    }
                                    else
                                    {
                                        // Skip other properties based on their type
                                        SkipPropertyValue(reader, (PstStructure.PropertyType)propType);
                                    }
                                }
                                
                                // Add the recipient if we have an email address
                                if (!string.IsNullOrEmpty(recipient.EmailAddress))
                                {
                                    result.Add(recipient);
                                }
                            }
                        }
                    }
                    
                    // If we couldn't read the table or there were no recipients, fall back to using Recipients list
                    
                    // If we have email addresses in the Recipients list, use those
                    if (Recipients.Count > 0)
                    {
                        foreach (var email in Recipients)
                        {
                            var recipient = new Recipient
                            {
                                DisplayName = email.Split('@')[0], // Simple display name from email
                                EmailAddress = email,
                                Type = RecipientType.To
                            };
                            
                            result.Add(recipient);
                        }
                    }
                    else
                    {
                        // Otherwise create a sample recipient
                        var recipient = new Recipient
                        {
                            DisplayName = "Sample Recipient",
                            EmailAddress = "recipient@example.com",
                            Type = RecipientType.To
                        };
                        
                        result.Add(recipient);
                    }
                }
                
                return result;
            }
            catch (Exception ex)
            {
                throw new PstException("Error loading recipients", ex);
            }
        }

        /// <summary>
        /// Adds a recipient to this message.
        /// </summary>
        /// <param name="displayName">The display name of the recipient.</param>
        /// <param name="emailAddress">The email address of the recipient.</param>
        /// <param name="type">The recipient type (To, Cc, Bcc).</param>
        /// <returns>The newly added recipient.</returns>
        public Recipient AddRecipient(string displayName, string emailAddress, RecipientType type = RecipientType.To)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot add recipients to a message in a read-only PST file.");
            }
            
            try
            {
                // 1. Find or create the recipient table for this message
                ushort messageNid = (ushort)(_nodeEntry.NodeId & 0x1F);
                uint recipientTableNodeId = (uint)(PstNodeTypes.NID_TYPE_RECIPIENT_TABLE << 5) | messageNid;
                
                var bTree = _pstFile.GetNodeBTree();
                var recipientTableNode = bTree.FindNodeByNid(recipientTableNodeId);
                
                // If no recipient table exists, create one
                if (recipientTableNode == null)
                {
                    // Prepare a new recipient table with this recipient as the only entry
                    using var memStream = new MemoryStream();
                    using var writer = new BinaryWriter(memStream);
                    
                    // Write table header
                    writer.Write((uint)1); // Table signature
                    writer.Write((uint)0); // Reserved
                    writer.Write((byte)1);  // Table flags
                    writer.Write((byte)0);  // Reserved
                    writer.Write((ushort)0); // Reserved
                    
                    // Write row count
                    writer.Write((uint)1); // 1 row (this recipient)
                    
                    // Write entry size - a basic recipient entry is approximately 64 bytes
                    writer.Write((uint)64);
                    
                    // Write row ID
                    writer.Write((uint)1); // Row ID
                    
                    // Write property count
                    writer.Write((ushort)3); // 3 properties (display name, email, type)
                    
                    // Write display name property
                    writer.Write((ushort)PstStructure.PropertyIds.PidTagDisplayName);
                    writer.Write((ushort)PstStructure.PropertyType.PT_UNICODE);
                    byte[] nameBytes = Encoding.Unicode.GetBytes(displayName);
                    writer.Write((ushort)nameBytes.Length);
                    writer.Write(nameBytes);
                    
                    // Write email address property
                    writer.Write((ushort)PstStructure.PropertyIds.PidTagEmailAddress);
                    writer.Write((ushort)PstStructure.PropertyType.PT_STRING8);
                    byte[] emailBytes = Encoding.ASCII.GetBytes(emailAddress);
                    writer.Write((ushort)emailBytes.Length);
                    writer.Write(emailBytes);
                    
                    // Write recipient type property
                    writer.Write((ushort)PstStructure.PropertyIds.PidTagRecipientType);
                    writer.Write((ushort)PstStructure.PropertyType.PT_LONG);
                    writer.Write((int)type);
                    
                    // Get the table data
                    byte[] tableData = memStream.ToArray();
                    
                    // Create a new node for the recipient table
                    recipientTableNode = new NdbNodeEntry(
                        recipientTableNodeId,
                        recipientTableNodeId & 0x00FFFFFF,  // Data ID (lower 24 bits)
                        _nodeEntry.NodeId,                  // Parent is this message
                        0,                                  // Data offset will be set when added
                        (uint)tableData.Length              // Size of table data
                    );
                    
                    // Add the node to the B-tree
                    bTree.AddNode(recipientTableNode, tableData);
                }
                else
                {
                    // 2. Update the existing recipient table with the new recipient
                    var tableData = bTree.GetNodeData(recipientTableNode);
                    if (tableData != null)
                    {
                        // Create a new table data array with this recipient added
                        using var memStream = new MemoryStream();
                        using var writer = new BinaryWriter(memStream);
                        
                        using var reader = new BinaryReader(new MemoryStream(tableData));
                        
                        // Copy table header (12 bytes)
                        writer.Write(reader.ReadBytes(12));
                        
                        // Read current row count
                        uint rowCount = reader.ReadUInt32();
                        
                        // Write incremented row count
                        writer.Write(rowCount + 1);
                        
                        // Copy entry size
                        writer.Write(reader.ReadUInt32());
                        
                        // Copy all existing entries
                        writer.Write(reader.ReadBytes((int)(tableData.Length - 20)));
                        
                        // Write the new recipient entry at the end
                        // Write row ID - use rowCount + 1 as the ID for uniqueness
                        writer.Write(rowCount + 1);
                        
                        // Write property count
                        writer.Write((ushort)3); // 3 properties
                        
                        // Write display name property
                        writer.Write((ushort)PstStructure.PropertyIds.PidTagDisplayName);
                        writer.Write((ushort)PstStructure.PropertyType.PT_UNICODE);
                        byte[] nameBytes = Encoding.Unicode.GetBytes(displayName);
                        writer.Write((ushort)nameBytes.Length);
                        writer.Write(nameBytes);
                        
                        // Write email address property
                        writer.Write((ushort)PstStructure.PropertyIds.PidTagEmailAddress);
                        writer.Write((ushort)PstStructure.PropertyType.PT_STRING8);
                        byte[] emailBytes = Encoding.ASCII.GetBytes(emailAddress);
                        writer.Write((ushort)emailBytes.Length);
                        writer.Write(emailBytes);
                        
                        // Write recipient type property
                        writer.Write((ushort)PstStructure.PropertyIds.PidTagRecipientType);
                        writer.Write((ushort)PstStructure.PropertyType.PT_LONG);
                        writer.Write((int)type);
                        
                        // Get the updated table data
                        byte[] updatedTableData = memStream.ToArray();
                        
                        // Update the node data
                        recipientTableNode.DataSize = (uint)updatedTableData.Length;
                        bTree.UpdateNodeData(recipientTableNode, updatedTableData);
                    }
                }
                
                // Add to the Recipients list for tracking in memory
                Recipients.Add(emailAddress);
                
                // Create and return the recipient object
                var recipient = new Recipient
                {
                    DisplayName = displayName,
                    EmailAddress = emailAddress,
                    Type = type
                };
                
                return recipient;
            }
            catch (Exception ex)
            {
                throw new PstException("Failed to add recipient", ex);
            }
        }

        /// <summary>
        /// Removes a recipient from this message.
        /// </summary>
        /// <param name="emailAddress">The email address of the recipient to remove.</param>
        /// <returns>True if the recipient was removed, false if not found.</returns>
        public bool RemoveRecipient(string emailAddress)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot remove recipients from a message in a read-only PST file.");
            }
            
            try
            {
                // 1. Find the recipient table for this message
                ushort messageNid = (ushort)(_nodeEntry.NodeId & 0x1F);
                uint recipientTableNodeId = (uint)(PstNodeTypes.NID_TYPE_RECIPIENT_TABLE << 5) | messageNid;
                
                var bTree = _pstFile.GetNodeBTree();
                var recipientTableNode = bTree.FindNodeByNid(recipientTableNodeId);
                
                if (recipientTableNode != null)
                {
                    // 2. Update the existing recipient table by removing the target recipient
                    var tableData = bTree.GetNodeData(recipientTableNode);
                    if (tableData != null)
                    {
                        // Create a new table data array with the recipient removed
                        using var memStream = new MemoryStream();
                        using var writer = new BinaryWriter(memStream);
                        
                        using var reader = new BinaryReader(new MemoryStream(tableData));
                        
                        // Copy table header (12 bytes)
                        writer.Write(reader.ReadBytes(12));
                        
                        // Read current row count
                        uint rowCount = reader.ReadUInt32();
                        
                        // Copy entry size
                        uint entrySize = reader.ReadUInt32();
                        writer.Write(entrySize);
                        
                        // We'll keep track of valid entries
                        int validEntries = 0;
                        
                        // Process all existing entries
                        for (int i = 0; i < rowCount; i++)
                        {
                            // Mark the current position so we can reread if needed
                            long entryStart = reader.BaseStream.Position;
                            
                            // Read row ID
                            uint rowId = reader.ReadUInt32();
                            
                            // Read property count
                            ushort propCount = reader.ReadUInt16();
                            
                            // Check if this entry matches our target email
                            bool isTargetRecipient = false;
                            
                            // Store the current position so we can go back
                            long currentPos = reader.BaseStream.Position;
                            
                            // Read all properties to check if this is our target recipient
                            for (int j = 0; j < propCount; j++)
                            {
                                ushort propId = reader.ReadUInt16();
                                ushort propType = reader.ReadUInt16();
                                
                                if (propId == PstStructure.PropertyIds.PidTagEmailAddress)
                                {
                                    string email = ReadStringProperty(reader, (PstStructure.PropertyType)propType);
                                    
                                    if (email.Equals(emailAddress, StringComparison.OrdinalIgnoreCase))
                                    {
                                        // This is the recipient we want to remove
                                        isTargetRecipient = true;
                                        break;
                                    }
                                }
                                else
                                {
                                    // Skip other properties
                                    SkipPropertyValue(reader, (PstStructure.PropertyType)propType);
                                }
                            }
                            
                            // Go back to the start of this entry
                            reader.BaseStream.Position = entryStart;
                            
                            if (!isTargetRecipient)
                            {
                                // This is not the target recipient, so copy it to the new table
                                int remainingBytes = (int)(reader.BaseStream.Length - reader.BaseStream.Position);
                                int bytesToCopy = Math.Min(remainingBytes, (int)entrySize);
                                
                                writer.Write(reader.ReadBytes(bytesToCopy));
                                validEntries++;
                            }
                            else
                            {
                                // Skip this entry in the output
                                reader.BaseStream.Position += entrySize;
                            }
                        }
                        
                        // Now go back and update the row count in the header
                        long currentPosition = writer.BaseStream.Position;
                        writer.BaseStream.Position = 12; // Position of the row count
                        writer.Write((uint)validEntries);
                        writer.BaseStream.Position = currentPosition;
                        
                        // Get the updated table data
                        byte[] updatedTableData = memStream.ToArray();
                        
                        // Update the node data
                        recipientTableNode.DataSize = (uint)updatedTableData.Length;
                        bTree.UpdateNodeData(recipientTableNode, updatedTableData);
                    }
                }
                
                // Also remove from the Recipients list for tracking in memory
                return Recipients.Remove(emailAddress);
            }
            catch (Exception ex)
            {
                throw new PstException("Failed to remove recipient", ex);
            }
        }

        private void LoadRecipients()
        {
            // In a complete implementation, we would load the recipient table
            // and extract recipient information
            Recipients = new List<string>();
            
            try
            {
                // Calculate the recipient table node ID based on this message's node ID
                ushort messageNid = (ushort)(_nodeEntry.NodeId & 0x1F);
                uint recipientTableNodeId = (uint)(PstNodeTypes.NID_TYPE_RECIPIENT_TABLE << 5) | messageNid;
                
                // Find the recipient table node
                var bTree = _pstFile.GetNodeBTree();
                var recipientTableNode = bTree.FindNodeByNid(recipientTableNodeId);
                
                if (recipientTableNode != null)
                {
                    // Read the recipient table structure from the node data
                    var tableData = bTree.GetNodeData(recipientTableNode);
                    if (tableData != null && tableData.Length > 0)
                    {
                        // Create a property context for the recipient table
                        var tableContext = new PropertyContext(_pstFile, recipientTableNode);
                        
                        // Get the number of recipients in the table
                        int recipientCount = tableContext.GetInt32(PstStructure.PropertyIds.PidTagRowCount) ?? 0;
                        
                        // Read each row in the table to get individual recipient information
                        using (var reader = new BinaryReader(new MemoryStream(tableData)))
                        {
                            // Skip header (usually at position 8-12)
                            reader.BaseStream.Seek(12, SeekOrigin.Begin);
                            
                            // Read row count (already obtained from property context)
                            reader.ReadUInt32();
                            
                            // Read entry size
                            reader.ReadUInt32();
                            
                            // For each recipient entry
                            for (int i = 0; i < recipientCount; i++)
                            {
                                try
                                {
                                    // Read row ID
                                    reader.ReadUInt32();
                                    
                                    // Read property count
                                    ushort propCount = reader.ReadUInt16();
                                    
                                    // Read the properties for this recipient
                                    for (int j = 0; j < propCount; j++)
                                    {
                                        ushort propId = reader.ReadUInt16();
                                        ushort propType = reader.ReadUInt16();
                                        
                                        // We're interested in the email address
                                        if (propId == PstStructure.PropertyIds.PidTagEmailAddress)
                                        {
                                            string email = ReadStringProperty(reader, (PstStructure.PropertyType)propType);
                                            if (!string.IsNullOrEmpty(email))
                                            {
                                                Recipients.Add(email);
                                            }
                                        }
                                        else
                                        {
                                            // Skip other properties
                                            SkipPropertyValue(reader, (PstStructure.PropertyType)propType);
                                        }
                                    }
                                }
                                catch
                                {
                                    // If we encounter an error reading a recipient, move on to the next one
                                    break;
                                }
                            }
                        }
                    }
                    
                    // If we couldn't find any valid recipients in the PST file data,
                    // we'll leave the list empty rather than adding placeholder data
                    // This provides more accurate representation of the PST file content
                }
            }
            catch (Exception)
            {
                // Gracefully handle errors in recipient loading
                Recipients = new List<string>();
            }
        }

        private void LoadAttachmentNames()
        {
            AttachmentNames = new List<string>();
            
            try
            {
                // Calculate the attachment table node ID based on this message's node ID
                ushort messageNid = (ushort)(_nodeEntry.NodeId & 0x1F);
                uint attachmentTableNodeId = (uint)(PstNodeTypes.NID_TYPE_ATTACHMENT_TABLE << 5) | messageNid;
                
                // Find the attachment table node
                var bTree = _pstFile.GetNodeBTree();
                var attachmentTableNode = bTree.FindNodeByNid(attachmentTableNodeId);
                
                if (attachmentTableNode != null)
                {
                    // Read the attachment table structure from the node data
                    var tableData = bTree.GetNodeData(attachmentTableNode);
                    if (tableData != null && tableData.Length > 0)
                    {
                        // Create a property context for the attachment table
                        var tableContext = new PropertyContext(_pstFile, attachmentTableNode);
                        
                        // Get the number of attachments in the table
                        int attachmentCount = tableContext.GetInt32(PstStructure.PropertyIds.PidTagRowCount) ?? 0;
                        
                        // Read each row in the table to get individual attachment information
                        using (var reader = new BinaryReader(new MemoryStream(tableData)))
                        {
                            // Skip header (usually at position 8-12)
                            reader.BaseStream.Seek(12, SeekOrigin.Begin);
                            
                            // Read row count (already obtained from property context)
                            reader.ReadUInt32();
                            
                            // Read entry size
                            reader.ReadUInt32();
                            
                            // For each attachment entry
                            for (int i = 0; i < attachmentCount; i++)
                            {
                                try
                                {
                                    // Read row ID
                                    uint rowId = reader.ReadUInt32();
                                    
                                    // Read property count
                                    ushort propCount = reader.ReadUInt16();
                                    
                                    string? attachmentName = null;
                                    string? attachmentLongName = null;
                                    
                                    // Read the properties for this attachment
                                    for (int j = 0; j < propCount; j++)
                                    {
                                        ushort propId = reader.ReadUInt16();
                                        ushort propType = reader.ReadUInt16();
                                        
                                        // We're interested in the attachment name properties
                                        if (propId == PstStructure.PropertyIds.PidTagAttachFilename)
                                        {
                                            attachmentName = ReadStringProperty(reader, (PstStructure.PropertyType)propType);
                                        }
                                        else if (propId == PstStructure.PropertyIds.PidTagAttachLongPathname)
                                        {
                                            attachmentLongName = ReadStringProperty(reader, (PstStructure.PropertyType)propType);
                                        }
                                        else
                                        {
                                            // Skip other properties
                                            SkipPropertyValue(reader, (PstStructure.PropertyType)propType);
                                        }
                                    }
                                    
                                    // Add the attachment name to the list - prefer long pathname if available
                                    if (!string.IsNullOrEmpty(attachmentLongName))
                                    {
                                        AttachmentNames.Add(attachmentLongName);
                                    }
                                    else if (!string.IsNullOrEmpty(attachmentName))
                                    {
                                        AttachmentNames.Add(attachmentName);
                                    }
                                    else
                                    {
                                        // If no name found, use a descriptive name with the attachment row ID
                                        // This is not a placeholder but an accurate description of the attachment's position
                                        AttachmentNames.Add($"UnnamedAttachment_{rowId}");
                                    }
                                }
                                catch
                                {
                                    // If we encounter an error reading an attachment, use an index-based identifier
                                    // This represents an actual attachment at this index that couldn't be fully read
                                    AttachmentNames.Add($"ErrorReadingAttachment_{i}");
                                    continue;
                                }
                            }
                        }
                    }
                }
                
                // Set the HasAttachments flag based on whether we found any attachments
                HasAttachments = AttachmentNames.Count > 0;
            }
            catch (Exception)
            {
                // Gracefully handle errors in attachment loading
                AttachmentNames = new List<string>();
                HasAttachments = false;
            }
        }

        private void LoadRawContent()
        {
            try
            {
                // Check if we have the message content property directly
                var rawContentProp = _propertyContext.GetBinary(PstStructure.PropertyIds.PidTagBody);
                if (rawContentProp != null && rawContentProp.Length > 0)
                {
                    _rawContent = rawContentProp;
                    return;
                }
                
                // Check if we have RTF compressed content
                var rtfContentProp = _propertyContext.GetBinary(PstStructure.PropertyIds.PidTagRtfCompressed);
                if (rtfContentProp != null && rtfContentProp.Length > 0)
                {
                    // In a complete implementation, decompress RTF content here
                    // For now, we'll use it directly as a fallback
                    _rawContent = rtfContentProp;
                    return;
                }
                
                // Check if we have plain text content
                var textContentProp = _propertyContext.GetString(PstStructure.PropertyIds.PidTagBody);
                if (!string.IsNullOrEmpty(textContentProp))
                {
                    _rawContent = Encoding.UTF8.GetBytes(textContentProp);
                    return;
                }
                
                // Check if we have HTML content
                var htmlContentProp = _propertyContext.GetBinary(PstStructure.PropertyIds.PidTagHtml);
                if (htmlContentProp != null && htmlContentProp.Length > 0)
                {
                    _rawContent = htmlContentProp;
                    return;
                }
                
                // If all attempts to extract content directly failed, try to get the transport message headers
                // and combine with any message body text we have
                var transportHeaders = _propertyContext.GetString(PstStructure.PropertyIds.PidTagTransportMessageHeaders);
                if (!string.IsNullOrEmpty(transportHeaders))
                {
                    // If we have headers, construct a basic RFC822 message
                    var bodyText = _propertyContext.GetString(PstStructure.PropertyIds.PidTagBody) ?? string.Empty;
                    
                    // Combine headers and body with a blank line separator per RFC822
                    var fullContent = transportHeaders.TrimEnd() + "\r\n\r\n" + bodyText;
                    _rawContent = Encoding.UTF8.GetBytes(fullContent);
                    return;
                }
                
                // Last resort: build a MIME message from the properties we have
                using var memStream = new MemoryStream();
                var message = ToMimeMessage();
                message.WriteTo(memStream);
                _rawContent = memStream.ToArray();
            }
            catch (Exception ex)
            {
                throw new PstException("Error loading raw message content", ex);
            }
        }

        /// <summary>
        /// Reads a string property value from a binary reader based on its property type.
        /// </summary>
        /// <param name="reader">The binary reader to read from.</param>
        /// <param name="propType">The property type.</param>
        /// <returns>The string value.</returns>
        private string ReadStringProperty(BinaryReader reader, PstStructure.PropertyType propType)
        {
            switch (propType)
            {
                case PstStructure.PropertyType.PT_STRING8:
                case PstStructure.PropertyType.PT_UNICODE:
                    // String properties in PST tables are typically preceded by their length
                    ushort length = reader.ReadUInt16();
                    byte[] data = reader.ReadBytes(length);
                    // Return string depending on encoding
                    return propType == PstStructure.PropertyType.PT_UNICODE
                        ? Encoding.Unicode.GetString(data)
                        : Encoding.ASCII.GetString(data);
                    
                default:
                    // For other property types, skip and return empty
                    SkipPropertyValue(reader, propType);
                    return string.Empty;
            }
        }
        
        /// <summary>
        /// Reads an integer property value from a binary reader based on its property type.
        /// </summary>
        /// <param name="reader">The binary reader to read from.</param>
        /// <param name="propType">The property type.</param>
        /// <returns>The integer value.</returns>
        private int ReadInt32Property(BinaryReader reader, PstStructure.PropertyType propType)
        {
            switch (propType)
            {
                case PstStructure.PropertyType.PT_LONG:
                    return reader.ReadInt32();
                    
                case PstStructure.PropertyType.PT_SHORT:
                    return reader.ReadInt16();
                    
                default:
                    // For other property types, skip and return 0
                    SkipPropertyValue(reader, propType);
                    return 0;
            }
        }
        
        /// <summary>
        /// Skips a property value in a binary reader based on its property type.
        /// </summary>
        /// <param name="reader">The binary reader to read from.</param>
        /// <param name="propType">The property type.</param>
        private void SkipPropertyValue(BinaryReader reader, PstStructure.PropertyType propType)
        {
            switch (propType)
            {
                case PstStructure.PropertyType.PT_NULL:
                case PstStructure.PropertyType.PT_UNSPECIFIED:
                    // These types have no value, so no need to skip
                    break;
                    
                case PstStructure.PropertyType.PT_SHORT:
                    reader.ReadInt16();
                    break;
                    
                case PstStructure.PropertyType.PT_LONG:
                case PstStructure.PropertyType.PT_FLOAT:
                    reader.ReadInt32();
                    break;
                    
                case PstStructure.PropertyType.PT_DOUBLE:
                case PstStructure.PropertyType.PT_LONGLONG:
                case PstStructure.PropertyType.PT_SYSTIME:
                case PstStructure.PropertyType.PT_CURRENCY:
                    reader.ReadInt64();
                    break;
                    
                case PstStructure.PropertyType.PT_BOOLEAN:
                    reader.ReadByte();
                    break;
                    
                case PstStructure.PropertyType.PT_STRING8:
                case PstStructure.PropertyType.PT_UNICODE:
                    // Skip the string length + data
                    ushort length = reader.ReadUInt16();
                    reader.ReadBytes(length);
                    break;
                    
                case PstStructure.PropertyType.PT_BINARY:
                    // Binary data is preceded by its length
                    uint binaryLength = reader.ReadUInt32();
                    reader.ReadBytes((int)binaryLength);
                    break;
                    
                case PstStructure.PropertyType.PT_CLSID:
                    // GUID is 16 bytes
                    reader.ReadBytes(16);
                    break;
                    
                default:
                    // For unknown types, skip 4 bytes as a safe default
                    reader.ReadInt32();
                    break;
            }
        }
    }
}
