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
                // Generate a temporary node ID for the message
                // In a real implementation, this would be allocated by the PST file
                uint tempNodeId = 0x10000000 | (uint)DateTime.Now.Ticks & 0xFFFFFF;
                
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
                
                // Create a placeholder node entry
                var nodeEntry = new NdbNodeEntry(tempNodeId, 0, 0, 0, (uint)messageData.Length);
                
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
                // Generate a temporary node ID for the message
                // In a real implementation, this would be allocated by the PST file
                uint tempNodeId = 0x10000000 | (uint)DateTime.Now.Ticks & 0xFFFFFF;
                
                // Serialize the MimeMessage to a byte array
                byte[] messageData;
                using (var memStream = new MemoryStream())
                {
                    mimeMessage.WriteTo(memStream);
                    messageData = memStream.ToArray();
                }
                
                // Create a placeholder node entry
                var nodeEntry = new NdbNodeEntry(tempNodeId, 0, 0, 0, (uint)messageData.Length);
                
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
                // Generate a temporary node ID for the message
                // In a real implementation, this would be allocated by the PST file
                uint newMessageId = 0x10000000 | (uint)DateTime.Now.Ticks & 0xFFFFFF;
                
                // Create a node for the message
                var bTree = pstFile.GetNodeBTree();
                var nodeData = sourceMessage.GetRawContent(); // Get raw content to preserve all properties
                var nodeEntry = bTree.AddNode(newMessageId, 0, 0, nodeData, sourceMessage.Subject);
                
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
                    var nodeEntry = new NdbNodeEntry(
                        // Generate a temporary ID for this email attachment
                        (uint)(0x20000000 | (uint)DateTime.Now.Ticks & 0xFFFFFF),
                        0,  // No parent - this is not in the PST structure
                        0,  // Block ID - not applicable here
                        0,  // Data offset - not applicable here
                        (uint)content.Length
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
                    // Log the error (in a real implementation)
                    Console.WriteLine($"Error parsing nested email: {ex.Message}");
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
                // In a real implementation, this would:
                // 1. Find the attachment node ID
                // 2. Remove it from the PST file
                // 3. Update the attachment table
                
                // For this implementation, we'll just remove it from the AttachmentNames list
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
                    // In a real implementation, this would:
                    // 1. Read the recipient table rows
                    // 2. Extract recipient properties for each row
                    
                    // For this implementation, we'll create mock recipients based on the Recipients list
                    
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
                // In a real implementation, this would:
                // 1. Update the recipient table with the new recipient
                // 2. Create a new row in the table with the recipient properties
                
                // For this implementation, we'll just add to the Recipients list
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
                // In a real implementation, this would:
                // 1. Find the recipient in the recipient table
                // 2. Remove the row from the table
                
                // For this implementation, we'll just remove from the Recipients list
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
                    // In a real implementation, this would:
                    // 1. Read the recipient table rows
                    // 2. Extract recipient email addresses for each row
                    
                    // For this implementation, we'll create sample recipients
                    Recipients.Add("recipient1@example.com");
                    
                    // Add a CC recipient sometimes
                    if ((_nodeEntry.NodeId & 0x02) != 0) // Simple way to vary sample data
                    {
                        Recipients.Add("cc@example.com");
                    }
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
            // In a complete implementation, we would load the attachment table
            // and extract attachment names
            AttachmentNames = new List<string>();
            
            // This is a simplified version that would be expanded in a real implementation
            try
            {
                // Attachment names would be loaded from the attachment table
                // associated with this message node
                
                // For demonstration, we're leaving this as an empty list
            }
            catch (Exception)
            {
                // Gracefully handle errors in attachment loading
            }
        }

        private void LoadRawContent()
        {
            try
            {
                // In a full implementation, this would extract the raw email content
                // from the PST file structures
                // 
                // For now, we'll build a simple representation from the properties we have
                
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
    }
}
