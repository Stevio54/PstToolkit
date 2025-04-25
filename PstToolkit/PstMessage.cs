using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using MimeKit;
using PstToolkit.Exceptions;
using PstToolkit.Utils;

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

            // Load the message properties
            LoadProperties();
        }

        /// <summary>
        /// Creates a new email message that can be added to a PST file.
        /// </summary>
        /// <param name="subject">The subject of the email.</param>
        /// <param name="body">The body text of the email.</param>
        /// <param name="senderEmail">The sender's email address.</param>
        /// <param name="senderName">The sender's display name.</param>
        /// <returns>A new PstMessage instance.</returns>
        public static PstMessage Create(string subject, string body, string senderEmail, string senderName)
        {
            // This would be implemented in a full implementation
            // For now, we'll return null as the creation requires a context in the PST file
            throw new NotImplementedException("Creating new messages outside of a PST context is not yet implemented");
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
        /// Gets the attachments of the message.
        /// </summary>
        /// <returns>A list of attachments as byte arrays with their filenames.</returns>
        public List<(string Filename, byte[] Content)> GetAttachments()
        {
            var result = new List<(string Filename, byte[] Content)>();
            
            // In a full implementation, this would extract attachments from the PST structure
            // For now, we'll return an empty list
            
            return result;
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

        private void LoadRecipients()
        {
            // In a complete implementation, we would load the recipient table
            // and extract recipient information
            Recipients = new List<string>();
            
            // This is a simplified version that would be expanded in a real implementation
            try
            {
                // Recipients would be loaded from the recipient table
                // associated with this message node
                
                // For demonstration, we're leaving this as an empty list
            }
            catch (Exception)
            {
                // Gracefully handle errors in recipient loading
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
