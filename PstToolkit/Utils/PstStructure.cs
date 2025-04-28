using System;

namespace PstToolkit.Utils
{
    /// <summary>
    /// Constants and utilities related to the PST file structure.
    /// </summary>
    internal static class PstStructure
    {
        /// <summary>
        /// Property tags for common PST properties.
        /// </summary>
        public static class PropTags
        {
            // Common property tags
            public const ushort PR_DISPLAY_NAME = 0x3001;
            public const ushort PR_SUBJECT = 0x0037;
            public const ushort PR_NORMALIZED_SUBJECT = 0x0E1D;
            public const ushort PR_SENDER_NAME = 0x0C1A;
            public const ushort PR_SENT_REPRESENTING_NAME = 0x0042;
            public const ushort PR_SENDER_EMAIL_ADDRESS = 0x0C1F;
            public const ushort PR_SENT_REPRESENTING_EMAIL_ADDRESS = 0x0065;
            public const ushort PR_CLIENT_SUBMIT_TIME = 0x0039;
            public const ushort PR_MESSAGE_DELIVERY_TIME = 0x0E06;
            public const ushort PR_MESSAGE_SIZE = 0x0E08;
            public const ushort PR_HASATTACH = 0x0E1B;
        }
        
        /// <summary>
        /// Property ID constants for common PST properties.
        /// </summary>
        public static class PropertyIds
        {
            // Common Properties
            public const ushort PidTagDisplayName = 0x3001;
            public const ushort PidTagSubject = 0x0037;
            public const ushort PidTagMessageClass = 0x001A;
            public const ushort PidTagLastModificationTime = 0x3008;
            public const ushort PidTagCreationTime = 0x3007;
            
            // Message Properties
            public const ushort PidTagSenderName = 0x0C1A;
            public const ushort PidTagSenderEmailAddress = 0x0C1F;
            public const ushort PidTagSentRepresentingName = 0x0042;
            public const ushort PidTagSentRepresentingEmailAddress = 0x0065;
            public const ushort PidTagMessageDeliveryTime = 0x0E06;
            public const ushort PidTagClientSubmitTime = 0x0039;
            public const ushort PidTagMessageFlags = 0x0E07;
            public const ushort PidTagMessageSize = 0x0E08;
            public const ushort PidTagBody = 0x1000;
            public const ushort PidTagHtml = 0x1013;
            public const ushort PidTagRtfCompressed = 0x1009;
            public const ushort PidTagTransportMessageHeaders = 0x007D;
            public const ushort PidTagImportance = 0x0017;
            
            // Recipient Properties
            public const ushort PidTagEmailAddress = 0x3003;
            public const ushort PidTagRecipientType = 0x0C15;
            
            // Attachment Properties
            public const ushort PidTagAttachmentSize = 0x0E20;
            public const ushort PidTagAttachFilename = 0x3704;
            public const ushort PidTagAttachLongFilename = 0x3707;
            public const ushort PidTagAttachMethod = 0x3705;
            public const ushort PidTagAttachData = 0x3701;
            public const ushort PidTagAttachDataBinary = 0x3701;
            public const ushort PidTagAttachDataObject = 0x3702;
            public const ushort PidTagAttachPathname = 0x3708;
            public const ushort PidTagAttachLongPathname = 0x370D;
            public const ushort PidTagAttachMimeTag = 0x370E;
            public const ushort PidTagAttachmentCount = 0x0E08; // Number of attachments
            
            // Folder Properties
            public const ushort PidTagContainerClass = 0x3613;
            public const ushort PidTagContentCount = 0x3602;
            public const ushort PidTagContentUnreadCount = 0x3603;
            public const ushort PidTagSubfolders = 0x360A;
            
            // Table Properties
            public const ushort PidTagRowCount = 0x3002; // Row count in a table
            public const ushort PidTagRowType = 0x3004; // Row type in a table
        }

        /// <summary>
        /// Property types used in PST property contexts.
        /// </summary>
        public enum PropertyType : ushort
        {
            /// <summary>No type</summary>
            PT_UNSPECIFIED = 0x0000,
            
            /// <summary>Null value</summary>
            PT_NULL = 0x0001,
            
            /// <summary>2-byte integer</summary>
            PT_SHORT = 0x0002,
            
            /// <summary>4-byte integer</summary>
            PT_LONG = 0x0003,
            
            /// <summary>4-byte floating point</summary>
            PT_FLOAT = 0x0004,
            
            /// <summary>8-byte floating point</summary>
            PT_DOUBLE = 0x0005,
            
            /// <summary>Currency value</summary>
            PT_CURRENCY = 0x0006,
            
            /// <summary>Application time</summary>
            PT_APPTIME = 0x0007,
            
            /// <summary>4-byte error value</summary>
            PT_ERROR = 0x000A,
            
            /// <summary>1-byte boolean</summary>
            PT_BOOLEAN = 0x000B,
            
            /// <summary>Object/embedded object</summary>
            PT_OBJECT = 0x000D,
            
            /// <summary>8-byte integer</summary>
            PT_LONGLONG = 0x0014,
            
            /// <summary>String</summary>
            PT_STRING8 = 0x001E,
            
            /// <summary>Unicode string</summary>
            PT_UNICODE = 0x001F,
            
            /// <summary>File time</summary>
            PT_SYSTIME = 0x0040,
            
            /// <summary>GUID</summary>
            PT_CLSID = 0x0048,
            
            /// <summary>Binary data</summary>
            PT_BINARY = 0x0102,
        }

        /// <summary>
        /// Creates a complete property ID by combining a property ID and type.
        /// </summary>
        /// <param name="id">The property ID.</param>
        /// <param name="type">The property type.</param>
        /// <returns>The complete property ID.</returns>
        public static uint MakePropertyId(ushort id, PropertyType type)
        {
            return ((uint)type << 16) | id;
        }

        /// <summary>
        /// Extracts the property ID from a complete property ID.
        /// </summary>
        /// <param name="propertyId">The complete property ID.</param>
        /// <returns>The property ID.</returns>
        public static ushort GetPropertyId(uint propertyId)
        {
            return (ushort)(propertyId & 0xFFFF);
        }

        /// <summary>
        /// Extracts the property type from a complete property ID.
        /// </summary>
        /// <param name="propertyId">The complete property ID.</param>
        /// <returns>The property type.</returns>
        public static PropertyType GetPropertyType(uint propertyId)
        {
            return (PropertyType)((propertyId >> 16) & 0xFFFF);
        }
    }
}
