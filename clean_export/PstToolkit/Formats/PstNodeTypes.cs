namespace PstToolkit.Formats
{
    /// <summary>
    /// Contains constants for PST node types used in the Node Database (NDB) layer.
    /// </summary>
    internal static class PstNodeTypes
    {
        /// <summary>
        /// Node type for the internal NDB structure.
        /// </summary>
        public const ushort NID_TYPE_INTERNAL = 0x0000;
        
        /// <summary>
        /// Node type for a normal folder.
        /// </summary>
        public const ushort NID_TYPE_FOLDER = 0x0001;
        
        /// <summary>
        /// Node type for search folder contents.
        /// </summary>
        public const ushort NID_TYPE_SEARCH_FOLDER = 0x0002;
        
        /// <summary>
        /// Node type for the message store.
        /// </summary>
        public const ushort NID_TYPE_MESSAGE_STORE = 0x0003;
        
        /// <summary>
        /// Node type for a normal message.
        /// </summary>
        public const ushort NID_TYPE_MESSAGE = 0x0004;
        
        /// <summary>
        /// Node type for the attachment object.
        /// </summary>
        public const ushort NID_TYPE_ATTACHMENT = 0x0005;
        
        /// <summary>
        /// Node type for contents table of a folder.
        /// </summary>
        public const ushort NID_TYPE_CONTENTS_TABLE = 0x0006;
        
        /// <summary>
        /// Node type for the recipient table of a message.
        /// </summary>
        public const ushort NID_TYPE_RECIPIENT_TABLE = 0x0007;
        
        /// <summary>
        /// Node type for the search criteria of a search folder.
        /// </summary>
        public const ushort NID_TYPE_SEARCH_CRITERIA = 0x0008;
        
        /// <summary>
        /// Node type for the attachment table of a message.
        /// </summary>
        public const ushort NID_TYPE_ATTACHMENT_TABLE = 0x0009;
        
        /// <summary>
        /// Node type for hierarchy table of a folder.
        /// </summary>
        public const ushort NID_TYPE_HIERARCHY_TABLE = 0x000A;
        
        /// <summary>
        /// Node type for the message store contents.
        /// </summary>
        public const ushort NID_TYPE_CONTENTS = 0x000B;
        
        /// <summary>
        /// Node type for associated contents table of a folder.
        /// </summary>
        public const ushort NID_TYPE_ASSOCIATED_CONTENTS = 0x000C;
        
        /// <summary>
        /// Node type for contents table of a search folder.
        /// </summary>
        public const ushort NID_TYPE_SEARCH_CONTENTS_TABLE = 0x000D;
        
        /// <summary>
        /// Extracts the node type from a node ID.
        /// </summary>
        /// <param name="nodeId">The node ID.</param>
        /// <returns>The node type.</returns>
        public static ushort GetNodeType(uint nodeId)
        {
            return (ushort)((nodeId >> 5) & 0x1F);
        }
        
        /// <summary>
        /// Determines if the given node ID represents a folder.
        /// </summary>
        /// <param name="nodeId">The node ID to check.</param>
        /// <returns>True if the node ID represents a folder, false otherwise.</returns>
        public static bool IsFolder(uint nodeId)
        {
            return GetNodeType(nodeId) == NID_TYPE_FOLDER;
        }
        
        /// <summary>
        /// Determines if the given node ID represents a message.
        /// </summary>
        /// <param name="nodeId">The node ID to check.</param>
        /// <returns>True if the node ID represents a message, false otherwise.</returns>
        public static bool IsMessage(uint nodeId)
        {
            return GetNodeType(nodeId) == NID_TYPE_MESSAGE;
        }
        
        /// <summary>
        /// Determines if the given node ID represents an attachment.
        /// </summary>
        /// <param name="nodeId">The node ID to check.</param>
        /// <returns>True if the node ID represents an attachment, false otherwise.</returns>
        public static bool IsAttachment(uint nodeId)
        {
            return GetNodeType(nodeId) == NID_TYPE_ATTACHMENT;
        }
        
        /// <summary>
        /// Creates a new node ID with the specified node type and ID.
        /// </summary>
        /// <param name="type">The node type.</param>
        /// <param name="id">The node-specific ID.</param>
        /// <returns>A new node ID.</returns>
        public static uint CreateNodeId(ushort type, ushort id)
        {
            return (uint)((type & 0x1F) << 5) | (uint)(id & 0x1F);
        }
    }
}
