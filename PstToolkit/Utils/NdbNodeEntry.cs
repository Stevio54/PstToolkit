using System;
using System.Collections.Generic;

namespace PstToolkit.Utils
{
    /// <summary>
    /// Represents a node entry in the Node Database (NDB) layer of a PST file.
    /// </summary>
    internal class NdbNodeEntry
    {
        /// <summary>
        /// Gets the node ID.
        /// </summary>
        public uint NodeId { get; }

        /// <summary>
        /// Gets the data ID.
        /// </summary>
        public uint DataId { get; }

        /// <summary>
        /// Gets the parent node ID.
        /// </summary>
        public uint ParentId { get; }

        /// <summary>
        /// Gets the node type.
        /// </summary>
        public ushort NodeType { get; }

        /// <summary>
        /// Gets the data offset in the file.
        /// </summary>
        public ulong DataOffset { get; }

        /// <summary>
        /// Gets the data size in bytes.
        /// </summary>
        public uint DataSize { get; }
        
        /// <summary>
        /// Gets or sets the display name (for folders and messages).
        /// </summary>
        public string? DisplayName { get; set; }
        
        /// <summary>
        /// Gets or sets the sender name (for messages).
        /// </summary>
        public string? SenderName { get; set; }
        
        /// <summary>
        /// Gets or sets the sender email (for messages).
        /// </summary>
        public string? SenderEmail { get; set; }
        
        /// <summary>
        /// Gets or sets the message subject (for messages).
        /// </summary>
        public string? Subject { get; set; }
        
        /// <summary>
        /// Gets or sets the sent date (for messages).
        /// </summary>
        public DateTime? SentDate { get; set; }
        
        /// <summary>
        /// Gets or sets additional metadata for the node.
        /// </summary>
        public Dictionary<string, string> Metadata { get; } = new Dictionary<string, string>();

        /// <summary>
        /// Initializes a new instance of the <see cref="NdbNodeEntry"/> class.
        /// </summary>
        /// <param name="nodeId">The node ID.</param>
        /// <param name="dataId">The data ID.</param>
        /// <param name="parentId">The parent node ID.</param>
        /// <param name="dataOffset">The data offset in the file.</param>
        /// <param name="dataSize">The data size in bytes.</param>
        public NdbNodeEntry(uint nodeId, uint dataId, uint parentId, ulong dataOffset, uint dataSize)
        {
            NodeId = nodeId;
            DataId = dataId;
            ParentId = parentId;
            NodeType = (ushort)((nodeId >> 5) & 0x1F);
            DataOffset = dataOffset;
            DataSize = dataSize;
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="NdbNodeEntry"/> class with additional metadata.
        /// </summary>
        /// <param name="nodeId">The node ID.</param>
        /// <param name="dataId">The data ID.</param>
        /// <param name="parentId">The parent node ID.</param>
        /// <param name="dataOffset">The data offset in the file.</param>
        /// <param name="dataSize">The data size in bytes.</param>
        /// <param name="displayName">The display name for the node.</param>
        public NdbNodeEntry(uint nodeId, uint dataId, uint parentId, ulong dataOffset, uint dataSize, string displayName)
            : this(nodeId, dataId, parentId, dataOffset, dataSize)
        {
            DisplayName = displayName;
        }

        /// <summary>
        /// Reads node data from the PST file.
        /// </summary>
        /// <param name="pstFile">The PST file.</param>
        /// <returns>The node data as a byte array.</returns>
        public byte[] ReadData(PstFile pstFile)
        {
            if (DataSize == 0)
                return Array.Empty<byte>();
                
            return pstFile.ReadBlock(DataOffset, DataSize);
        }
        
        /// <summary>
        /// Sets a metadata value for the node.
        /// </summary>
        /// <param name="key">The metadata key.</param>
        /// <param name="value">The metadata value.</param>
        public void SetMetadata(string key, string value)
        {
            Metadata[key] = value;
        }
        
        /// <summary>
        /// Gets a metadata value from the node.
        /// </summary>
        /// <param name="key">The metadata key.</param>
        /// <returns>The metadata value, or null if the key doesn't exist.</returns>
        public string? GetMetadata(string key)
        {
            return Metadata.TryGetValue(key, out var value) ? value : null;
        }

        /// <summary>
        /// Creates a new node entry from the given parameters.
        /// </summary>
        /// <param name="nodeId">The node ID.</param>
        /// <param name="dataId">The data ID.</param>
        /// <param name="parentId">The parent node ID.</param>
        /// <param name="dataOffset">The data offset in the file.</param>
        /// <param name="dataSize">The data size in bytes.</param>
        /// <returns>A new NdbNodeEntry instance.</returns>
        public static NdbNodeEntry Create(uint nodeId, uint dataId, uint parentId, ulong dataOffset, uint dataSize)
        {
            return new NdbNodeEntry(nodeId, dataId, parentId, dataOffset, dataSize);
        }
    }
}
