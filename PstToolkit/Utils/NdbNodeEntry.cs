using System;

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
