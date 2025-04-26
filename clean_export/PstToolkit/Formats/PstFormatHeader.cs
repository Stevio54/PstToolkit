using System;
using PstToolkit.Utils;

namespace PstToolkit.Formats
{
    /// <summary>
    /// PST file format types.
    /// </summary>
    public enum PstFormatType
    {
        /// <summary>ANSI PST file format (97-2002)</summary>
        Ansi = 0,
        
        /// <summary>Unicode PST file format (2003+)</summary>
        Unicode = 1
    }

    /// <summary>
    /// Represents the header of a PST file that contains format information.
    /// </summary>
    internal class PstFormatHeader
    {
        private const uint ANSI_SIGNATURE = 0x4E444221;    // '!BDN'
        private const uint UNICODE_SIGNATURE = 0x4E444242; // 'BBDN'

        /// <summary>
        /// Gets the PST file signature.
        /// </summary>
        public uint Signature { get; private set; }

        /// <summary>
        /// Gets the PST file format type (ANSI or Unicode).
        /// </summary>
        public PstFormatType FormatType { get; private set; }

        /// <summary>
        /// Gets whether the PST file is in ANSI format.
        /// </summary>
        public bool IsAnsi => FormatType == PstFormatType.Ansi;

        /// <summary>
        /// Gets the PST file format major version.
        /// </summary>
        public ushort MajorVersion { get; private set; }

        /// <summary>
        /// Gets the PST file format minor version.
        /// </summary>
        public ushort MinorVersion { get; private set; }

        /// <summary>
        /// Gets the root node ID of the node B-tree.
        /// </summary>
        public uint NodeBTreeRoot { get; private set; }

        /// <summary>
        /// Gets the root node ID of the block B-tree.
        /// </summary>
        public uint BlockBTreeRoot { get; private set; }

        /// <summary>
        /// Gets the ID of the root folder.
        /// </summary>
        public uint RootFolderId { get; private set; }

        /// <summary>
        /// Gets whether the PST header is valid.
        /// </summary>
        public bool IsValid { get; private set; }

        /// <summary>
        /// Reads the PST file header from the given reader.
        /// </summary>
        /// <param name="reader">The binary reader to read from.</param>
        /// <returns>A PstFormatHeader object containing the header information.</returns>
        public static PstFormatHeader Read(PstBinaryReader reader)
        {
            var header = new PstFormatHeader();
            
            try
            {
                // Read the signature (first 4 bytes)
                reader.BaseStream.Position = 0;
                header.Signature = reader.ReadUInt32();
                
                // Determine the format type based on the signature
                switch (header.Signature)
                {
                    case ANSI_SIGNATURE:
                        header.FormatType = PstFormatType.Ansi;
                        break;
                    case UNICODE_SIGNATURE:
                        header.FormatType = PstFormatType.Unicode;
                        break;
                    default:
                        header.IsValid = false;
                        return header;
                }
                
                // Read version information
                reader.BaseStream.Position = 10;
                header.MajorVersion = reader.ReadUInt16();
                header.MinorVersion = reader.ReadUInt16();
                
                // Read B-tree root information
                if (header.FormatType == PstFormatType.Ansi)
                {
                    reader.BaseStream.Position = 0xA4;
                    header.NodeBTreeRoot = reader.ReadUInt32();
                    header.BlockBTreeRoot = reader.ReadUInt32();
                    
                    // Root folder ID is typically 0x21 for ANSI PST
                    header.RootFolderId = 0x21;
                }
                else
                {
                    reader.BaseStream.Position = 0xC4;
                    header.NodeBTreeRoot = reader.ReadUInt32();
                    header.BlockBTreeRoot = reader.ReadUInt32();
                    
                    // Root folder ID is typically 0x42 for Unicode PST
                    header.RootFolderId = 0x42;
                }
                
                // For a real implementation, we would validate more extensively
                header.IsValid = true;
            }
            catch
            {
                header.IsValid = false;
            }
            
            return header;
        }

        /// <summary>
        /// Creates a new PST file header with the specified format type.
        /// </summary>
        /// <param name="formatType">The PST format type to create.</param>
        /// <returns>A PstFormatHeader object containing the header information.</returns>
        public static PstFormatHeader CreateNew(PstFormatType formatType)
        {
            var header = new PstFormatHeader
            {
                FormatType = formatType,
                Signature = formatType == PstFormatType.Ansi ? ANSI_SIGNATURE : UNICODE_SIGNATURE,
                MajorVersion = formatType == PstFormatType.Ansi ? (ushort)14 : (ushort)23,
                MinorVersion = 0,
                IsValid = true
            };
            
            // Set root nodes
            if (formatType == PstFormatType.Ansi)
            {
                header.NodeBTreeRoot = 0x0D;
                header.BlockBTreeRoot = 0x0E;
                header.RootFolderId = 0x21;
            }
            else
            {
                header.NodeBTreeRoot = 0x1D;
                header.BlockBTreeRoot = 0x1E;
                header.RootFolderId = 0x42;
            }
            
            return header;
        }

        /// <summary>
        /// Writes the PST file header to the given writer.
        /// </summary>
        /// <param name="writer">The binary writer to write to.</param>
        public void Write(PstBinaryWriter writer)
        {
            // In a real implementation, this would write the complete PST header structure
            // including all the necessary metadata fields
            
            // For now, this is a simplified placeholder implementation
            writer.BaseStream.Position = 0;
            
            // Write signature
            writer.Write(Signature);
            
            // Write placeholder for the rest of the header
            byte[] headerBytes = new byte[IsAnsi ? 0x200 : 0x400];
            
            // Set version info
            headerBytes[10] = (byte)(MajorVersion & 0xFF);
            headerBytes[11] = (byte)(MajorVersion >> 8);
            headerBytes[12] = (byte)(MinorVersion & 0xFF);
            headerBytes[13] = (byte)(MinorVersion >> 8);
            
            // Set B-tree root info
            int rootOffset = IsAnsi ? 0xA4 : 0xC4;
            BitConverter.GetBytes(NodeBTreeRoot).CopyTo(headerBytes, rootOffset);
            BitConverter.GetBytes(BlockBTreeRoot).CopyTo(headerBytes, rootOffset + 4);
            
            // Write the header
            writer.Write(headerBytes);
        }
    }
}
