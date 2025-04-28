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
        /// Gets the starting page offset for the B-tree heap structure.
        /// </summary>
        public uint BTreeOnHeapStartPage { get; private set; }

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
                    
                    // B-tree start page is typically at offset 0x100 for ANSI PST
                    header.BTreeOnHeapStartPage = 0x100;
                    
                    // Root folder ID is typically 0x21 for ANSI PST
                    header.RootFolderId = 0x21;
                }
                else
                {
                    reader.BaseStream.Position = 0xC4;
                    header.NodeBTreeRoot = reader.ReadUInt32();
                    header.BlockBTreeRoot = reader.ReadUInt32();
                    
                    // B-tree start page is typically at offset 0x200 for Unicode PST
                    header.BTreeOnHeapStartPage = 0x200;
                    
                    // Root folder ID is typically 0x42 for Unicode PST
                    header.RootFolderId = 0x42;
                }
                
                // Perform extended validation
                header.IsValid = header.ValidateHeader();
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
                header.BTreeOnHeapStartPage = 0x100;
                header.RootFolderId = 0x21;
            }
            else
            {
                header.NodeBTreeRoot = 0x1D;
                header.BlockBTreeRoot = 0x1E;
                header.BTreeOnHeapStartPage = 0x200;
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
            // Move to the beginning of the file
            writer.BaseStream.Position = 0;
            
            // Create the header bytes array with appropriate size based on format
            int headerSize = IsAnsi ? 0x200 : 0x400;
            byte[] headerBytes = new byte[headerSize];
            
            // Write signature (first 4 bytes)
            BitConverter.GetBytes(Signature).CopyTo(headerBytes, 0);
            
            // Write format type info (offset 8)
            headerBytes[8] = (byte)(IsAnsi ? 0 : 1);
            
            // Write version info (offset 10)
            BitConverter.GetBytes(MajorVersion).CopyTo(headerBytes, 10);
            BitConverter.GetBytes(MinorVersion).CopyTo(headerBytes, 12);
            
            // Fill in other standard header fields
            
            // Write file size placeholder (offset 0x78 for ANSI, 0xB8 for Unicode)
            int fileSizeOffset = IsAnsi ? 0x78 : 0xB8;
            BitConverter.GetBytes((ulong)headerSize).CopyTo(headerBytes, fileSizeOffset);
            
            // Set B-tree root info
            int rootOffset = IsAnsi ? 0xA4 : 0xC4;
            BitConverter.GetBytes(NodeBTreeRoot).CopyTo(headerBytes, rootOffset);
            BitConverter.GetBytes(BlockBTreeRoot).CopyTo(headerBytes, rootOffset + 4);
            
            // Write a placeholder value for the number of nodes in the B-tree
            BitConverter.GetBytes(1).CopyTo(headerBytes, rootOffset + 8);
            
            // Set B-tree heap start page
            BitConverter.GetBytes(BTreeOnHeapStartPage).CopyTo(headerBytes, rootOffset + 16);
            
            // Write root folder ID at appropriate location
            int rootFolderOffset = IsAnsi ? 0xC4 : 0xE4;
            BitConverter.GetBytes(RootFolderId).CopyTo(headerBytes, rootFolderOffset);
            
            // Add a format verification value ("!BDN" or "BBDN")
            if (IsAnsi)
            {
                headerBytes[0x1FC] = 0x21; // '!'
                headerBytes[0x1FD] = 0x42; // 'B'
                headerBytes[0x1FE] = 0x44; // 'D'
                headerBytes[0x1FF] = 0x4E; // 'N'
            }
            else
            {
                headerBytes[0x3FC] = 0x42; // 'B'
                headerBytes[0x3FD] = 0x42; // 'B'
                headerBytes[0x3FE] = 0x44; // 'D'
                headerBytes[0x3FF] = 0x4E; // 'N'
            }
            
            // Write the complete header
            writer.Write(headerBytes);
        }
        
        /// <summary>
        /// Performs extended validation of the PST header.
        /// </summary>
        /// <returns>True if the header is valid, false otherwise.</returns>
        private bool ValidateHeader()
        {
            // Validate the signature
            bool isValidSignature = false;
            
            if (FormatType == PstFormatType.Ansi && Signature == ANSI_SIGNATURE)
            {
                isValidSignature = true;
            }
            else if (FormatType == PstFormatType.Unicode && Signature == UNICODE_SIGNATURE)
            {
                isValidSignature = true;
            }
            
            if (!isValidSignature)
            {
                return false;
            }
            
            // Validate version
            if (FormatType == PstFormatType.Ansi)
            {
                // ANSI PST should have version 11-14
                if (MajorVersion < 11 || MajorVersion > 14)
                {
                    return false;
                }
            }
            else
            {
                // Unicode PST should have version 15+
                if (MajorVersion < 15)
                {
                    return false;
                }
            }
            
            // Validate B-tree roots
            if (NodeBTreeRoot == 0 || BlockBTreeRoot == 0)
            {
                return false;
            }
            
            // Validate root folder ID
            if (RootFolderId == 0)
            {
                return false;
            }
            
            // Validate BTreeOnHeapStartPage
            if (BTreeOnHeapStartPage == 0)
            {
                return false;
            }
            
            return true;
        }
    }
}
