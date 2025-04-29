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
        private const uint PST_SIGNATURE = 0x4E444221;    // '!BDN' - Common signature for both ANSI and Unicode PST files

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
                
                // Validate that the signature matches the PST file format
                if (header.Signature != PST_SIGNATURE)
                {
                    header.IsValid = false;
                    return header;
                }
                
                // Read version information at offset 10
                reader.BaseStream.Position = 10;
                header.MajorVersion = reader.ReadUInt16();
                header.MinorVersion = reader.ReadUInt16();
                
                // Determine format type based on the major version
                // According to Microsoft spec: version 23+ = Unicode, version 14 = ANSI
                header.FormatType = (header.MajorVersion >= 23) ? PstFormatType.Unicode : PstFormatType.Ansi;
                
                // Read B-tree root information according to Microsoft PST File Format spec
                if (header.FormatType == PstFormatType.Ansi)
                {
                    // ANSI Format (PST files 97-2002)
                    reader.BaseStream.Position = 0x0C0;
                    header.NodeBTreeRoot = reader.ReadUInt32();
                    reader.BaseStream.Position = 0x0C4;
                    header.BlockBTreeRoot = reader.ReadUInt32();
                    
                    // B-tree start page is at offset 0x0E0 in ANSI format
                    reader.BaseStream.Position = 0x0E0;
                    header.BTreeOnHeapStartPage = reader.ReadUInt32();
                    
                    // Root folder ID (NID_ROOT_FOLDER) is 0x21 for ANSI PST
                    header.RootFolderId = 0x21;
                }
                else
                {
                    // Unicode Format (PST files 2003+)
                    reader.BaseStream.Position = 0x0D0;
                    header.NodeBTreeRoot = reader.ReadUInt32();
                    reader.BaseStream.Position = 0x0D4;
                    header.BlockBTreeRoot = reader.ReadUInt32();
                    
                    // B-tree start page is at offset 0x0F0 in Unicode format
                    reader.BaseStream.Position = 0x0F0;
                    header.BTreeOnHeapStartPage = reader.ReadUInt32();
                    
                    // Root folder ID (NID_ROOT_FOLDER) is 0x42 for Unicode PST
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
                Signature = PST_SIGNATURE, // Both ANSI and Unicode PST files use the same signature
                MajorVersion = formatType == PstFormatType.Ansi ? (ushort)14 : (ushort)23,
                MinorVersion = 0,
                IsValid = true
            };
            
            // Set root nodes based on Microsoft PST file format spec
            if (formatType == PstFormatType.Ansi)
            {
                // ANSI format (PST 97-2002)
                header.NodeBTreeRoot = 0x0D;            // NDB_ROOT_NID Node ID
                header.BlockBTreeRoot = 0x0E;           // NDB_ROOT_BB Block ID
                header.BTreeOnHeapStartPage = 0x4000;   // Start of B-tree on heap (allocation)
                header.RootFolderId = 0x21;             // NID_ROOT_FOLDER
            }
            else
            {
                // Unicode format (PST 2003+)
                header.NodeBTreeRoot = 0x1D;            // NDB_ROOT_NID Node ID
                header.BlockBTreeRoot = 0x1E;           // NDB_ROOT_BB Block ID
                header.BTreeOnHeapStartPage = 0x8000;   // Start of B-tree on heap (allocation)
                header.RootFolderId = 0x42;             // NID_ROOT_FOLDER
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
            
            // Write version info (offset 10) - This is what determines ANSI vs Unicode format
            BitConverter.GetBytes(MajorVersion).CopyTo(headerBytes, 10);
            BitConverter.GetBytes(MinorVersion).CopyTo(headerBytes, 12);
            
            // Fill in other standard header fields
            
            // Write file size according to the PST File Format Specification
            int fileSizeOffset = IsAnsi ? 0x0A8 : 0x0B8; // dwFileSize offset
            // Use the current stream length as the file size
            ulong fileSize = (ulong)writer.BaseStream.Length;
            // Ensure minimum size is the header size
            fileSize = Math.Max(fileSize, (ulong)headerSize);
            BitConverter.GetBytes(fileSize).CopyTo(headerBytes, fileSizeOffset);
            
            // Set B-tree root info according to the Microsoft spec
            if (IsAnsi)
            {
                // ANSI Format B-tree node values
                BitConverter.GetBytes(NodeBTreeRoot).CopyTo(headerBytes, 0x0C0); // Root.nid NDB_ROOT_NID
                BitConverter.GetBytes(BlockBTreeRoot).CopyTo(headerBytes, 0x0C4); // Root.nid NDB_ROOT_BB
                
                // Write the initial number of nodes in the B-tree
                // For a new PST file, we typically start with root nodes plus a few system nodes
                uint initialNodeCount = 5;
                BitConverter.GetBytes(initialNodeCount).CopyTo(headerBytes, 0x0C8); // cEntries
                
                // Set B-tree heap start page
                BitConverter.GetBytes(BTreeOnHeapStartPage).CopyTo(headerBytes, 0x0E0); // bidUnused (start of BBTREE)
            }
            else
            {
                // Unicode Format B-tree node values
                BitConverter.GetBytes(NodeBTreeRoot).CopyTo(headerBytes, 0x0D0); // Root.nid NDB_ROOT_NID
                BitConverter.GetBytes(BlockBTreeRoot).CopyTo(headerBytes, 0x0D4); // Root.nid NDB_ROOT_BB
                
                // Write the initial number of nodes in the B-tree
                // For a new PST file, we typically start with root nodes plus a few system nodes
                uint initialNodeCount = 7;
                BitConverter.GetBytes(initialNodeCount).CopyTo(headerBytes, 0x0D8); // cEntries
                
                // Set B-tree heap start page
                BitConverter.GetBytes(BTreeOnHeapStartPage).CopyTo(headerBytes, 0x0F0); // bidUnused (start of BBTREE)
            }
            
            // Root folder ID - Store in a special field for reference during hierarchy table creation
            // Not part of the standard header but needed for proper operation
            if (IsAnsi)
            {
                BitConverter.GetBytes(RootFolderId).CopyTo(headerBytes, 0x184);
            }
            else
            {
                BitConverter.GetBytes(RootFolderId).CopyTo(headerBytes, 0x1A4);
            }
            
            // Add a format verification value "!BDN" at the end of the header for both formats
            if (IsAnsi)
            {
                headerBytes[0x1FC] = 0x21; // '!'
                headerBytes[0x1FD] = 0x42; // 'B'
                headerBytes[0x1FE] = 0x44; // 'D'
                headerBytes[0x1FF] = 0x4E; // 'N'
            }
            else
            {
                headerBytes[0x3FC] = 0x21; // '!'
                headerBytes[0x3FD] = 0x42; // 'B'
                headerBytes[0x3FE] = 0x44; // 'D'
                headerBytes[0x3FF] = 0x4E; // 'N'
            }
            
            // Write the complete header with all needed PST structure data
            writer.Write(headerBytes);
            
            // Flush to ensure all data is written
            writer.Flush();
        }
        
        /// <summary>
        /// Performs extended validation of the PST header.
        /// </summary>
        /// <returns>True if the header is valid, false otherwise.</returns>
        private bool ValidateHeader()
        {
            // Validate the signature (both ANSI and Unicode PST files use the same '!BDN' signature)
            if (Signature != PST_SIGNATURE)
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
