using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Buffers;

namespace PstToolkit.Utils
{
    /// <summary>
    /// Binary reader with PST-specific functionality, optimized for large files.
    /// </summary>
    internal class PstBinaryReader : BinaryReader
    {
        private const int BufferSize = 8192; // 8KB buffer - good balance for performance
        
        /// <summary>
        /// Initializes a new instance of the <see cref="PstBinaryReader"/> class.
        /// </summary>
        /// <param name="stream">The stream to read from.</param>
        /// <param name="leaveOpen">Whether to leave the stream open when the reader is disposed.</param>
        public PstBinaryReader(Stream stream, bool leaveOpen = false)
            : base(stream, Encoding.ASCII, leaveOpen)
        {
        }

        /// <summary>
        /// Reads a variable-length string from the current stream with optimized memory usage.
        /// </summary>
        /// <param name="length">The length of the string in bytes.</param>
        /// <param name="encoding">The character encoding to use.</param>
        /// <returns>The string that was read.</returns>
        public string ReadString(int length, Encoding encoding)
        {
            if (length <= 0)
                return string.Empty;
                
            // For small strings, use regular ReadBytes
            if (length < 1024) 
            {
                byte[] buffer = ReadBytes(length);
                return encoding.GetString(buffer);
            }
            
            // For larger strings, use a shared buffer to reduce allocations
            byte[] sharedBuffer = ArrayPool<byte>.Shared.Rent(length);
            try
            {
                int bytesRead = Read(sharedBuffer, 0, length);
                return encoding.GetString(sharedBuffer, 0, bytesRead);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(sharedBuffer);
            }
        }

        /// <summary>
        /// Reads a null-terminated string from the current stream with optimized memory usage.
        /// </summary>
        /// <param name="encoding">The character encoding to use.</param>
        /// <returns>The string that was read.</returns>
        public string ReadNullTerminatedString(Encoding encoding)
        {
            // Use a pooled buffer to reduce allocations for null-terminated strings
            byte[] buffer = ArrayPool<byte>.Shared.Rent(256); // Initial size estimate
            try
            {
                int position = 0;
                byte b;
                
                while ((b = ReadByte()) != 0)
                {
                    // If buffer is full, expand it by renting a larger one
                    if (position >= buffer.Length)
                    {
                        byte[] newBuffer = ArrayPool<byte>.Shared.Rent(buffer.Length * 2);
                        Buffer.BlockCopy(buffer, 0, newBuffer, 0, position);
                        ArrayPool<byte>.Shared.Return(buffer);
                        buffer = newBuffer;
                    }
                    
                    buffer[position++] = b;
                }
                
                return encoding.GetString(buffer, 0, position);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        /// <summary>
        /// Reads a GUID from the current stream.
        /// </summary>
        /// <returns>The GUID that was read.</returns>
        public Guid ReadGuid()
        {
            byte[] buffer = ReadBytes(16);
            return new Guid(buffer);
        }

        /// <summary>
        /// Reads a 64-bit file time from the current stream.
        /// </summary>
        /// <returns>A DateTime representing the file time that was read.</returns>
        public DateTime ReadFileTime()
        {
            long fileTime = ReadInt64();
            
            if (fileTime == 0)
                return DateTime.MinValue;
                
            try
            {
                return DateTime.FromFileTime(fileTime);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// Reads a length-prefixed string from the current stream.
        /// </summary>
        /// <param name="encoding">The character encoding to use.</param>
        /// <returns>The string that was read.</returns>
        public string ReadLengthPrefixedString(Encoding encoding)
        {
            ushort length = ReadUInt16();
            return ReadString(length, encoding);
        }

        /// <summary>
        /// Reads a specified number of bytes from the stream, starting at a specified position,
        /// with optimizations for large data blocks.
        /// </summary>
        /// <param name="position">The position in the stream to begin reading from.</param>
        /// <param name="count">The number of bytes to read.</param>
        /// <returns>A byte array containing the read bytes.</returns>
        public byte[] ReadBytesAt(long position, int count)
        {
            if (count <= 0)
                return Array.Empty<byte>();
                
            long originalPosition = BaseStream.Position;
            try
            {
                BaseStream.Position = position;
                
                // For small blocks, use standard ReadBytes
                if (count < BufferSize) 
                {
                    return ReadBytes(count);
                }
                
                // For large blocks, read in chunks to optimize memory usage and I/O
                byte[] result = new byte[count];
                int bytesRead = 0;
                int remaining = count;
                
                while (remaining > 0)
                {
                    int chunkSize = Math.Min(BufferSize, remaining);
                    int read = BaseStream.Read(result, bytesRead, chunkSize);
                    
                    if (read == 0)
                        break; // End of stream
                        
                    bytesRead += read;
                    remaining -= read;
                }
                
                return result;
            }
            finally
            {
                // Always restore the original position
                BaseStream.Position = originalPosition;
            }
        }
        
        /// <summary>
        /// Reads a large block of data from the stream efficiently.
        /// </summary>
        /// <param name="count">The number of bytes to read.</param>
        /// <returns>A byte array containing the read bytes.</returns>
        public byte[] ReadLargeBlock(int count)
        {
            if (count <= 0)
                return Array.Empty<byte>();
                
            // For small blocks, use the standard method
            if (count < BufferSize)
            {
                return ReadBytes(count);
            }
            
            // For large blocks, read in chunks to optimize memory usage
            byte[] result = new byte[count];
            int bytesRead = 0;
            int remaining = count;
            
            while (remaining > 0)
            {
                int chunkSize = Math.Min(BufferSize, remaining);
                int read = BaseStream.Read(result, bytesRead, chunkSize);
                
                if (read == 0)
                    break; // End of stream
                    
                bytesRead += read;
                remaining -= read;
            }
            
            return result;
        }
        
        /// <summary>
        /// Reads a property value from the stream based on its data type.
        /// </summary>
        /// <param name="propertyType">The type of property to read.</param>
        /// <param name="valueSize">Size in bytes for variable-length properties.</param>
        /// <returns>The property value with the appropriate type.</returns>
        public object ReadPropertyValue(PstStructure.PropertyType propertyType, int valueSize = 0)
        {
            switch (propertyType)
            {
                case PstStructure.PropertyType.PT_UNSPECIFIED:
                case PstStructure.PropertyType.PT_NULL:
                    // Return empty string instead of null to avoid CS8603 warning
                    return string.Empty;
                    
                case PstStructure.PropertyType.PT_SHORT:
                    return ReadInt16();
                    
                case PstStructure.PropertyType.PT_LONG:
                    return ReadInt32();
                    
                case PstStructure.PropertyType.PT_FLOAT:
                    return ReadSingle();
                    
                case PstStructure.PropertyType.PT_DOUBLE:
                    return ReadDouble();
                    
                case PstStructure.PropertyType.PT_CURRENCY:
                    // Currency is an 8-byte integer scaled by 10,000
                    return ReadInt64() / 10000.0m;
                    
                case PstStructure.PropertyType.PT_APPTIME:
                    // Application time is a double representing days since Dec 30, 1899
                    double days = ReadDouble();
                    return new DateTime(1899, 12, 30).AddDays(days);
                    
                case PstStructure.PropertyType.PT_ERROR:
                    return ReadUInt32(); // Error code
                    
                case PstStructure.PropertyType.PT_BOOLEAN:
                    return ReadByte() != 0;
                    
                case PstStructure.PropertyType.PT_OBJECT:
                    if (valueSize > 0)
                    {
                        return ReadLargeBlock(valueSize);
                    }
                    return Array.Empty<byte>();
                    
                case PstStructure.PropertyType.PT_LONGLONG:
                    return ReadInt64();
                    
                case PstStructure.PropertyType.PT_STRING8:
                    if (valueSize > 0)
                    {
                        return ReadString(valueSize, Encoding.ASCII);
                    }
                    return string.Empty;
                    
                case PstStructure.PropertyType.PT_UNICODE:
                    if (valueSize > 0)
                    {
                        return ReadString(valueSize, Encoding.Unicode);
                    }
                    return string.Empty;
                    
                case PstStructure.PropertyType.PT_SYSTIME:
                    // FILETIME: 64-bit value representing 100-nanosecond intervals since January 1, 1601 UTC
                    return ReadFileTime();
                    
                case PstStructure.PropertyType.PT_CLSID:
                    return ReadGuid();
                    
                case PstStructure.PropertyType.PT_BINARY:
                    if (valueSize > 0)
                    {
                        return ReadLargeBlock(valueSize);
                    }
                    return Array.Empty<byte>();
                    
                default:
                    // Unknown property type, read as binary
                    if (valueSize > 0)
                    {
                        return ReadLargeBlock(valueSize);
                    }
                    return Array.Empty<byte>();
            }
        }
    }
}
