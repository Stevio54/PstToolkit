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
    }
}
