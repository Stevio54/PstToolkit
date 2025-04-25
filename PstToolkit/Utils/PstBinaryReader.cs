using System;
using System.IO;
using System.Text;

namespace PstToolkit.Utils
{
    /// <summary>
    /// Binary reader with PST-specific functionality.
    /// </summary>
    internal class PstBinaryReader : BinaryReader
    {
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
        /// Reads a variable-length string from the current stream.
        /// </summary>
        /// <param name="length">The length of the string in bytes.</param>
        /// <param name="encoding">The character encoding to use.</param>
        /// <returns>The string that was read.</returns>
        public string ReadString(int length, Encoding encoding)
        {
            byte[] buffer = ReadBytes(length);
            return encoding.GetString(buffer);
        }

        /// <summary>
        /// Reads a null-terminated string from the current stream.
        /// </summary>
        /// <param name="encoding">The character encoding to use.</param>
        /// <returns>The string that was read.</returns>
        public string ReadNullTerminatedString(Encoding encoding)
        {
            var bytes = new List<byte>();
            byte b;
            
            while ((b = ReadByte()) != 0)
            {
                bytes.Add(b);
            }
            
            return encoding.GetString(bytes.ToArray());
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
        /// Reads a specified number of bytes from the stream, starting at a specified position.
        /// </summary>
        /// <param name="position">The position in the stream to begin reading from.</param>
        /// <param name="count">The number of bytes to read.</param>
        /// <returns>A byte array containing the read bytes.</returns>
        public byte[] ReadBytesAt(long position, int count)
        {
            long originalPosition = BaseStream.Position;
            BaseStream.Position = position;
            
            byte[] buffer = ReadBytes(count);
            
            BaseStream.Position = originalPosition;
            return buffer;
        }
    }
}
