using System;
using System.IO;
using System.Text;

namespace PstToolkit.Utils
{
    /// <summary>
    /// Binary writer with PST-specific functionality.
    /// </summary>
    internal class PstBinaryWriter : BinaryWriter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PstBinaryWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <param name="leaveOpen">Whether to leave the stream open when the writer is disposed.</param>
        public PstBinaryWriter(Stream stream, bool leaveOpen = false)
            : base(stream, Encoding.ASCII, leaveOpen)
        {
        }

        /// <summary>
        /// Writes a string to the current stream with the specified encoding.
        /// </summary>
        /// <param name="value">The string to write.</param>
        /// <param name="encoding">The character encoding to use.</param>
        public void WriteString(string value, Encoding encoding)
        {
            byte[] buffer = encoding.GetBytes(value);
            Write(buffer);
        }

        /// <summary>
        /// Writes a null-terminated string to the current stream with the specified encoding.
        /// </summary>
        /// <param name="value">The string to write.</param>
        /// <param name="encoding">The character encoding to use.</param>
        public void WriteNullTerminatedString(string value, Encoding encoding)
        {
            WriteString(value, encoding);
            Write((byte)0);
        }

        /// <summary>
        /// Writes a GUID to the current stream.
        /// </summary>
        /// <param name="value">The GUID to write.</param>
        public void WriteGuid(Guid value)
        {
            Write(value.ToByteArray());
        }

        /// <summary>
        /// Writes a DateTime as a 64-bit file time to the current stream.
        /// </summary>
        /// <param name="value">The DateTime to write.</param>
        public void WriteFileTime(DateTime value)
        {
            if (value == DateTime.MinValue)
            {
                Write((long)0);
                return;
            }
            
            try
            {
                Write(value.ToFileTime());
            }
            catch
            {
                Write((long)0);
            }
        }

        /// <summary>
        /// Writes a length-prefixed string to the current stream.
        /// </summary>
        /// <param name="value">The string to write.</param>
        /// <param name="encoding">The character encoding to use.</param>
        public void WriteLengthPrefixedString(string value, Encoding encoding)
        {
            byte[] buffer = encoding.GetBytes(value);
            Write((ushort)buffer.Length);
            Write(buffer);
        }

        /// <summary>
        /// Writes bytes to the stream at the specified position without changing the stream position.
        /// </summary>
        /// <param name="position">The position in the stream to write to.</param>
        /// <param name="buffer">The bytes to write.</param>
        public void WriteBytesAt(long position, byte[] buffer)
        {
            long originalPosition = BaseStream.Position;
            BaseStream.Position = position;
            
            Write(buffer);
            
            BaseStream.Position = originalPosition;
        }
    }
}
