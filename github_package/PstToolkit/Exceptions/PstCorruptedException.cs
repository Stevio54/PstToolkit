using System;

namespace PstToolkit.Exceptions
{
    /// <summary>
    /// Exception thrown when the PST file is found to be corrupted.
    /// </summary>
    public class PstCorruptedException : PstException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PstCorruptedException"/> class.
        /// </summary>
        public PstCorruptedException() : base("The PST file is corrupted.")
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstCorruptedException"/> class with a specified error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public PstCorruptedException(string message) : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstCorruptedException"/> class with a specified error message
        /// and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        public PstCorruptedException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
