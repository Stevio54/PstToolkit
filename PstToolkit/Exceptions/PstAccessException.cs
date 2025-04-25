using System;

namespace PstToolkit.Exceptions
{
    /// <summary>
    /// Exception thrown when there is an error accessing the PST file.
    /// </summary>
    public class PstAccessException : PstException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PstAccessException"/> class.
        /// </summary>
        public PstAccessException() : base("Could not access the PST file.")
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstAccessException"/> class with a specified error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public PstAccessException(string message) : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstAccessException"/> class with a specified error message
        /// and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        public PstAccessException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
