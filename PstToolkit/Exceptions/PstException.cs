using System;

namespace PstToolkit.Exceptions
{
    /// <summary>
    /// Base exception for all PST-related errors.
    /// </summary>
    public class PstException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PstException"/> class.
        /// </summary>
        public PstException() : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstException"/> class with a specified error message.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        public PstException(string message) : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PstException"/> class with a specified error message
        /// and a reference to the inner exception that is the cause of this exception.
        /// </summary>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        public PstException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
