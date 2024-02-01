using System;
using System.Diagnostics;

namespace ditjson
{
    /// <inheritdoc/>
    [DebuggerDisplay($"{{{nameof(GetDebuggerDisplay)}(),nq}}")]
    public class NtdsException : Exception
    {
        public NtdsException()
        {
        }

        public NtdsException(string? message) : base(message)
        {
        }

        public NtdsException(string? message, Exception? innerException) : base(message, innerException)
        {
        }

        private string GetDebuggerDisplay() => ToString();
    }
}
