using System;
using System.Collections.Generic;
using System.Text;
using PstToolkit.Exceptions;
using static PstToolkit.Utils.PstStructure;

namespace PstToolkit.Utils
{
    /// <summary>
    /// Provides access to properties stored in a PST node.
    /// </summary>
    internal class PropertyContext
    {
        private readonly PstFile _pstFile;
        private readonly NdbNodeEntry _nodeEntry;
        private Dictionary<uint, object>? _properties;

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyContext"/> class.
        /// </summary>
        /// <param name="pstFile">The PST file.</param>
        /// <param name="nodeEntry">The node entry.</param>
        public PropertyContext(PstFile pstFile, NdbNodeEntry nodeEntry)
        {
            _pstFile = pstFile;
            _nodeEntry = nodeEntry;
        }

        /// <summary>
        /// Gets a string property.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <returns>The string value if found, or null if not found.</returns>
        public string? GetString(ushort propertyId)
        {
            EnsurePropertiesLoaded();
            
            // Try as ANSI string (PT_STRING8)
            uint key = MakePropertyId(propertyId, PropertyType.PT_STRING8);
            if (_properties!.TryGetValue(key, out var value) && value is string str)
            {
                return str;
            }
            
            // Try as Unicode string (PT_UNICODE)
            key = MakePropertyId(propertyId, PropertyType.PT_UNICODE);
            if (_properties.TryGetValue(key, out value) && value is string unicodeStr)
            {
                return unicodeStr;
            }
            
            return null;
        }

        /// <summary>
        /// Gets an integer property.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <returns>The integer value if found, or null if not found.</returns>
        public int? GetInt32(ushort propertyId)
        {
            EnsurePropertiesLoaded();
            
            uint key = MakePropertyId(propertyId, PropertyType.PT_LONG);
            if (_properties!.TryGetValue(key, out var value) && value is int intValue)
            {
                return intValue;
            }
            
            return null;
        }

        /// <summary>
        /// Gets a long integer property.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <returns>The long integer value if found, or null if not found.</returns>
        public long? GetInt64(ushort propertyId)
        {
            EnsurePropertiesLoaded();
            
            uint key = MakePropertyId(propertyId, PropertyType.PT_LONGLONG);
            if (_properties!.TryGetValue(key, out var value) && value is long longValue)
            {
                return longValue;
            }
            
            return null;
        }

        /// <summary>
        /// Gets a DateTime property.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <returns>The DateTime value if found, or null if not found.</returns>
        public DateTime? GetDateTime(ushort propertyId)
        {
            EnsurePropertiesLoaded();
            
            uint key = MakePropertyId(propertyId, PropertyType.PT_SYSTIME);
            if (_properties!.TryGetValue(key, out var value) && value is DateTime dtValue)
            {
                return dtValue;
            }
            
            return null;
        }

        /// <summary>
        /// Gets a boolean property.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <returns>The boolean value if found, or null if not found.</returns>
        public bool? GetBoolean(ushort propertyId)
        {
            EnsurePropertiesLoaded();
            
            uint key = MakePropertyId(propertyId, PropertyType.PT_BOOLEAN);
            if (_properties!.TryGetValue(key, out var value) && value is bool boolValue)
            {
                return boolValue;
            }
            
            return null;
        }

        /// <summary>
        /// Gets a binary property.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <returns>The binary value as a byte array if found, or null if not found.</returns>
        public byte[]? GetBytes(ushort propertyId)
        {
            EnsurePropertiesLoaded();
            
            uint key = MakePropertyId(propertyId, PropertyType.PT_BINARY);
            if (_properties!.TryGetValue(key, out var value) && value is byte[] byteValue)
            {
                return byteValue;
            }
            
            return null;
        }

        private void EnsurePropertiesLoaded()
        {
            if (_properties != null)
                return;
                
            _properties = new Dictionary<uint, object>();
            
            try
            {
                LoadProperties();
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error loading properties from node", ex);
            }
        }

        private void LoadProperties()
        {
            // In a full implementation, this would:
            // 1. Read the property context from the node
            // 2. Parse the property table
            // 3. Store the properties in the _properties dictionary
            
            // For this demonstration, we'll just initialize an empty dictionary
            _properties = new Dictionary<uint, object>();
            
            // A complete implementation would require parsing the complex
            // property storage format used in PST files
        }
    }
}
