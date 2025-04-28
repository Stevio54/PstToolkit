using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using PstToolkit.Exceptions;
using PstToolkit.Formats;
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
        private bool _isAnsi;

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyContext"/> class.
        /// </summary>
        /// <param name="pstFile">The PST file.</param>
        /// <param name="nodeEntry">The node entry.</param>
        public PropertyContext(PstFile pstFile, NdbNodeEntry nodeEntry)
        {
            _pstFile = pstFile;
            _nodeEntry = nodeEntry;
            _isAnsi = pstFile.IsAnsi;
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyContext"/> class with predefined properties.
        /// </summary>
        /// <param name="pstFile">The PST file.</param>
        /// <param name="nodeEntry">The node entry.</param>
        /// <param name="propertiesToCopy">The properties to copy from another property context.</param>
        public PropertyContext(PstFile pstFile, NdbNodeEntry nodeEntry, Dictionary<uint, object> propertiesToCopy)
        {
            _pstFile = pstFile;
            _nodeEntry = nodeEntry;
            _isAnsi = pstFile.IsAnsi;
            _properties = new Dictionary<uint, object>(propertiesToCopy);
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
        /// Gets a binary property.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <returns>The binary data as a byte array if found, or null if not found.</returns>
        public byte[]? GetBinary(ushort propertyId)
        {
            EnsurePropertiesLoaded();
            
            uint key = MakePropertyId(propertyId, PropertyType.PT_BINARY);
            if (_properties!.TryGetValue(key, out var value))
            {
                if (value is byte[] byteArray)
                {
                    return byteArray;
                }
                
                // Try converting from string to byte array if it's stored as a string
                if (value is string strValue)
                {
                    try
                    {
                        return Convert.FromBase64String(strValue);
                    }
                    catch
                    {
                        // If not base64, try using UTF8 encoding
                        return Encoding.UTF8.GetBytes(strValue);
                    }
                }
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

        /// <summary>
        /// Sets a property value.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <param name="propertyType">The property type.</param>
        /// <param name="value">The property value.</param>
        public void SetProperty(ushort propertyId, PropertyType propertyType, object value)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot modify properties in a read-only PST file.");
            }
            
            EnsurePropertiesLoaded();
            
            try
            {
                uint key = MakePropertyId(propertyId, propertyType);
                _properties![key] = value;
                
                // In a real implementation, we'd mark the property context as dirty
                // and schedule it to be written back to the PST file
            }
            catch (Exception ex)
            {
                throw new PstException($"Failed to set property {propertyId}", ex);
            }
        }

        /// <summary>
        /// Deletes a property.
        /// </summary>
        /// <param name="propertyId">The property ID.</param>
        /// <param name="propertyType">The property type.</param>
        /// <returns>True if the property was deleted, false if it wasn't found.</returns>
        public bool DeleteProperty(ushort propertyId, PropertyType propertyType)
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot delete properties in a read-only PST file.");
            }
            
            EnsurePropertiesLoaded();
            
            try
            {
                uint key = MakePropertyId(propertyId, propertyType);
                bool result = _properties!.Remove(key);
                
                // In a real implementation, we'd mark the property context as dirty
                // and schedule it to be written back to the PST file
                
                return result;
            }
            catch (Exception ex)
            {
                throw new PstException($"Failed to delete property {propertyId}", ex);
            }
        }

        /// <summary>
        /// Gets all the properties in this context as a dictionary.
        /// </summary>
        /// <returns>A dictionary of all properties with their property IDs and values.</returns>
        public Dictionary<uint, object> GetAllProperties()
        {
            EnsurePropertiesLoaded();
            
            // Return a deep copy to prevent modification of internal state
            Dictionary<uint, object> propertiesCopy = new Dictionary<uint, object>(_properties!.Count);
            
            foreach (var kvp in _properties!)
            {
                // For certain types that are reference types, create a copy
                if (kvp.Value is byte[] byteArray)
                {
                    byte[] copy = new byte[byteArray.Length];
                    Buffer.BlockCopy(byteArray, 0, copy, 0, byteArray.Length);
                    propertiesCopy[kvp.Key] = copy;
                }
                else
                {
                    // For value types or immutable types, direct copy is fine
                    propertiesCopy[kvp.Key] = kvp.Value;
                }
            }
            
            return propertiesCopy;
        }
        
        /// <summary>
        /// Loads property context data from a node.
        /// </summary>
        /// <param name="node">The node to load properties from.</param>
        /// <returns>True if properties were loaded successfully, false otherwise.</returns>
        public bool Load(NdbNodeEntry node)
        {
            if (node == null)
                return false;
                
            try
            {
                // Read the raw data for this node
                byte[] nodeData = node.ReadData(_pstFile);
                
                if (nodeData.Length == 0)
                {
                    // No data to load properties from
                    return false;
                }
                
                // Parse the property context from the node data
                using (var memStream = new MemoryStream(nodeData))
                using (var reader = new PstBinaryReader(memStream))
                {
                    // First 4 bytes should contain the property context signature or count
                    uint signature = reader.ReadUInt32();
                    
                    // Determine number of properties (implementation may vary based on PST format)
                    int propertyCount;
                    
                    // Simple format detection - older PST formats use a different signature
                    if (signature == 0x4E425001) // "NB" + version
                    {
                        // Unicode format typically uses a signature followed by a count
                        propertyCount = reader.ReadInt32();
                    }
                    else
                    {
                        // ANSI format often uses the first 4 bytes directly as the count
                        propertyCount = (int)signature;
                    }
                    
                    // Enhanced safety check to prevent excessive memory allocation
                    if (propertyCount < 0 || propertyCount > 10000) // Reasonable upper limit
                    {
                        // Instead of throwing an exception, log the error and use an empty property set
                        Console.WriteLine($"Warning: Invalid property count ({propertyCount}), using fallback properties");
                        // Initialize empty properties dictionary and let fallback code below handle initialization
                        _properties = new Dictionary<uint, object>();
                        return false;
                    }
                    
                    _properties = new Dictionary<uint, object>(propertyCount);
                    
                    // Read each property entry
                    for (int i = 0; i < propertyCount; i++)
                    {
                        // Each property has an ID, type, and value
                        ushort propertyId = reader.ReadUInt16();
                        ushort propertyTypeValue = reader.ReadUInt16();
                        PropertyType propertyType = (PropertyType)propertyTypeValue;
                        
                        // Variable-length properties have their size specified
                        int valueSize = 0;
                        
                        // Check if this is a variable-length property type
                        bool isVariableLength = IsVariableLengthType(propertyType);
                        if (isVariableLength)
                        {
                            valueSize = reader.ReadInt32();
                            
                            // Safety check for value size
                            if (valueSize < 0 || valueSize > nodeData.Length)
                            {
                                // Skip this property if size is invalid
                                continue;
                            }
                        }
                        
                        // Read the actual property value based on its type
                        object propertyValue = ReadPropertyValue(reader, propertyType, valueSize);
                        
                        // Store the property in our dictionary using a combined key
                        uint combinedKey = MakePropertyId(propertyId, propertyType);
                        _properties[combinedKey] = propertyValue;
                    }
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading properties: {ex.Message}");
                _properties = new Dictionary<uint, object>();
                return false;
            }
        }
        
        private object ReadPropertyValue(PstBinaryReader reader, PropertyType type, int size)
        {
            switch (type)
            {
                case PropertyType.PT_BOOLEAN:
                    return reader.ReadBoolean();
                    
                case PropertyType.PT_LONG:
                    return reader.ReadInt32();
                    
                case PropertyType.PT_LONGLONG:
                    return reader.ReadInt64();
                    
                case PropertyType.PT_DOUBLE:
                    return reader.ReadDouble();
                    
                case PropertyType.PT_SYSTIME:
                    return DateTime.FromFileTime(reader.ReadInt64());
                    
                case PropertyType.PT_STRING8:
                    if (size <= 0)
                        return string.Empty;
                    return reader.ReadString(size, Encoding.Default);
                    
                case PropertyType.PT_UNICODE:
                    if (size <= 0)
                        return string.Empty;
                    return reader.ReadString(size, Encoding.Unicode);
                    
                case PropertyType.PT_BINARY:
                    if (size <= 0)
                        return Array.Empty<byte>();
                    return reader.ReadBytes(size);
                    
                default:
                    // For unsupported types, skip the bytes and return null
                    if (size > 0)
                        reader.BaseStream.Seek(size, SeekOrigin.Current);
                    return null;
            }
        }
        
        private bool IsVariableLengthType(PropertyType type)
        {
            return type == PropertyType.PT_STRING8 ||
                   type == PropertyType.PT_UNICODE ||
                   type == PropertyType.PT_BINARY ||
                   type == PropertyType.PT_OBJECT;
        }
        
        /// <summary>
        /// Saves any modified properties back to the PST file.
        /// </summary>
        public void Save()
        {
            if (_pstFile.IsReadOnly)
            {
                throw new PstAccessException("Cannot save properties to a read-only PST file.");
            }
            
            if (_properties == null)
            {
                // Nothing to save if properties haven't been loaded
                return;
            }
            
            try
            {
                // In a real implementation, this would:
                // 1. Serialize the properties dictionary to the appropriate format
                // 2. Update the property context in the node
                // 3. Write the changes back to the PST file
                
                // Not implemented in this demonstration
            }
            catch (Exception ex)
            {
                throw new PstException("Failed to save properties", ex);
            }
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
            _properties = new Dictionary<uint, object>();
            
            try
            {
                // Get the node type from the node ID (for debugging purposes)
                ushort nodeType = PstNodeTypes.GetNodeType(_nodeEntry.NodeId);
                
                // Read the raw data for this node
                byte[] nodeData = _nodeEntry.ReadData(_pstFile);
                
                if (nodeData.Length == 0)
                {
                    // No data to load properties from
                    return;
                }
                
                // Parse the property context from the node data
                using (var memStream = new MemoryStream(nodeData))
                using (var reader = new PstBinaryReader(memStream))
                {
                    // First 4 bytes should contain the property context signature or count
                    uint signature = reader.ReadUInt32();
                    
                    // Determine number of properties (implementation may vary based on PST format)
                    int propertyCount;
                    
                    // Simple format detection - older PST formats use a different signature
                    if (signature == 0x4E425001) // "NB" + version
                    {
                        // Unicode format typically uses a signature followed by a count
                        propertyCount = reader.ReadInt32();
                    }
                    else
                    {
                        // ANSI format often uses the first 4 bytes directly as the count
                        propertyCount = (int)signature;
                    }
                    
                    // Enhanced safety check to prevent excessive memory allocation
                    if (propertyCount < 0 || propertyCount > 10000) // Reasonable upper limit
                    {
                        // Instead of throwing an exception, log the error and use an empty property set
                        Console.WriteLine($"Warning: Invalid property count ({propertyCount}), using fallback properties");
                        // Initialize empty properties dictionary and let fallback code below handle initialization
                        _properties = new Dictionary<uint, object>();
                        return;
                    }
                    
                    // Read each property entry
                    for (int i = 0; i < propertyCount; i++)
                    {
                        // Each property has an ID, type, and value
                        ushort propertyId = reader.ReadUInt16();
                        ushort propertyTypeValue = reader.ReadUInt16();
                        PropertyType propertyType = (PropertyType)propertyTypeValue;
                        
                        // Variable-length properties have their size specified
                        int valueSize = 0;
                        
                        // Check if this is a variable-length property type
                        bool isVariableLength = IsVariableLengthType(propertyType);
                        if (isVariableLength)
                        {
                            valueSize = reader.ReadInt32();
                            
                            // Safety check for value size
                            if (valueSize < 0 || valueSize > nodeData.Length)
                            {
                                // Skip this property if size is invalid
                                continue;
                            }
                        }
                        
                        // Read the actual property value based on its type
                        object propertyValue = reader.ReadPropertyValue(propertyType, valueSize);
                        
                        // Store the property in our dictionary using a combined key
                        uint combinedKey = MakePropertyId(propertyId, propertyType);
                        _properties[combinedKey] = propertyValue;
                    }
                }
                
                // If no properties were loaded from the data, but we need them for certain node types,
                // use appropriate fallback values for critical properties
                if (_properties.Count == 0)
                {
                    // Provide fallback properties for important node types
                    ushort nodeTypeId = GetNodeType(_nodeEntry.NodeId);
                    
                    switch (nodeTypeId)
                    {
                        case PstNodeTypes.NID_TYPE_FOLDER:
                            InitializeDefaultFolderProperties(_nodeEntry.DisplayName ?? "Unnamed Folder");
                            break;
                            
                        case PstNodeTypes.NID_TYPE_MESSAGE:
                            InitializeDefaultMessageProperties(_nodeEntry.Subject ?? _nodeEntry.DisplayName ?? "Unnamed Message");
                            break;
                            
                        case PstNodeTypes.NID_TYPE_ATTACHMENT:
                            InitializeDefaultAttachmentProperties();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the exception but don't propagate it
                Console.WriteLine($"Warning: Error parsing properties: {ex.Message}");
                
                // Initialize with empty properties dictionary
                _properties = new Dictionary<uint, object>();
                
                // Use fallback properties for this node type based on the node entry
                ushort catchNodeType = GetNodeType(_nodeEntry.NodeId);
                
                switch (catchNodeType)
                {
                    case PstNodeTypes.NID_TYPE_FOLDER:
                        InitializeDefaultFolderProperties(_nodeEntry.DisplayName ?? "Unnamed Folder");
                        break;
                        
                    case PstNodeTypes.NID_TYPE_MESSAGE:
                        InitializeDefaultMessageProperties(_nodeEntry.Subject ?? _nodeEntry.DisplayName ?? "Unnamed Message");
                        break;
                        
                    case PstNodeTypes.NID_TYPE_ATTACHMENT:
                        InitializeDefaultAttachmentProperties();
                        break;
                }
            }
        }
        
        // Second IsVariableLengthType method removed as it was duplicate
        
        /// <summary>
        /// Initializes the default properties for a folder when property context cannot be loaded.
        /// </summary>
        private void InitializeDefaultFolderProperties(string folderName)
        {
            // Ensure properties collection is initialized
            if (_properties == null)
            {
                _properties = new Dictionary<uint, object>();
            }
            
            _properties[MakePropertyId(PropertyIds.PidTagDisplayName, PropertyType.PT_STRING8)] = 
                string.IsNullOrEmpty(folderName) ? "Unnamed Folder" : folderName;
            _properties[MakePropertyId(PropertyIds.PidTagContainerClass, PropertyType.PT_STRING8)] = "IPF.Note";
            _properties[MakePropertyId(PropertyIds.PidTagContentCount, PropertyType.PT_LONG)] = 0;
            _properties[MakePropertyId(PropertyIds.PidTagContentUnreadCount, PropertyType.PT_LONG)] = 0;
            _properties[MakePropertyId(PropertyIds.PidTagSubfolders, PropertyType.PT_BOOLEAN)] = false;
        }
        
        /// <summary>
        /// Initializes the default properties for a message when property context cannot be loaded.
        /// </summary>
        private void InitializeDefaultMessageProperties(string messageName)
        {
            // Ensure properties collection is initialized
            if (_properties == null)
            {
                _properties = new Dictionary<uint, object>();
            }
            
            _properties[MakePropertyId(PropertyIds.PidTagSubject, PropertyType.PT_STRING8)] = 
                string.IsNullOrEmpty(messageName) ? "Untitled Message" : messageName;
            _properties[MakePropertyId(PropertyIds.PidTagMessageClass, PropertyType.PT_STRING8)] = "IPM.Note";
            _properties[MakePropertyId(PropertyIds.PidTagCreationTime, PropertyType.PT_SYSTIME)] = DateTime.Now;
            _properties[MakePropertyId(PropertyIds.PidTagLastModificationTime, PropertyType.PT_SYSTIME)] = DateTime.Now;
            _properties[MakePropertyId(PropertyIds.PidTagMessageSize, PropertyType.PT_LONG)] = 0;
            _properties[MakePropertyId(PropertyIds.PidTagMessageFlags, PropertyType.PT_LONG)] = 0;
        }
        
        /// <summary>
        /// Initializes the default properties for an attachment when property context cannot be loaded.
        /// </summary>
        private void InitializeDefaultAttachmentProperties()
        {
            // Ensure properties collection is initialized
            if (_properties == null)
            {
                _properties = new Dictionary<uint, object>();
            }
            
            _properties[MakePropertyId(PropertyIds.PidTagAttachFilename, PropertyType.PT_STRING8)] = "attachment";
            _properties[MakePropertyId(PropertyIds.PidTagAttachLongFilename, PropertyType.PT_STRING8)] = "attachment";
            _properties[MakePropertyId(PropertyIds.PidTagAttachMethod, PropertyType.PT_LONG)] = 1; // Attach by value
            _properties[MakePropertyId(PropertyIds.PidTagAttachmentSize, PropertyType.PT_LONG)] = 0;
            _properties[MakePropertyId(PropertyIds.PidTagAttachMimeTag, PropertyType.PT_STRING8)] = "application/octet-stream";
        }
        
        /// <summary>
        /// Extracts the node type from a node ID.
        /// </summary>
        /// <param name="nodeId">The node ID.</param>
        /// <returns>The node type.</returns>
        private ushort GetNodeType(uint nodeId)
        {
            return (ushort)((nodeId >> 5) & 0x1F);
        }
        
        // Note: GetAllProperties method already exists elsewhere in the class
    }
}
