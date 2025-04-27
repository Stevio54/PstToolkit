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
            
            // In a real implementation, this would parse the property context
            // from the node data. We'll provide a mock implementation that
            // returns some typical properties based on the node type.
            
            try
            {
                // Get the node type from the node ID
                ushort nodeType = GetNodeType(_nodeEntry.NodeId);
                
                // Read the raw data for this node
                byte[] nodeData = _nodeEntry.ReadData(_pstFile);
                
                if (nodeData.Length == 0)
                {
                    // No data to load properties from
                    return;
                }
                
                // In a real implementation, we would parse the property context from nodeData
                // For now, initialize with mock properties based on node type
                
                switch (nodeType)
                {
                    case PstNodeTypes.NID_TYPE_FOLDER:
                        LoadFolderProperties();
                        break;
                        
                    case PstNodeTypes.NID_TYPE_MESSAGE:
                        LoadMessageProperties();
                        break;
                        
                    case PstNodeTypes.NID_TYPE_ATTACHMENT:
                        LoadAttachmentProperties();
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new PstCorruptedException("Error parsing properties", ex);
            }
        }
        
        private void LoadFolderProperties()
        {
            // Mock implementation for folder properties based on node ID
            uint nodeId = _nodeEntry.NodeId;
            
            // Create some default folder properties depending on the node ID
            if (nodeId == (_isAnsi ? 0x21u : 0x42u)) // Root folder
            {
                _properties![MakePropertyId(PropertyIds.PidTagDisplayName, PropertyType.PT_STRING8)] = "Root Folder";
                _properties[MakePropertyId(PropertyIds.PidTagContainerClass, PropertyType.PT_STRING8)] = "";
                _properties[MakePropertyId(PropertyIds.PidTagContentCount, PropertyType.PT_LONG)] = 0;
                _properties[MakePropertyId(PropertyIds.PidTagContentUnreadCount, PropertyType.PT_LONG)] = 0;
                _properties[MakePropertyId(PropertyIds.PidTagSubfolders, PropertyType.PT_BOOLEAN)] = true;
            }
            else if (nodeId == 0x122) // Typical Inbox
            {
                _properties![MakePropertyId(PropertyIds.PidTagDisplayName, PropertyType.PT_STRING8)] = "Inbox";
                _properties[MakePropertyId(PropertyIds.PidTagContainerClass, PropertyType.PT_STRING8)] = "IPM.Note";
                _properties[MakePropertyId(PropertyIds.PidTagContentCount, PropertyType.PT_LONG)] = 10;
                _properties[MakePropertyId(PropertyIds.PidTagContentUnreadCount, PropertyType.PT_LONG)] = 2;
                _properties[MakePropertyId(PropertyIds.PidTagSubfolders, PropertyType.PT_BOOLEAN)] = false;
            }
            else if (nodeId == 0x222) // Typical Sent Items
            {
                _properties![MakePropertyId(PropertyIds.PidTagDisplayName, PropertyType.PT_STRING8)] = "Sent Items";
                _properties[MakePropertyId(PropertyIds.PidTagContainerClass, PropertyType.PT_STRING8)] = "IPM.Note";
                _properties[MakePropertyId(PropertyIds.PidTagContentCount, PropertyType.PT_LONG)] = 5;
                _properties[MakePropertyId(PropertyIds.PidTagContentUnreadCount, PropertyType.PT_LONG)] = 0;
                _properties[MakePropertyId(PropertyIds.PidTagSubfolders, PropertyType.PT_BOOLEAN)] = false;
            }
            else // Generic folder
            {
                _properties![MakePropertyId(PropertyIds.PidTagDisplayName, PropertyType.PT_STRING8)] = $"Folder {nodeId:X}";
                _properties[MakePropertyId(PropertyIds.PidTagContainerClass, PropertyType.PT_STRING8)] = "IPM.Note";
                _properties[MakePropertyId(PropertyIds.PidTagContentCount, PropertyType.PT_LONG)] = 0;
                _properties[MakePropertyId(PropertyIds.PidTagContentUnreadCount, PropertyType.PT_LONG)] = 0;
                _properties[MakePropertyId(PropertyIds.PidTagSubfolders, PropertyType.PT_BOOLEAN)] = false;
            }
        }
        
        private void LoadMessageProperties()
        {
            // Mock implementation for message properties
            uint nodeId = _nodeEntry.NodeId;
            
            // Create default message properties
            _properties![MakePropertyId(PropertyIds.PidTagSubject, PropertyType.PT_STRING8)] = $"Sample Message {nodeId:X}";
            _properties[MakePropertyId(PropertyIds.PidTagSenderName, PropertyType.PT_STRING8)] = "John Doe";
            _properties[MakePropertyId(PropertyIds.PidTagSenderEmailAddress, PropertyType.PT_STRING8)] = "john.doe@example.com";
            _properties[MakePropertyId(PropertyIds.PidTagClientSubmitTime, PropertyType.PT_SYSTIME)] = DateTime.Now.AddDays(-1);
            _properties[MakePropertyId(PropertyIds.PidTagMessageDeliveryTime, PropertyType.PT_SYSTIME)] = DateTime.Now.AddDays(-1).AddMinutes(5);
            _properties[MakePropertyId(PropertyIds.PidTagMessageFlags, PropertyType.PT_LONG)] = 1; // Read
            _properties[MakePropertyId(PropertyIds.PidTagMessageSize, PropertyType.PT_LONG)] = 4096;
            _properties[MakePropertyId(PropertyIds.PidTagImportance, PropertyType.PT_LONG)] = 1; // Normal
            _properties[MakePropertyId(PropertyIds.PidTagBody, PropertyType.PT_STRING8)] = "This is a sample message body.";
        }
        
        private void LoadAttachmentProperties()
        {
            // Mock implementation for attachment properties
            uint nodeId = _nodeEntry.NodeId;
            
            // Create default attachment properties
            _properties![MakePropertyId(PropertyIds.PidTagAttachFilename, PropertyType.PT_STRING8)] = $"attachment{nodeId:X}.txt";
            _properties[MakePropertyId(PropertyIds.PidTagAttachLongFilename, PropertyType.PT_STRING8)] = $"attachment{nodeId:X}.txt";
            _properties[MakePropertyId(PropertyIds.PidTagAttachmentSize, PropertyType.PT_LONG)] = 1024;
            _properties[MakePropertyId(PropertyIds.PidTagAttachMethod, PropertyType.PT_LONG)] = 1; // Embedded
        }
        
        private ushort GetNodeType(uint nodeId)
        {
            return (ushort)((nodeId >> 5) & 0x1F);
        }
    }
}
