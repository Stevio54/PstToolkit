using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PstToolkit
{
    /// <summary>
    /// Provides filtering capabilities for messages in PST files.
    /// </summary>
    public class MessageFilter
    {
        private List<Func<PstMessage, bool>> _filters = new List<Func<PstMessage, bool>>();
        
        /// <summary>
        /// Initializes a new instance of the <see cref="MessageFilter"/> class.
        /// </summary>
        public MessageFilter()
        {
        }
        
        /// <summary>
        /// Adds a condition to filter messages.
        /// </summary>
        /// <param name="property">The property to filter on.</param>
        /// <param name="op">The filter operator.</param>
        /// <param name="value">The value to compare against.</param>
        /// <returns>The filter instance for method chaining.</returns>
        public MessageFilter AddCondition(string property, FilterOperator op, object value)
        {
            property = property.ToLowerInvariant();
            
            switch (property)
            {
                case "subject":
                    AddSubjectCondition(op, value.ToString());
                    break;
                case "sender":
                case "from":
                case "senderemail":
                    AddSenderCondition(op, value.ToString());
                    break;
                case "date":
                case "sentdate":
                    if (value is DateTime dateValue)
                    {
                        AddDateCondition(op, dateValue);
                    }
                    else if (DateTime.TryParse(value.ToString(), out DateTime parsedDate))
                    {
                        AddDateCondition(op, parsedDate);
                    }
                    break;
                case "size":
                    if (value is long longValue)
                    {
                        AddSizeCondition(op, longValue);
                    }
                    else if (long.TryParse(value.ToString(), out long parsedSize))
                    {
                        AddSizeCondition(op, parsedSize);
                    }
                    break;
                case "hasattachments":
                    if (value is bool boolValue)
                    {
                        _filters.Add(m => m.HasAttachments == boolValue);
                    }
                    else if (bool.TryParse(value.ToString(), out bool parsedBool))
                    {
                        _filters.Add(m => m.HasAttachments == parsedBool);
                    }
                    break;
                case "body":
                case "content":
                    AddBodyCondition(op, value.ToString());
                    break;
                default:
                    // For unknown properties, try generic matching on any property
                    AddCustomCondition(property, op, value.ToString());
                    break;
            }
            
            return this;
        }
        
        /// <summary>
        /// Adds a condition with a range to filter messages.
        /// </summary>
        /// <param name="property">The property to filter on.</param>
        /// <param name="op">The filter operator (should be Between).</param>
        /// <param name="value1">The first value of the range.</param>
        /// <param name="value2">The second value of the range.</param>
        /// <returns>The filter instance for method chaining.</returns>
        public MessageFilter AddCondition(string property, FilterOperator op, object value1, object value2)
        {
            if (op != FilterOperator.Between)
            {
                throw new ArgumentException("Range filtering requires the Between operator");
            }
            
            property = property.ToLowerInvariant();
            
            switch (property)
            {
                case "date":
                case "sentdate":
                    if (value1 is DateTime startDate && value2 is DateTime endDate)
                    {
                        _filters.Add(m => m.SentDate >= startDate && m.SentDate <= endDate);
                    }
                    break;
                case "size":
                    if (value1 is long minSize && value2 is long maxSize)
                    {
                        _filters.Add(m => m.Size >= minSize && m.Size <= maxSize);
                    }
                    else if (long.TryParse(value1.ToString(), out long parsedMinSize) && 
                             long.TryParse(value2.ToString(), out long parsedMaxSize))
                    {
                        _filters.Add(m => m.Size >= parsedMinSize && m.Size <= parsedMaxSize);
                    }
                    break;
            }
            
            return this;
        }
        
        private void AddSubjectCondition(FilterOperator op, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                // Don't add any filter for empty values
                return;
            }
            
            switch (op)
            {
                case FilterOperator.Equals:
                    _filters.Add(m => m.Subject != null && m.Subject.Equals(value, StringComparison.OrdinalIgnoreCase));
                    break;
                case FilterOperator.Contains:
                    _filters.Add(m => m.Subject != null && m.Subject.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0);
                    break;
                case FilterOperator.StartsWith:
                    _filters.Add(m => m.Subject != null && m.Subject.StartsWith(value, StringComparison.OrdinalIgnoreCase));
                    break;
                case FilterOperator.EndsWith:
                    _filters.Add(m => m.Subject != null && m.Subject.EndsWith(value, StringComparison.OrdinalIgnoreCase));
                    break;
                case FilterOperator.RegexMatch:
                    try
                    {
                        var regex = new Regex(value, RegexOptions.IgnoreCase);
                        _filters.Add(m => m.Subject != null && regex.IsMatch(m.Subject));
                    }
                    catch
                    {
                        // Invalid regex pattern, use contains as fallback
                        _filters.Add(m => m.Subject != null && m.Subject.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0);
                    }
                    break;
            }
        }
        
        private void AddSenderCondition(FilterOperator op, string value)
        {
            switch (op)
            {
                case FilterOperator.Equals:
                    _filters.Add(m => m.SenderEmail.Equals(value, StringComparison.OrdinalIgnoreCase) || 
                                     m.SenderName.Equals(value, StringComparison.OrdinalIgnoreCase));
                    break;
                case FilterOperator.Contains:
                    _filters.Add(m => m.SenderEmail.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0 || 
                                     m.SenderName.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0);
                    break;
                case FilterOperator.StartsWith:
                    _filters.Add(m => m.SenderEmail.StartsWith(value, StringComparison.OrdinalIgnoreCase) || 
                                     m.SenderName.StartsWith(value, StringComparison.OrdinalIgnoreCase));
                    break;
                case FilterOperator.EndsWith:
                    _filters.Add(m => m.SenderEmail.EndsWith(value, StringComparison.OrdinalIgnoreCase) || 
                                     m.SenderName.EndsWith(value, StringComparison.OrdinalIgnoreCase));
                    break;
                case FilterOperator.RegexMatch:
                    try
                    {
                        var regex = new Regex(value, RegexOptions.IgnoreCase);
                        _filters.Add(m => regex.IsMatch(m.SenderEmail) || regex.IsMatch(m.SenderName));
                    }
                    catch
                    {
                        // Invalid regex pattern, use contains as fallback
                        _filters.Add(m => m.SenderEmail.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0 || 
                                         m.SenderName.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0);
                    }
                    break;
            }
        }
        
        private void AddDateCondition(FilterOperator op, DateTime value)
        {
            switch (op)
            {
                case FilterOperator.Equals:
                    _filters.Add(m => m.SentDate.Date == value.Date);
                    break;
                case FilterOperator.GreaterThan:
                    _filters.Add(m => m.SentDate > value);
                    break;
                case FilterOperator.LessThan:
                    _filters.Add(m => m.SentDate < value);
                    break;
            }
        }
        
        private void AddSizeCondition(FilterOperator op, long value)
        {
            switch (op)
            {
                case FilterOperator.Equals:
                    _filters.Add(m => m.Size == value);
                    break;
                case FilterOperator.GreaterThan:
                    _filters.Add(m => m.Size > value);
                    break;
                case FilterOperator.LessThan:
                    _filters.Add(m => m.Size < value);
                    break;
            }
        }
        
        private void AddBodyCondition(FilterOperator op, string value)
        {
            switch (op)
            {
                case FilterOperator.Contains:
                    _filters.Add(m => 
                        (m.BodyText != null && m.BodyText.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (m.BodyHtml != null && m.BodyHtml.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0));
                    break;
                case FilterOperator.RegexMatch:
                    try
                    {
                        var regex = new Regex(value, RegexOptions.IgnoreCase);
                        _filters.Add(m => 
                            (m.BodyText != null && regex.IsMatch(m.BodyText)) ||
                            (m.BodyHtml != null && regex.IsMatch(m.BodyHtml)));
                    }
                    catch
                    {
                        // Invalid regex pattern, use contains as fallback
                        _filters.Add(m => 
                            (m.BodyText != null && m.BodyText.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (m.BodyHtml != null && m.BodyHtml.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0));
                    }
                    break;
            }
        }
        
        private void AddCustomCondition(string property, FilterOperator op, string value)
        {
            // This would handle custom properties or fall back to a more generic approach
            // For now, just create a loose matching filter that checks if any message field contains the value
            _filters.Add(m => 
                (m.Subject != null && m.Subject.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (m.SenderName != null && m.SenderName.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (m.SenderEmail != null && m.SenderEmail.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (m.BodyText != null && m.BodyText.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0) ||
                (m.BodyHtml != null && m.BodyHtml.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0));
        }
        
        /// <summary>
        /// Applies all filters to the given message and returns true if the message passes all filters.
        /// </summary>
        /// <param name="message">The message to filter.</param>
        /// <returns>True if the message passes all filters, false otherwise.</returns>
        public bool Matches(PstMessage message)
        {
            // If no filters, all messages match
            if (_filters.Count == 0)
            {
                return true;
            }
            
            // Test each filter
            foreach (var filter in _filters)
            {
                if (!filter(message))
                {
                    return false;
                }
            }
            
            return true;
        }
        
        /// <summary>
        /// Filters a collection of messages and returns only those that match the filters.
        /// </summary>
        /// <param name="messages">The messages to filter.</param>
        /// <returns>A filtered collection of messages.</returns>
        public IEnumerable<PstMessage> Apply(IEnumerable<PstMessage> messages)
        {
            return messages.Where(Matches);
        }
        
        /// <summary>
        /// Clears all filters from the filter chain.
        /// </summary>
        /// <returns>The filter instance for method chaining.</returns>
        public MessageFilter ClearFilters()
        {
            _filters.Clear();
            return this;
        }
    }
    
    /// <summary>
    /// Defines filter operators for message filtering.
    /// </summary>
    public enum FilterOperator
    {
        /// <summary>
        /// Equal to (case-insensitive for strings).
        /// </summary>
        Equals,
        
        /// <summary>
        /// Contains the value (case-insensitive for strings).
        /// </summary>
        Contains,
        
        /// <summary>
        /// Starts with the value (case-insensitive for strings).
        /// </summary>
        StartsWith,
        
        /// <summary>
        /// Ends with the value (case-insensitive for strings).
        /// </summary>
        EndsWith,
        
        /// <summary>
        /// Greater than the value (for dates, numbers).
        /// </summary>
        GreaterThan,
        
        /// <summary>
        /// Less than the value (for dates, numbers).
        /// </summary>
        LessThan,
        
        /// <summary>
        /// Between two values (inclusive, for dates, numbers).
        /// </summary>
        Between,
        
        /// <summary>
        /// Matches a regular expression pattern (for strings).
        /// </summary>
        RegexMatch
    }
}