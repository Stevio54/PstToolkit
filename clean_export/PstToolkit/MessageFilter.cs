using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PstToolkit
{
    /// <summary>
    /// Defines the logical operator to use when combining multiple filters.
    /// </summary>
    public enum FilterLogic
    {
        /// <summary>
        /// All filters must match (AND).
        /// </summary>
        All,
        
        /// <summary>
        /// At least one filter must match (OR).
        /// </summary>
        Any
    }
    
    /// <summary>
    /// Provides filtering capabilities for messages in PST files.
    /// </summary>
    public class MessageFilter
    {
        private List<Func<PstMessage, bool>> _filters = new List<Func<PstMessage, bool>>();
        private FilterLogic _logic = FilterLogic.All; // Default to AND logic
        
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
        public MessageFilter AddCondition(string property, FilterOperator op, object? value)
        {
            if (property == null)
            {
                throw new ArgumentNullException(nameof(property));
            }
            
            property = property.ToLowerInvariant();
            
            // Handle null value safely
            string stringValue = value?.ToString() ?? string.Empty;
            
            switch (property)
            {
                case "subject":
                    AddSubjectCondition(op, stringValue);
                    break;
                case "sender":
                case "from":
                case "senderemail":
                    AddSenderCondition(op, stringValue);
                    break;
                case "date":
                case "sentdate":
                    if (value is DateTime dateValue)
                    {
                        AddDateCondition(op, dateValue);
                    }
                    else if (DateTime.TryParse(stringValue, out DateTime parsedDate))
                    {
                        AddDateCondition(op, parsedDate);
                    }
                    break;
                case "size":
                    if (value is long longValue)
                    {
                        AddSizeCondition(op, longValue);
                    }
                    else if (long.TryParse(stringValue, out long parsedSize))
                    {
                        AddSizeCondition(op, parsedSize);
                    }
                    break;
                case "hasattachments":
                    if (value is bool boolValue)
                    {
                        _filters.Add(m => m.HasAttachments == boolValue);
                    }
                    else if (bool.TryParse(stringValue, out bool parsedBool))
                    {
                        _filters.Add(m => m.HasAttachments == parsedBool);
                    }
                    break;
                case "body":
                case "content":
                    AddBodyCondition(op, stringValue);
                    break;
                default:
                    // For unknown properties, try generic matching on any property
                    AddCustomCondition(property, op, stringValue);
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
        public MessageFilter AddCondition(string property, FilterOperator op, object? value1, object? value2)
        {
            if (property == null)
            {
                throw new ArgumentNullException(nameof(property));
            }
            
            if (op != FilterOperator.Between)
            {
                throw new ArgumentException("Range filtering requires the Between operator");
            }
            
            property = property.ToLowerInvariant();
            
            // Handle null values safely
            string stringValue1 = value1?.ToString() ?? string.Empty;
            string stringValue2 = value2?.ToString() ?? string.Empty;
            
            switch (property)
            {
                case "date":
                case "sentdate":
                    if (value1 is DateTime startDate && value2 is DateTime endDate)
                    {
                        _filters.Add(m => m.SentDate >= startDate && m.SentDate <= endDate);
                    }
                    else 
                    {
                        DateTime parsedStartDate, parsedEndDate;
                        bool startParsed = DateTime.TryParse(stringValue1, out parsedStartDate);
                        bool endParsed = DateTime.TryParse(stringValue2, out parsedEndDate);
                        
                        if (startParsed && endParsed)
                        {
                            _filters.Add(m => m.SentDate >= parsedStartDate && m.SentDate <= parsedEndDate);
                        }
                    }
                    break;
                case "size":
                    if (value1 is long minSize && value2 is long maxSize)
                    {
                        _filters.Add(m => m.Size >= minSize && m.Size <= maxSize);
                    }
                    else 
                    {
                        long parsedMinSize, parsedMaxSize;
                        bool minParsed = long.TryParse(stringValue1, out parsedMinSize);
                        bool maxParsed = long.TryParse(stringValue2, out parsedMaxSize);
                        
                        if (minParsed && maxParsed)
                        {
                            _filters.Add(m => m.Size >= parsedMinSize && m.Size <= parsedMaxSize);
                        }
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
            if (string.IsNullOrEmpty(value))
            {
                // Don't add any filter for empty values
                return;
            }
            
            switch (op)
            {
                case FilterOperator.Equals:
                    _filters.Add(m => 
                        (m.SenderEmail != null && m.SenderEmail.Equals(value, StringComparison.OrdinalIgnoreCase)) || 
                        (m.SenderName != null && m.SenderName.Equals(value, StringComparison.OrdinalIgnoreCase)));
                    break;
                case FilterOperator.Contains:
                    _filters.Add(m => 
                        (m.SenderEmail != null && m.SenderEmail.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0) || 
                        (m.SenderName != null && m.SenderName.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0));
                    break;
                case FilterOperator.StartsWith:
                    _filters.Add(m => 
                        (m.SenderEmail != null && m.SenderEmail.StartsWith(value, StringComparison.OrdinalIgnoreCase)) || 
                        (m.SenderName != null && m.SenderName.StartsWith(value, StringComparison.OrdinalIgnoreCase)));
                    break;
                case FilterOperator.EndsWith:
                    _filters.Add(m => 
                        (m.SenderEmail != null && m.SenderEmail.EndsWith(value, StringComparison.OrdinalIgnoreCase)) || 
                        (m.SenderName != null && m.SenderName.EndsWith(value, StringComparison.OrdinalIgnoreCase)));
                    break;
                case FilterOperator.RegexMatch:
                    try
                    {
                        var regex = new Regex(value, RegexOptions.IgnoreCase);
                        _filters.Add(m => 
                            (m.SenderEmail != null && regex.IsMatch(m.SenderEmail)) || 
                            (m.SenderName != null && regex.IsMatch(m.SenderName)));
                    }
                    catch
                    {
                        // Invalid regex pattern, use contains as fallback
                        _filters.Add(m => 
                            (m.SenderEmail != null && m.SenderEmail.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0) || 
                            (m.SenderName != null && m.SenderName.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0));
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
            if (string.IsNullOrEmpty(value))
            {
                // Don't add any filter for empty values
                return;
            }
            
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
            if (string.IsNullOrEmpty(value))
            {
                // Don't add any filter for empty values
                return;
            }
            
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
        /// Sets the logic operator to use when combining multiple filter conditions.
        /// </summary>
        /// <param name="logic">The filter logic to use (All/AND or Any/OR).</param>
        /// <returns>The filter instance for method chaining.</returns>
        public MessageFilter SetLogic(FilterLogic logic)
        {
            _logic = logic;
            return this;
        }
        
        /// <summary>
        /// Applies all filters to the given message and returns true if the message passes the filters.
        /// </summary>
        /// <param name="message">The message to filter.</param>
        /// <returns>True if the message passes the filters based on the logic operator, false otherwise.</returns>
        public bool Matches(PstMessage message)
        {
            // If no filters, all messages match
            if (_filters.Count == 0)
            {
                return true;
            }
            
            // Apply filter based on the current logic operator
            if (_logic == FilterLogic.All)
            {
                // AND logic - all filters must match
                foreach (var filter in _filters)
                {
                    if (!filter(message))
                    {
                        return false;
                    }
                }
                return true;
            }
            else
            {
                // OR logic - at least one filter must match
                foreach (var filter in _filters)
                {
                    if (filter(message))
                    {
                        return true;
                    }
                }
                return false;
            }
        }
        
        /// <summary>
        /// Filters a collection of messages and returns only those that match the filters.
        /// </summary>
        /// <param name="messages">The messages to filter.</param>
        /// <returns>A filtered collection of messages.</returns>
        public IEnumerable<PstMessage> Apply(IEnumerable<PstMessage> messages)
        {
            // If no filters, return all messages immediately
            if (_filters.Count == 0)
            {
                return messages;
            }
            
            // For small collections, use simple filtering
            var messageList = messages as IList<PstMessage> ?? messages.ToList();
            if (messageList.Count < 100)
            {
                return messageList.Where(Matches);
            }
            
            // For larger collections, use parallel processing with batch size optimization
            return ApplyParallel(messageList);
        }
        
        /// <summary>
        /// Applies filters in parallel for better performance on large message sets.
        /// </summary>
        /// <param name="messages">The messages to filter.</param>
        /// <returns>A filtered collection of messages.</returns>
        private IEnumerable<PstMessage> ApplyParallel(IList<PstMessage> messages)
        {
            // Optimize batch size based on message count
            int batchSize = DetermineBatchSize(messages.Count);
            
            // Use chunking for better memory usage
            var result = new ConcurrentBag<PstMessage>();
            
            Parallel.ForEach(
                Partitioner.Create(0, messages.Count, batchSize),
                range =>
                {
                    for (int i = range.Item1; i < range.Item2; i++)
                    {
                        var message = messages[i];
                        if (Matches(message))
                        {
                            result.Add(message);
                        }
                    }
                });
                
            return result;
        }
        
        /// <summary>
        /// Determines optimal batch size based on collection size.
        /// </summary>
        private int DetermineBatchSize(int totalCount)
        {
            // For very large sets (50K+ messages), use larger batches
            if (totalCount > 50000)
                return 1000;
                
            // For large sets (10K-50K messages), use medium batches
            if (totalCount > 10000)
                return 500;
                
            // For medium sets (1K-10K messages), use small batches
            if (totalCount > 1000)
                return 250;
                
            // For smaller sets, use very small batches
            return 100;
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