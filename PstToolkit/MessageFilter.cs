using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PstToolkit
{
    /// <summary>
    /// Filtering options for message properties
    /// </summary>
    public enum FilterOperator
    {
        /// <summary>
        /// Contains the specified text (case-insensitive)
        /// </summary>
        Contains,
        
        /// <summary>
        /// Starts with the specified text (case-insensitive)
        /// </summary>
        StartsWith,
        
        /// <summary>
        /// Ends with the specified text (case-insensitive)
        /// </summary>
        EndsWith,
        
        /// <summary>
        /// Equals the specified text (case-insensitive)
        /// </summary>
        Equals,
        
        /// <summary>
        /// Matches the specified regular expression pattern
        /// </summary>
        RegexMatch,
        
        /// <summary>
        /// Greater than the specified value (for dates and numeric values)
        /// </summary>
        GreaterThan,
        
        /// <summary>
        /// Less than the specified value (for dates and numeric values)
        /// </summary>
        LessThan,
        
        /// <summary>
        /// Between the specified range (for dates and numeric values)
        /// </summary>
        Between,
        
        /// <summary>
        /// Has any value (not null or empty)
        /// </summary>
        HasValue,
        
        /// <summary>
        /// Is in a list of values
        /// </summary>
        InList
    }
    
    /// <summary>
    /// Represents a filter condition for message properties
    /// </summary>
    public class FilterCondition
    {
        /// <summary>
        /// The property to filter by (e.g., "Subject", "SenderEmail", "ReceivedDate")
        /// </summary>
        public string Property { get; set; }
        
        /// <summary>
        /// The filter operator to apply
        /// </summary>
        public FilterOperator Operator { get; set; }
        
        /// <summary>
        /// The value to compare against
        /// </summary>
        public object? Value { get; set; }
        
        /// <summary>
        /// The second value for range operations (e.g., Between)
        /// </summary>
        public object? Value2 { get; set; }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="FilterCondition"/> class
        /// </summary>
        /// <param name="property">The property to filter by</param>
        /// <param name="operator">The filter operator to apply</param>
        /// <param name="value">The value to compare against</param>
        public FilterCondition(string property, FilterOperator @operator, object? value)
        {
            Property = property;
            Operator = @operator;
            Value = value;
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="FilterCondition"/> class with a second value for range operations
        /// </summary>
        /// <param name="property">The property to filter by</param>
        /// <param name="operator">The filter operator to apply</param>
        /// <param name="value">The first value to compare against</param>
        /// <param name="value2">The second value for range operations</param>
        public FilterCondition(string property, FilterOperator @operator, object? value, object? value2)
            : this(property, @operator, value)
        {
            Value2 = value2;
        }
    }
    
    /// <summary>
    /// Represents a group of filter conditions with a logical operator
    /// </summary>
    public class FilterGroup
    {
        /// <summary>
        /// The conditions in this group
        /// </summary>
        public List<FilterCondition> Conditions { get; set; } = new List<FilterCondition>();
        
        /// <summary>
        /// Whether to combine conditions with AND (true) or OR (false)
        /// </summary>
        public bool UseAnd { get; set; } = true;
        
        /// <summary>
        /// Nested filter groups for complex filtering
        /// </summary>
        public List<FilterGroup> NestedGroups { get; set; } = new List<FilterGroup>();
        
        /// <summary>
        /// Whether to combine nested groups with AND (true) or OR (false)
        /// </summary>
        public bool CombineNestedWithAnd { get; set; } = true;
        
        /// <summary>
        /// Adds a condition to this filter group
        /// </summary>
        /// <param name="condition">The condition to add</param>
        /// <returns>This filter group for chaining</returns>
        public FilterGroup AddCondition(FilterCondition condition)
        {
            Conditions.Add(condition);
            return this;
        }
        
        /// <summary>
        /// Adds a nested filter group to this filter group
        /// </summary>
        /// <param name="nestedGroup">The nested group to add</param>
        /// <returns>This filter group for chaining</returns>
        public FilterGroup AddNestedGroup(FilterGroup nestedGroup)
        {
            NestedGroups.Add(nestedGroup);
            return this;
        }
    }
    
    /// <summary>
    /// A flexible filtering system for PST messages based on various criteria
    /// </summary>
    public class MessageFilter
    {
        private readonly FilterGroup _rootGroup = new FilterGroup();
        
        /// <summary>
        /// Creates a new instance of the MessageFilter class
        /// </summary>
        public MessageFilter()
        {
        }
        
        /// <summary>
        /// Creates a new instance of the MessageFilter class with an initial filter condition
        /// </summary>
        /// <param name="property">The property to filter by</param>
        /// <param name="operator">The filter operator to apply</param>
        /// <param name="value">The value to compare against</param>
        public MessageFilter(string property, FilterOperator @operator, object? value)
        {
            _rootGroup.AddCondition(new FilterCondition(property, @operator, value));
        }
        
        /// <summary>
        /// Creates a new filter group for complex filtering
        /// </summary>
        /// <param name="useAnd">Whether to combine conditions with AND (true) or OR (false)</param>
        /// <returns>A new filter group</returns>
        public FilterGroup CreateGroup(bool useAnd = true)
        {
            var group = new FilterGroup { UseAnd = useAnd };
            _rootGroup.AddNestedGroup(group);
            return group;
        }
        
        /// <summary>
        /// Adds a filter condition to the root group
        /// </summary>
        /// <param name="property">The property to filter by</param>
        /// <param name="operator">The filter operator to apply</param>
        /// <param name="value">The value to compare against</param>
        /// <returns>This message filter for chaining</returns>
        public MessageFilter AddCondition(string property, FilterOperator @operator, object? value)
        {
            _rootGroup.AddCondition(new FilterCondition(property, @operator, value));
            return this;
        }
        
        /// <summary>
        /// Adds a filter condition with a range to the root group
        /// </summary>
        /// <param name="property">The property to filter by</param>
        /// <param name="operator">The filter operator to apply</param>
        /// <param name="value">The first value to compare against</param>
        /// <param name="value2">The second value for range operations</param>
        /// <returns>This message filter for chaining</returns>
        public MessageFilter AddCondition(string property, FilterOperator @operator, object? value, object? value2)
        {
            _rootGroup.AddCondition(new FilterCondition(property, @operator, value, value2));
            return this;
        }
        
        /// <summary>
        /// Sets whether to combine conditions in the root group with AND (true) or OR (false)
        /// </summary>
        /// <param name="useAnd">Whether to use AND (true) or OR (false)</param>
        /// <returns>This message filter for chaining</returns>
        public MessageFilter CombineWithAnd(bool useAnd = true)
        {
            _rootGroup.UseAnd = useAnd;
            return this;
        }
        
        /// <summary>
        /// Checks if a message matches this filter
        /// </summary>
        /// <param name="message">The message to check</param>
        /// <returns>True if the message matches the filter, false otherwise</returns>
        public bool Matches(PstMessage message)
        {
            return MatchesGroup(message, _rootGroup);
        }
        
        private bool MatchesGroup(PstMessage message, FilterGroup group)
        {
            // If there are no conditions or nested groups, the message matches by default
            if (group.Conditions.Count == 0 && group.NestedGroups.Count == 0)
                return true;
            
            // Check conditions in this group
            bool conditionsMatch = group.Conditions.Count == 0;
            
            if (group.Conditions.Count > 0)
            {
                if (group.UseAnd)
                {
                    // All conditions must match (AND)
                    conditionsMatch = group.Conditions.All(c => MatchesCondition(message, c));
                }
                else
                {
                    // At least one condition must match (OR)
                    conditionsMatch = group.Conditions.Any(c => MatchesCondition(message, c));
                }
            }
            
            // Check nested groups
            bool nestedGroupsMatch = group.NestedGroups.Count == 0;
            
            if (group.NestedGroups.Count > 0)
            {
                if (group.CombineNestedWithAnd)
                {
                    // All nested groups must match (AND)
                    nestedGroupsMatch = group.NestedGroups.All(g => MatchesGroup(message, g));
                }
                else
                {
                    // At least one nested group must match (OR)
                    nestedGroupsMatch = group.NestedGroups.Any(g => MatchesGroup(message, g));
                }
            }
            
            // Combine the results based on the group's UseAnd property
            return group.UseAnd 
                ? (conditionsMatch && nestedGroupsMatch) 
                : (conditionsMatch || nestedGroupsMatch);
        }
        
        private bool MatchesCondition(PstMessage message, FilterCondition condition)
        {
            // Get the property value from the message
            object? propertyValue = GetPropertyValue(message, condition.Property);
            
            // If the property value is null, it only matches HasValue (which will be false)
            if (propertyValue == null)
                return condition.Operator == FilterOperator.HasValue && (condition.Value is bool b && !b);
                
            // Handle different operators
            switch (condition.Operator)
            {
                case FilterOperator.Contains:
                    return propertyValue.ToString()?.IndexOf(condition.Value?.ToString() ?? "", StringComparison.OrdinalIgnoreCase) >= 0;
                    
                case FilterOperator.StartsWith:
                    return propertyValue.ToString()?.StartsWith(condition.Value?.ToString() ?? "", StringComparison.OrdinalIgnoreCase) == true;
                    
                case FilterOperator.EndsWith:
                    return propertyValue.ToString()?.EndsWith(condition.Value?.ToString() ?? "", StringComparison.OrdinalIgnoreCase) == true;
                    
                case FilterOperator.Equals:
                    return string.Equals(propertyValue.ToString(), condition.Value?.ToString(), StringComparison.OrdinalIgnoreCase);
                    
                case FilterOperator.RegexMatch:
                    return condition.Value != null && 
                           Regex.IsMatch(propertyValue.ToString() ?? "", condition.Value.ToString() ?? "");
                    
                case FilterOperator.GreaterThan:
                    return CompareValues(propertyValue, condition.Value) > 0;
                    
                case FilterOperator.LessThan:
                    return CompareValues(propertyValue, condition.Value) < 0;
                    
                case FilterOperator.Between:
                    return CompareValues(propertyValue, condition.Value) >= 0 && 
                           CompareValues(propertyValue, condition.Value2) <= 0;
                    
                case FilterOperator.HasValue:
                    bool hasValue = !string.IsNullOrEmpty(propertyValue.ToString());
                    return condition.Value is bool checkHasValue ? hasValue == checkHasValue : hasValue;
                    
                case FilterOperator.InList:
                    if (condition.Value is IEnumerable<object> list)
                    {
                        return list.Any(item => string.Equals(propertyValue.ToString(), 
                                        item.ToString(), StringComparison.OrdinalIgnoreCase));
                    }
                    else if (condition.Value is string csv)
                    {
                        string[] items = csv.Split(',');
                        return items.Any(item => string.Equals(propertyValue.ToString(), 
                                        item.Trim(), StringComparison.OrdinalIgnoreCase));
                    }
                    return false;
                    
                default:
                    return false;
            }
        }
        
        private object? GetPropertyValue(PstMessage message, string propertyName)
        {
            // Handle common message properties
            switch (propertyName.ToLowerInvariant())
            {
                case "subject":
                    return message.Subject;
                    
                case "sendername":
                case "sender.name":
                case "sender_name":
                case "from.name":
                case "fromname":
                    return message.SenderName;
                    
                case "senderemail":
                case "sender.email":
                case "sender_email":
                case "from.email":
                case "fromemail":
                    return message.SenderEmail;
                
                case "recipients":
                case "to":
                    return message.Recipients;
                    
                case "body":
                case "bodytext":
                case "body_text":
                case "messagetext":
                case "message_text":
                    return message.BodyText;
                    
                case "receiveddate":
                case "received_date":
                case "received":
                case "date":
                    return message.ReceivedDate;
                    
                case "sentdate":
                case "sent_date":
                case "sent":
                    return message.SentDate;
                    
                case "importance":
                    return message.Importance;
                    
                case "hasattachments":
                case "has_attachments":
                    return message.HasAttachments;
                    
                case "size":
                case "messagesize":
                case "message_size":
                    return message.Size;
                    
                default:
                    // For custom properties, we don't have direct access
                    // This would need to be implemented based on the property context
                    // of the message if needed
                    return null;
            }
        }
        
        private int CompareValues(object? value1, object? value2)
        {
            if (value1 == null && value2 == null)
                return 0;
                
            if (value1 == null)
                return -1;
                
            if (value2 == null)
                return 1;
                
            // Handle DateTime comparison
            if (value1 is DateTime dt1 && value2 is string dt2Str)
            {
                if (DateTime.TryParse(dt2Str, out DateTime dt2))
                    return dt1.CompareTo(dt2);
            }
            else if (value1 is DateTime dt1b && value2 is DateTime dt2b)
            {
                return dt1b.CompareTo(dt2b);
            }
            
            // Handle numeric comparison
            if (value1 is IComparable comparable1 && value2 is IConvertible convertible2)
            {
                try
                {
                    Type type1 = value1.GetType();
                    object converted2 = Convert.ChangeType(value2, type1);
                    return comparable1.CompareTo(converted2);
                }
                catch
                {
                    // If conversion fails, fall back to string comparison
                }
            }
            
            // Fall back to string comparison
            return string.Compare(value1.ToString(), value2?.ToString(), StringComparison.OrdinalIgnoreCase);
        }
    }
}