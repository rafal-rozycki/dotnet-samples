using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Custom.Libraries.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace Custom.Libraries.Logging
{
    /// <summary>
    /// Takes an arbitrary object and is able to crawl the object to output the fields on the object into a log format.
    /// </summary>
    public static class LogFormatter
    {
        /// <summary>
        /// The beginning character for a representation of an object.
        /// </summary>
        public const string OBJECT_BEGINS_WITH_CHARACTER = "{";

        /// <summary>
        /// The ending character for a representation of an object.
        /// </summary>
        public const string OBJECT_ENDS_WITH_CHARACTER = "}";

        /// <summary>
        /// The beginning character for a representation of an array.
        /// </summary>
        public const string ARRAY_BEGINS_WITH_CHARACTER = "[";

        /// <summary>
        /// The ending character for a representation of an array.
        /// </summary>
        public const string ARRAY_ENDS_WITH_CHARACTER = "]";

        /// <summary>
        /// The separator character for name/value pairs (between name and value).
        /// </summary>
        public const string NAME_VALUE_SEPARATOR_CHARACTER = ":";
        
        /// <summary>
        /// The separator character for between name/value pairs.
        /// </summary>
        public const string VALUE_PAIR_SEPARATOR_CHARACTER = ",";

        /// <summary>
        /// The null representation used by the log formatter when it encounters a null object at the root level.
        /// </summary>
        public const string NULL_OBJECT_REPRESENTATION = OBJECT_BEGINS_WITH_CHARACTER + OBJECT_ENDS_WITH_CHARACTER;

        /// <summary>
        /// The null representation used by the log formatter when it encounters a null object nested inside another object.
        /// </summary>
        public const string NULL_REPRESENTATION = "null";

        /// <summary>
        /// The empty representation used by the log formatter when it encounters an empty array.
        /// </summary>
        public const string EMPTY_ARRAY_REPRESENTATION = ARRAY_BEGINS_WITH_CHARACTER + ARRAY_ENDS_WITH_CHARACTER;

#if DEBUG
        /// <summary>
        /// The default indentation amount.
        /// </summary>
        public const string LOG_INDENT = "  ";

        /// <summary>
        /// The default spacing used between name/value pairs
        /// </summary>
        public static string NameValueWhiteSpace = " ";

        /// <summary>
        /// The default line end characters.
        /// </summary>
        public static string LineEnd = System.Environment.NewLine;
#else
        /// <summary>
        /// The default indentation amount.
        /// </summary>
        public const string LOG_INDENT = "";

        /// <summary>
        /// The default spacing used between name/value pairs
        /// </summary>
        public static string NameValueWhiteSpace = String.Empty;

        /// <summary>
        /// The default line end characters.
        /// </summary>
        public static string LineEnd = String.Empty;
#endif

        /// <summary>
        /// Converts an object to a string representation in JSON format.
        /// </summary>
        /// <param name="objectToFormat">The object to format.</param>
        /// <param name="outerIndent">The outer indent.</param>
        /// <returns>System.String. If the object is null, returns NULL_OBJECT_REPRESENTATION. If an error occurs, returns an empty string.</returns>
        /// <remarks>
        /// This method uses reflection to traverse and serialize each member on the passed object.
        /// </remarks>
        public static string ToLog(this object objectToFormat, string outerIndent = "")
        {
            string result = String.Empty;
#if DEBUG
            bool validateJson = false;
#endif
            try
            {
                // Base scenario: Null object passed in
                if (objectToFormat == null)
                {
                    result = NULL_OBJECT_REPRESENTATION;
                    return result;
                }

                Type type = objectToFormat.GetType();

                // If it's a primitive type, string, or DateTime, just return it as string.
                // XElement is also handled as a string type since we don't want to traverse the entire tree.
                if (type.IsPrimitive || type == typeof(string) || type == typeof(DateTime) || type == typeof(XElement))
                {
                    result = FormatValueType(objectToFormat);
                    return result;
                }

                // For enumerable types, call ToLog recursively on each object
                IEnumerable objectAsEnumerable = objectToFormat as IEnumerable;
                if (objectAsEnumerable != null)
                {
                    result = EnumerableToLogImpl(objectAsEnumerable);
#if DEBUG
                    validateJson = true;
#endif
                    return result;
                }

                // Otherwise, get the members and loop through them
                var members = GetMembers(objectToFormat);
                members = members.OrderBy(m => m.Name).ToArray(); //order by name
                result = ToLogImpl(outerIndent, members, objectToFormat);
#if DEBUG
                validateJson = true;
#endif
                return result;
            }
            catch (Exception)
            {
                return String.Empty;
            }
#if DEBUG
            // The finally{} block is only for development, where we want to make sure that we're serializing to valid JSON format
            finally
            {
                Exception ex;
                if (validateJson && !TryValidateJson(result, out ex))
                {
                    log4net.ILog logger = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
                    logger.Warn(String.Format("LogFormatter.ToLog: Result string is not valid JSON: {0}", result), ex);
                }
            }
#endif
        }
        
        /// <summary>
        /// Converts an object to a string representation in JSON format (concrete implementation).
        /// </summary>
        /// <param name="outerIndent">The outer indent.</param>
        /// <param name="members">The members.</param>
        /// <param name="objectToFormat">The object to format.</param>
        /// <returns>System.String.</returns>
        private static string ToLogImpl(string outerIndent, MemberInfo[] members, object objectToFormat)
        {
            if (objectToFormat == null) { return NULL_OBJECT_REPRESENTATION; }

            string innerIndent = LOG_INDENT + outerIndent;

            StringBuilder sb = new StringBuilder();

            sb.Append(outerIndent + OBJECT_BEGINS_WITH_CHARACTER + LineEnd);

            if (!members.Any())
            {
                return objectToFormat.ToString();
            }

            MemberInfo lastField = members.Last();
            foreach (MemberInfo field in members)
            {
                FieldMaskAttribute fieldMask = field.GetCustomAttributeCached<FieldMaskAttribute>();
                bool isLastField = lastField.Equals(field);

                if (fieldMask == null)
                {
                    Type currentType = null;
                    object value = null;

                    if (field is PropertyInfo)
                    {
                        currentType = ((PropertyInfo)field).PropertyType;
                        value = ((PropertyInfo)field).GetValue(objectToFormat);
                    }
                    if (field is FieldInfo)
                    {
                        currentType = ((FieldInfo)field).FieldType;
                        value = ((FieldInfo)field).GetValue(objectToFormat);
                    }

                    if (value == null)
                    {
                        sb.Append(FormatNameValuePair(field.Name, NULL_REPRESENTATION, isLastField, innerIndent) + LineEnd);
                        continue;
                    }

                    if (
                        currentType != typeof(string) //strings are a prominent exception to enumerable types we want to explicitly enumerate over
                        && !(currentType.IsArray && (currentType.GetElementType().IsValueType || currentType.GetElementType() == typeof(string))) //also arrays that are value types (since .NET handles this)
                        && !(currentType.GetInterfaces().Contains(typeof(IEnumerable)) && currentType.IsGenericType && (currentType.GetGenericArguments()[0].IsValueType || currentType.GetGenericArguments()[0] == typeof(string))) //as well as IEnumerable<T> where T is a value type (since .NET handles this)
                        && (currentType.IsArray || currentType.GetInterfaces().Contains(typeof(IEnumerable))))
                    {
                        sb.Append(FormatNameValuePair(field.Name, EnumerableToLogImpl((IEnumerable)value, innerIndent), isLastField, innerIndent));
                    }
                    /* since the above handles non value type lists and arrays, let's handle every other type of list and array here, except string of course */
                    else if (currentType != typeof(string) && ((currentType.IsArray) || (currentType.GetInterfaces().Contains(typeof(IEnumerable)))))
                    {
                        StringBuilder iSb = new StringBuilder();
                        iSb.Append(innerIndent + OBJECT_BEGINS_WITH_CHARACTER + NameValueWhiteSpace);
                        foreach (var item in (IEnumerable)value)
                        {
                            iSb.Append(String.Format("{0}" + VALUE_PAIR_SEPARATOR_CHARACTER + NameValueWhiteSpace, item));
                        }
                        iSb.Append(NameValueWhiteSpace + OBJECT_ENDS_WITH_CHARACTER);

                        string enumerableString = iSb.ToString();
                        int lastCommaIndex = enumerableString.LastIndexOf(VALUE_PAIR_SEPARATOR_CHARACTER + NameValueWhiteSpace, StringComparison.OrdinalIgnoreCase);
                        if (lastCommaIndex > 0)
                        {
                            // NameValueWhiteSpace.Length + 1 is used so that we remove the correct number of characters
                            // depending on the whitespace used in DEBUG/RELEASE builds. Otherwise it strips out the trailing ARRAY_ENDS_WITH_CHARACTER
                            enumerableString = enumerableString.Remove(lastCommaIndex, NameValueWhiteSpace.Length + 1); // Remove the last comma pair
                        }
                        sb.Append(FormatNameValuePair(field.Name, enumerableString, isLastField, innerIndent) + LineEnd);
                    }
                    else if (currentType.IsValueType || currentType == typeof(string))
                    {
                        value = FormatValueType(value);
                        sb.Append(FormatNameValuePair(field.Name, value, isLastField, innerIndent) + LineEnd);
                    }
                    else
                    {
                        if (value != objectToFormat) //avoid circular references where some objects have a field that contains itself
                        {
                            sb.Append(FormatNameValuePair(field.Name, value.ToLog(innerIndent), isLastField, innerIndent));
                        }
                    }
                }
                else
                {
                    FieldMasker fm = new FieldMasker(fieldMask);
                    string maskedValue = fm.Mask(field, objectToFormat);
                    sb.Append(FormatNameValuePair(field.Name, maskedValue == null ? NULL_REPRESENTATION : "\"" + maskedValue + "\"", isLastField, innerIndent) + LineEnd);
                }
            }

            sb.Append(outerIndent + OBJECT_ENDS_WITH_CHARACTER + LineEnd);

            string result = sb.ToString();
            return result;
        }

        /// <summary>
        /// Enumerable to log implementation.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="outerIndent">The outer indent.</param>
        /// <returns>System.String.</returns>
        private static string EnumerableToLogImpl(IEnumerable list, string outerIndent = "")
        {
            if (list == null) { return NULL_OBJECT_REPRESENTATION; }

            string innerIndent = LOG_INDENT + outerIndent;

            StringBuilder sb = new StringBuilder();

            sb.Append(outerIndent + ARRAY_BEGINS_WITH_CHARACTER + LineEnd);

            foreach (object item in list)
            {
                if (item == null)
                {
                    sb.Append(NULL_REPRESENTATION);
                }
                else
                {
                    sb.Append(ToLog(item, innerIndent));
                }

                sb.Append(VALUE_PAIR_SEPARATOR_CHARACTER + NameValueWhiteSpace);
            }

            sb.Append(outerIndent + ARRAY_ENDS_WITH_CHARACTER + LineEnd);

            string sbAsString = sb.ToString();
            int lastCommaIndex = sbAsString.LastIndexOf(VALUE_PAIR_SEPARATOR_CHARACTER + NameValueWhiteSpace, StringComparison.OrdinalIgnoreCase);
            if (lastCommaIndex > 0)
            {
                // NameValueWhiteSpace.Length + 1 is used so that we remove the correct number of characters
                // depending on the whitespace used in DEBUG/RELEASE builds. Otherwise it strips out the trailing ARRAY_ENDS_WITH_CHARACTER
                sbAsString = sbAsString.Remove(lastCommaIndex, NameValueWhiteSpace.Length + 1); // Remove the last comma pair
            }

            return sbAsString;
        }

        /// <summary>
        /// Gets the members for a given object.
        /// </summary>
        /// <param name="_object">The object.</param>
        /// <returns>MemberInfo[].</returns>
        private static MemberInfo[] GetMembers(object _object)
        {
            Type t = _object.GetType();
            var members = t.FindMembers(MemberTypes.Field | MemberTypes.Property, BindingFlags.Instance | BindingFlags.Public, null, null);
            return members;
        }

        /// <summary>
        /// Gets the members for a given object.
        /// </summary>
        /// <typeparam name="TProjection">The type of the t projection.</typeparam>
        /// <returns>MemberInfo[].</returns>
        private static MemberInfo[] GetMembers<TProjection>()
        {
            var members = typeof(TProjection).FindMembers(MemberTypes.Field | MemberTypes.Property, BindingFlags.Instance | BindingFlags.Public, null, null);
            members = members.OrderBy(m => m.Name).ToArray(); //order by name
            return members;
        }

        /// <summary>
        /// Formats a name/value pair in JSON format.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        /// <param name="isLastField">if set to <c>true</c> this is the last field, and a trailing comma will not be added.</param>
        /// <param name="innerIndent">The inner indent spacing.</param>
        /// <returns>System.String.</returns>
        /// <remarks>
        /// If quotes are needed around the value, they must be explicitly passed in.
        /// </remarks>
        private static string FormatNameValuePair(string name, object value, bool isLastField, string innerIndent)
        {
            return isLastField
                ? String.Format(innerIndent + "\"{0}\"" + NAME_VALUE_SEPARATOR_CHARACTER + NameValueWhiteSpace + "{1}", name, value)
                : String.Format(innerIndent + "\"{0}\"" + NAME_VALUE_SEPARATOR_CHARACTER + NameValueWhiteSpace + "{1}" + VALUE_PAIR_SEPARATOR_CHARACTER, name, value);
        }

        /// <summary>
        /// Formats a value type in JSON format, adding quotes where needed. Numeric and boolean values do not get quotes.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>System.String.</returns>
        private static string FormatValueType(object value)
        {
            bool useQuotes = true; // Use quotes on value by default (string, DateTime, etc.)
            double valueAsDouble; // No quotes on numeric values
            bool valueAsBool; // No quotes on boolean values
            DateTime valueAsDateTime;

            if (value is string)
            {
                useQuotes = true;
            }
            else if (Double.TryParse(value.ToString(), out valueAsDouble))
            {
                value = valueAsDouble;
                useQuotes = false;
            }
            else if (Boolean.TryParse(value.ToString(), out valueAsBool))
            {
                value = valueAsBool.ToString().ToLower();
                useQuotes = false;
            }
            else if (DateTime.TryParse(value.ToString(), out valueAsDateTime))
            {
                value = valueAsDateTime.ToString("o"); // Convert to ISO 8601 format (ex. "2012-04-23T18:25:43.511Z")
            }

            if (useQuotes)
            {
                value = "\"" + value + "\"";
            }
            return value.ToString();
        }

#if DEBUG
        /// <summary>
        /// Determines whether a given string is in valid JSON format.
        /// </summary>
        /// <param name="s">The string input.</param>
        /// <param name="e">The exception.</param>
        /// <returns><c>true</c> if the string is in valid JSON format; otherwise, <c>false</c>.</returns>
        /// <remarks>
        /// Lifted from http://stackoverflow.com/a/14977915
        /// </remarks>
        private static bool TryValidateJson(string s, out Exception e)
        {
            e = null;
            s = s.Trim();

            // The reason to add checks for { or [ is based on the fact that JToken.Parse would parse the values such as "1234" or "'a string'" as a valid token
            if ((!s.StartsWith(OBJECT_BEGINS_WITH_CHARACTER) || !s.EndsWith(OBJECT_ENDS_WITH_CHARACTER)) && (!s.StartsWith(ARRAY_BEGINS_WITH_CHARACTER) || !s.EndsWith(ARRAY_ENDS_WITH_CHARACTER)))
            {
                return false;
            }

            try
            {
                JToken.Parse(s); // If Parse works, we're good to go
                return true;
            }
            // Any exception means the format is not valid JSON
            catch (JsonReaderException jsonEx)
            {
                e = jsonEx;
                return false;
            }
            catch (Exception ex)
            {
                e = ex;
                return false;
            }
        }
#endif
    }
}
