using System;
using System.Collections.Concurrent;
using System.Reflection;

namespace MorphosPowerPointAddIn.Utilities
{
    internal static class ComFontAccessorCache
    {
        private static readonly ConcurrentDictionary<string, PropertyInfo> GetterCache =
            new ConcurrentDictionary<string, PropertyInfo>(StringComparer.Ordinal);

        private static readonly ConcurrentDictionary<string, PropertyInfo> SetterCache =
            new ConcurrentDictionary<string, PropertyInfo>(StringComparer.Ordinal);

        public static bool TryGetString(object target, string propertyName, out string value)
        {
            value = string.Empty;
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                var property = GetterCache.GetOrAdd(BuildKey(target.GetType(), propertyName), _ => ResolveProperty(target.GetType(), propertyName));
                if (property == null || !property.CanRead)
                {
                    return false;
                }

                value = Convert.ToString(property.GetValue(target, null));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool TrySetString(object target, string propertyName, string value)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                var property = SetterCache.GetOrAdd(BuildKey(target.GetType(), propertyName), _ => ResolveProperty(target.GetType(), propertyName));
                if (property == null || !property.CanWrite)
                {
                    return false;
                }

                property.SetValue(target, value, null);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static PropertyInfo ResolveProperty(Type targetType, string propertyName)
        {
            return targetType == null
                ? null
                : targetType.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
        }

        private static string BuildKey(Type targetType, string propertyName)
        {
            return (targetType == null ? string.Empty : targetType.FullName ?? targetType.Name) + "|" + propertyName;
        }
    }
}
