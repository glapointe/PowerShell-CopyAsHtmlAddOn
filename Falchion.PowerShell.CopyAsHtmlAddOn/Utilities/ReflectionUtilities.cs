using System;
using System.Reflection;

namespace Falchion.PowerShell.CopyAsHtmlAddOn.Utilities
{
    internal class ReflectionUtilities
    {
        /// <summary>
        /// Gets all bindings.
        /// </summary>
        /// <value>All bindings.</value>
        internal static BindingFlags AllBindings
        {
            get
            {
                return BindingFlags.CreateInstance |
                BindingFlags.FlattenHierarchy |
                BindingFlags.GetField |
                BindingFlags.GetProperty |
                BindingFlags.IgnoreCase |
                BindingFlags.Instance |
                BindingFlags.InvokeMethod |
                BindingFlags.NonPublic |
                BindingFlags.Public |
                BindingFlags.SetField |
                BindingFlags.SetProperty |
                BindingFlags.Static;
            }
        }

        /// <summary>
        /// Gets the property value.
        /// </summary>
        /// <param name="o">The object whose property is to be retrieved.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <returns></returns>
        internal static object GetPropertyValue(object o, string propertyName)
        {
            return o.GetType().GetProperty(propertyName, AllBindings).GetValue(o, null);
        }
        internal static object GetFieldValue(object o, string fieldName)
        {
            return o.GetType().GetField(fieldName, AllBindings).GetValue(o);
        }

        /// <summary>
        /// Executes the method.
        /// </summary>
        /// <param name="objectType">The type.</param>
        /// <param name="methodName">Name of the method.</param>
        /// <param name="parameterTypes">The parameter types.</param>
        /// <param name="parameterValues">The parameter values.</param>
        /// <returns></returns>
        internal static object ExecuteMethod(Type objectType, string methodName, Type[] parameterTypes, object[] parameterValues)
        {
            return ExecuteMethod(objectType, null, methodName, parameterTypes, parameterValues);
        }
        /// <summary>
        /// Executes the method.
        /// </summary>
        /// <param name="obj">The obj.</param>
        /// <param name="methodName">Name of the method.</param>
        /// <param name="parameterTypes">The parameter types.</param>
        /// <param name="parameterValues">The parameter values.</param>
        /// <returns></returns>
        internal static object ExecuteMethod(object obj, string methodName, Type[] parameterTypes, object[] parameterValues)
        {
            return ExecuteMethod(obj.GetType(), obj, methodName, parameterTypes, parameterValues);
        }

        /// <summary>
        /// Executes the method.
        /// </summary>
        /// <param name="objectType">The type.</param>
        /// <param name="obj">The obj.</param>
        /// <param name="methodName">Name of the method.</param>
        /// <param name="parameterTypes">The parameter types.</param>
        /// <param name="parameterValues">The parameter values.</param>
        /// <returns></returns>
        internal static object ExecuteMethod(Type objectType, object obj, string methodName, Type[] parameterTypes, object[] parameterValues)
        {
            MethodInfo methodInfo = objectType.GetMethod(methodName, AllBindings, null, parameterTypes, null);
            try
            {
                return methodInfo.Invoke(obj, parameterValues);
            }
            catch (TargetInvocationException ex)
            {
                // Get and throw the real exception.
                throw ex.InnerException;
            }
        }
    }
}
