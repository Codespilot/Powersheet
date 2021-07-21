using System;
using System.Linq.Expressions;
using System.Reflection;

namespace Nerosoft.Powersheet
{
    internal static class ExpressionHelper
    {
        /// <summary>
        /// Extracts the property name from a property expression.
        /// </summary>
        /// <typeparam name="T">The object type containing the property specified in the expression.</typeparam>
        /// <param name="expression">The property expression (e.g. p =&gt; p.PropertyName)</param>
        /// <returns>The name of the property.</returns>
        /// <exception cref="ArgumentNullException">Thrown if the <paramref name="expression" /> is null.</exception>
        /// <exception cref="ArgumentException">Thrown when the expression is:<br />
        /// Not a <see cref="MemberExpression" /><br />
        /// The <see cref="MemberExpression" /> does not represent a property.<br />
        /// Or, the property is static.</exception>
        internal static string GetPropertyName<T>(Expression<Func<T, object>> expression)
        {
            if (expression == null)
            {
                throw new ArgumentNullException(nameof(expression));
            }

            switch (expression.Body)
            {
                case MemberExpression memberExpression:
                    var property = memberExpression.Member as PropertyInfo;
                    if (property == null)
                    {
                        throw new ArgumentException("The member access expression does not access a property.", nameof(expression));
                    }

                    var getMethod = property.GetMethod;
                    if (getMethod.IsStatic)
                    {
                        throw new ArgumentException("The referenced property is a static property.", nameof(expression));
                    }

                    return memberExpression.Member.Name;
                case UnaryExpression unaryExpression:
                    if (unaryExpression.Operand is not MemberExpression operand)
                    {
                        throw new ArgumentException("The expression is not a member access expression.", nameof(operand));
                    }

                    return operand.Member.Name;
                default:
                    throw new InvalidOperationException();
            }
        }
    }
}