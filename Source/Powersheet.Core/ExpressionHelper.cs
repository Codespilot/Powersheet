using System;
using System.Collections.Generic;
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

		internal static MemberInfo GetMemberAccess(this LambdaExpression memberAccessExpression)
			=> GetInternalMemberAccess<MemberInfo>(memberAccessExpression);

		private static TMemberInfo GetInternalMemberAccess<TMemberInfo>(this LambdaExpression memberAccessExpression)
			where TMemberInfo : MemberInfo
		{

			var parameterExpression = memberAccessExpression.Parameters[0];
			var memberInfo = parameterExpression.MatchSimpleMemberAccess<TMemberInfo>(memberAccessExpression.Body);

			if (memberInfo == null)
			{
				throw new ArgumentException();
			}

			var declaringType = memberInfo.DeclaringType;
			var parameterType = parameterExpression.Type;

			if (declaringType != null
				&& declaringType != parameterType
				&& declaringType.IsInterface
				&& declaringType.IsAssignableFrom(parameterType)
				&& memberInfo is PropertyInfo propertyInfo)
			{
				var propertyGetter = propertyInfo.GetMethod;
				var interfaceMapping = parameterType.GetTypeInfo().GetRuntimeInterfaceMap(declaringType);
				var index = Array.FindIndex(interfaceMapping.InterfaceMethods, p => p.Equals(propertyGetter));
				var targetMethod = interfaceMapping.TargetMethods[index];
				foreach (var runtimeProperty in parameterType.GetRuntimeProperties())
				{
					if (targetMethod.Equals(runtimeProperty.GetMethod))
					{
						return (TMemberInfo)(object)runtimeProperty;
					}
				}
			}

			return memberInfo;
		}

		internal static TMemberInfo MatchSimpleMemberAccess<TMemberInfo>(
				this Expression parameterExpression,
				Expression memberAccessExpression)
				where TMemberInfo : MemberInfo
		{
			var memberInfos = MatchMemberAccess<TMemberInfo>(parameterExpression, memberAccessExpression);

			return memberInfos?.Count == 1 ? memberInfos[0] : null;
		}

		private static IReadOnlyList<TMemberInfo> MatchMemberAccess<TMemberInfo>(
			this Expression parameterExpression,
			Expression memberAccessExpression)
			where TMemberInfo : MemberInfo
		{
			var memberInfos = new List<TMemberInfo>();

			MemberExpression memberExpression;
			var unwrappedExpression = RemoveTypeAs(RemoveConvert(memberAccessExpression));
			do
			{
				memberExpression = unwrappedExpression as MemberExpression;

				if (memberExpression?.Member is not TMemberInfo memberInfo)
				{
					return null;
				}

				memberInfos.Insert(0, memberInfo);

				unwrappedExpression = RemoveTypeAs(RemoveConvert(memberExpression.Expression));
			}
			while (unwrappedExpression != parameterExpression);

			return memberInfos;
		}

		internal static Expression RemoveTypeAs(this Expression expression)
		{
			while (expression?.NodeType == ExpressionType.TypeAs)
			{
				expression = ((UnaryExpression)RemoveConvert(expression)).Operand;
			}

			return expression;
		}

		private static Expression RemoveConvert(Expression expression)
		{
			if (expression is UnaryExpression unaryExpression
				&& (expression.NodeType == ExpressionType.Convert
					|| expression.NodeType == ExpressionType.ConvertChecked))
			{
				return RemoveConvert(unaryExpression.Operand);
			}

			return expression;
		}
	}
}