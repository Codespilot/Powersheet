using System;
using System.Globalization;
using System.Linq.Expressions;

namespace Nerosoft.Powersheet
{
    /// <summary>
    /// 表格选项扩展
    /// </summary>
    public static class SheetHandleOptionsExtensions
    {
        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="options"></param>
        /// <param name="nameExpression"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static SheetHandleOptions AddMapProfile<T>(this SheetHandleOptions options, Expression<Func<T, object>> nameExpression, string columnName)
        {
            var name = ExpressionHelper.GetPropertyName(nameExpression);
            return options.AddMapProfile(name, columnName);
        }

        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="options"></param>
        /// <param name="nameExpression"></param>
        /// <param name="columnName"></param>
        /// <param name="valueConvert"></param>
        /// <returns></returns>
        public static SheetHandleOptions AddMapProfile<T>(this SheetHandleOptions options, Expression<Func<T, object>> nameExpression, string columnName, Func<object, CultureInfo, object> valueConvert)
        {
            var name = ExpressionHelper.GetPropertyName(nameExpression);
            return options.AddMapProfile(name, columnName, valueConvert);
        }

        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="options"></param>
        /// <param name="nameExpression"></param>
        /// <param name="columnName"></param>
        /// <param name="valueConverter"></param>
        /// <returns></returns>
        public static SheetHandleOptions AddMapProfile<T>(this SheetHandleOptions options, Expression<Func<T, object>> nameExpression, string columnName, ICellValueConverter valueConverter)
        {
            var name = ExpressionHelper.GetPropertyName(nameExpression);
            return options.AddMapProfile(name, columnName, valueConverter);
        }

        /// <summary>
        /// 添加映射配置
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="options"></param>
        /// <param name="nameExpression"></param>
        /// <param name="columnName"></param>
        /// <param name="valueConverterType"></param>
        /// <returns></returns>
        public static SheetHandleOptions AddMapProfile<T>(this SheetHandleOptions options, Expression<Func<T, object>> nameExpression, string columnName, Type valueConverterType)
        {
            var name = ExpressionHelper.GetPropertyName(nameExpression);
            return options.AddMapProfile(name, columnName, valueConverterType);
        }


        /// <summary>
        /// 校验选项内容
        /// </summary>
        /// <param name="options"></param>
        public static void Validate(this SheetHandleOptions options)
        {
            if (options == null)
            {
                throw new NullReferenceException("Options is null.");
            }

            if (options.FirstColumnNumber < 1)
            {
                throw new InvalidOperationException("FirstColumnNumber must greater than 0.");
            }

            if (options.HeaderRowNumber < 1)
            {
                throw new InvalidOperationException("HeaderRowNumber must greater than 0.");
            }
        }
    }
}