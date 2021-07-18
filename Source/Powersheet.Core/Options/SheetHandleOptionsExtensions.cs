using System;
using System.Globalization;
using System.Linq.Expressions;

namespace Nerosoft.Powersheet
{
    public static class SheetHandleOptionsExtensions
    {
        public static SheetHandleOptions UseMapProfile(this SheetHandleOptions options, string name, string columnName)
        {
            if (options == null)
            {
                throw new NullReferenceException("Options is null.");
            }

            options.AddMapProfile(name, columnName);
            return options;
        }

        public static SheetHandleOptions UseMapProfile(this SheetHandleOptions options, string name, string columnName, Func<object, CultureInfo, object> valueConvert)
        {
            if (options == null)
            {
                throw new NullReferenceException("Options is null.");
            }

            options.AddMapProfile(name, columnName, valueConvert);
            return options;
        }

        public static SheetHandleOptions UseMapProfile(this SheetHandleOptions options, string name, string columnName, ICellValueConverter valueConverter)
        {
            if (options == null)
            {
                throw new NullReferenceException("Options is null.");
            }

            options.AddMapProfile(name, columnName, valueConverter);
            return options;
        }

        public static SheetHandleOptions UseMapProfile(this SheetHandleOptions options, string name, string columnName, Type valueConverterType)
        {
            if (options == null)
            {
                throw new NullReferenceException("Options is null.");
            }

            if (valueConverterType.IsSubclassOf(typeof(ICellValueConverter)))
            {
                throw new InvalidOperationException($"The value converter must implements {nameof(ICellValueConverter)}.");
            }

            var valueConverter = (ICellValueConverter) Activator.CreateInstance(valueConverterType);
            options.AddMapProfile(name, columnName, valueConverter);
            return options;
        }

        public static SheetHandleOptions UseMapProfile<T>(this SheetHandleOptions options, Expression<Func<T>> nameExpression, string columnName)
        {
            var name = ExpressionHelper.GetPropertyName(nameExpression);
            return options.UseMapProfile(name, columnName);
        }

        public static SheetHandleOptions UseMapProfile<T>(this SheetHandleOptions options, Expression<Func<T>> nameExpression, string columnName, Func<object, CultureInfo, object> valueConvert)
        {
            var name = ExpressionHelper.GetPropertyName(nameExpression);
            return options.UseMapProfile(name, columnName, valueConvert);
        }

        public static SheetHandleOptions UseMapProfile<T>(this SheetHandleOptions options, Expression<Func<T>> nameExpression, string columnName, ICellValueConverter valueConverter)
        {
            var name = ExpressionHelper.GetPropertyName(nameExpression);
            return options.UseMapProfile(name, columnName, valueConverter);
        }

        public static SheetHandleOptions UseMapProfile<T>(this SheetHandleOptions options, Expression<Func<T>> nameExpression, string columnName, Type valueConverterType)
        {
            var name = ExpressionHelper.GetPropertyName(nameExpression);
            return options.UseMapProfile(name, columnName, valueConverterType);
        }

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