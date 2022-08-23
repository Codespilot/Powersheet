using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq.Expressions;

namespace Nerosoft.Powersheet
{
    public class SheetHandleOptionsBuilder<TItem>
        where TItem : class
    {
        private readonly SheetItemTypeBuilder<TItem> _itemTypeBuilder = new();

        protected SheetHandleOptionsBuilder(SheetHandleOptions options)
        {
            Options = options;
        }

        protected SheetHandleOptions Options { get; }

        protected SheetHandleOptionsBuilder<TItem> ConfigureProfile(Action<SheetItemTypeBuilder<TItem>> buildAction)
        {
            buildAction(_itemTypeBuilder);
            foreach (var profile in _itemTypeBuilder.Profiles)
            {
                if (profile.Ignored)
                {
                    Options.IgnoreName(profile.PropertyName);
                }
                else
                {
                    Options.AddMapProfile(profile.PropertyName, profile.ColumnName, profile.ValueConverter);
                }
            }

            return this;
        }

        public SheetHandleOptionsBuilder<TItem> ConfigureOptions(Action<SheetHandleOptions> action)
        {
            action(Options);
            return this;
        }

        public static SheetHandleOptionsBuilder<TOptions, TItem> For<TOptions>()
            where TOptions : SheetHandleOptions, new()
        {
            var options = new TOptions();
            return new SheetHandleOptionsBuilder<TOptions, TItem>(options);
        }

        public static SheetHandleOptionsBuilder<TOptions, TItem> For<TOptions>(TOptions options)
            where TOptions : SheetHandleOptions
        {
            return new SheetHandleOptionsBuilder<TOptions, TItem>(options);
        }
    }

    public class SheetHandleOptionsBuilder<TOptions, TItem> : SheetHandleOptionsBuilder<TItem>
        where TOptions : SheetHandleOptions
        where TItem : class
    {
        internal SheetHandleOptionsBuilder(TOptions options)
            : base(options)
        {
        }

        public new TOptions Options => (TOptions)base.Options;

        public new SheetHandleOptionsBuilder<TOptions, TItem> ConfigureProfile(Action<SheetItemTypeBuilder<TItem>> buildAction)
        {
            return (SheetHandleOptionsBuilder<TOptions, TItem>)base.ConfigureProfile(buildAction);
        }

        public SheetHandleOptionsBuilder<TOptions, TItem> ConfigureOptions(Action<TOptions> action)
        {
            action(Options);
            return this;
        }
    }

    public sealed class SheetItemTypeBuilder<TItem>
    {
        private readonly Dictionary<string, SheetItemProfileBuilder> _profiles = new();

        public SheetItemProfileBuilder<TProperty> Property<TProperty>(Expression<Func<TItem, TProperty>> propertyExpression)
        {
            var propertyName = propertyExpression.GetMemberAccess().Name;
            var builder = new SheetItemProfileBuilder<TProperty>(propertyName);
            _profiles.Add(propertyName, builder);
            return builder;
        }

        internal IEnumerable<SheetItemProfileBuilder> Profiles => _profiles.Values;
    }

    public class SheetItemProfileBuilder
    {
        internal SheetItemProfileBuilder(string propertyName)
        {
            PropertyName = propertyName;
            ColumnName = propertyName;
        }

        public bool Ignored { get; protected set; }

        public string PropertyName { get; }

        public string ColumnName { get; protected set; }

        public IValueConverter ValueConverter { get; protected set; }
    }

    public class SheetItemProfileBuilder<TProperty> : SheetItemProfileBuilder
    {
        internal SheetItemProfileBuilder(string propertyName)
            : base(propertyName)
        {
        }

        public SheetItemProfileBuilder<TProperty> Ignore()
        {
            Ignored = true;
            return this;
        }

        public SheetItemProfileBuilder<TProperty> HasColumnName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentNullException(nameof(name));
            }

            ColumnName = name;
            return this;
        }

        public SheetItemProfileBuilder<TProperty> HasCellValueConverter(Func<object, CultureInfo, TProperty> valueConvert)
        {
            ValueConverter ??= new EmptyValueConverter<TProperty>();

            if (ValueConverter is EmptyValueConverter<TProperty> converter)
            {
                converter.SetCellValueConverter(valueConvert);
            }

            return this;
        }

        public SheetItemProfileBuilder<TProperty> HasItemValueConverter(Func<TProperty, CultureInfo, object> valueConvert)
        {
            ValueConverter ??= new EmptyValueConverter<TProperty>();

            if (ValueConverter is EmptyValueConverter<TProperty> converter)
            {
                converter.SetItemValueConverter(valueConvert);
            }

            return this;
        }

        public SheetItemProfileBuilder<TProperty> HasValueConverter(IValueConverter converter)
        {
            ValueConverter = converter;
            return this;
        }

        public SheetItemProfileBuilder<TProperty> HasValueConverter(IValueConverter<TProperty> converter)
        {
            return HasValueConverter((IValueConverter)converter);
        }

        public SheetItemProfileBuilder<TProperty> HasValueConverter(Type converterType)
        {
            if (converterType == null)
            {
                throw new ArgumentNullException(nameof(converterType));
            }

            if (converterType.IsSubclassOf(typeof(IValueConverter)))
            {
                throw new InvalidOperationException($"The value converter must implements {nameof(IValueConverter)}.");
            }

            var valueConverter = (IValueConverter)Activator.CreateInstance(converterType);
            return HasValueConverter(valueConverter);
        }

        public SheetItemProfileBuilder<TProperty> HasValueConverter<TConverter>()
            where TConverter : class
        {
            return HasValueConverter(typeof(TConverter));
        }
    }
}