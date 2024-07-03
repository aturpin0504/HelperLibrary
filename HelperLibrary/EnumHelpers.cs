using System.ComponentModel;
using System;
using System.Reflection;

namespace HelperLibrary
{
    public static class EnumHelpers
    {
        public static TEnum GetEnumValue<TEnum>(int value) where TEnum : struct, Enum
        {
            if (!Enum.IsDefined(typeof(TEnum), value))
            {
                throw new InvalidEnumArgumentException($"The value {value} is not defined in the enum {typeof(TEnum).Name}.");
            }

            return (TEnum)Enum.ToObject(typeof(TEnum), value);
        }

        public static string GetEnumDescription<TEnum>(int value) where TEnum : struct, Enum
        {
            if (!Enum.IsDefined(typeof(TEnum), value))
            {
                throw new InvalidEnumArgumentException($"The value {value} is not defined in the enum {typeof(TEnum).Name}.");
            }

            TEnum enumValue = GetEnumValue<TEnum>(value);
            FieldInfo fieldInfo = enumValue.GetType().GetField(enumValue.ToString());
            DescriptionAttribute[] attributes = (DescriptionAttribute[])fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
            return attributes.Length > 0 ? attributes[0].Description : enumValue.ToString();
        }
    }
}
