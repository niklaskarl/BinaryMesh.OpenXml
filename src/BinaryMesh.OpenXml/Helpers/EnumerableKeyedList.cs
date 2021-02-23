using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace BinaryMesh.OpenXml.Helpers
{
    internal sealed class EnumerableKeyedList<TEnum, TKey, TValue> : KeyedReadOnlyList<TKey, TValue>
    {
        private readonly IEnumerable<TEnum> enumerable;

        private readonly Func<TEnum, TKey> keySelector;

        private readonly Func<TEnum, TValue> valueSelector;

        public EnumerableKeyedList(IEnumerable<TEnum> enumerable, Func<TEnum, TKey> keySelector, Func<TEnum, TValue> valueSelector)
        {
            this.enumerable = enumerable;
            this.keySelector = keySelector;
            this.valueSelector = valueSelector;
        }

        public override TValue this[int index] => this.valueSelector(this.enumerable.Skip(index).First());

        public override TValue this[TKey key] => this.valueSelector(this.enumerable.FirstOrDefault(item => object.Equals(keySelector(item), key)));

        public override int Count => this.enumerable.Count();

        public override IEnumerable<TKey> Keys => enumerable.Select(keySelector);

        public override IEnumerable<TValue> Values => this;

        public override bool ContainsKey(TKey key)
        {
            return this.enumerable.Any(item => object.Equals(keySelector(item), key));
        }

        public override bool TryGetValue(TKey key, out TValue value)
        {
            foreach (TEnum result in this.enumerable.Where(item => object.Equals(keySelector(item), key)))
            {
                value = valueSelector(result);
                return true;
            }

            value = default(TValue);
            return false;
        }

        public override IEnumerator<TValue> GetEnumerator()
        {
            return enumerable.Select(valueSelector).GetEnumerator();
        }

        protected override IEnumerator<KeyValuePair<TKey, TValue>> GetDictionaryEnumerator()
        {
            return enumerable.Select(item => new KeyValuePair<TKey, TValue>(keySelector(item), valueSelector(item))).GetEnumerator();
        }
    }
}
