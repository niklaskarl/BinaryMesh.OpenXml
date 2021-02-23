using System;
using System.Collections;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml
{
    public abstract class KeyedReadOnlyList<TKey, TValue> : IReadOnlyList<TValue>, IReadOnlyDictionary<TKey, TValue>
    {
        public abstract TValue this[int index] { get; }

        public abstract TValue this[TKey key] { get; }

        public abstract int Count { get; }

        public abstract IEnumerable<TKey> Keys { get; }

        public abstract IEnumerable<TValue> Values { get; }

        public abstract bool ContainsKey(TKey key);

        public abstract bool TryGetValue(TKey key, out TValue value);

        public abstract IEnumerator<TValue> GetEnumerator();

        protected abstract IEnumerator<KeyValuePair<TKey, TValue>> GetDictionaryEnumerator();

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        IEnumerator<KeyValuePair<TKey, TValue>> IEnumerable<KeyValuePair<TKey, TValue>>.GetEnumerator()
        {
            return this.GetDictionaryEnumerator();
        }
    }
}
