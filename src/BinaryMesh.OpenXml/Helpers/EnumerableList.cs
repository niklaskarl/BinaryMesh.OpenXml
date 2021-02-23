using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace BinaryMesh.OpenXml.Helpers
{
    internal sealed class EnumerableList<TEnum, TSelect> : IReadOnlyList<TSelect>, IReadOnlyCollection<TSelect>, IEnumerable<TSelect>, IEnumerable
    {
        private readonly IEnumerable<TEnum> enumerable;

        private readonly Func<TEnum, TSelect> selector;

        public EnumerableList(IEnumerable<TEnum> enumerable, Func<TEnum, TSelect> selector)
        {
            this.enumerable = enumerable;
            this.selector = selector;
        }

        public TSelect this[int index] => this.selector(this.enumerable.Skip(index).First());

        public int Count => this.enumerable.Count();

        public IEnumerator<TSelect> GetEnumerator()
        {
            return enumerable.Select(selector).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
