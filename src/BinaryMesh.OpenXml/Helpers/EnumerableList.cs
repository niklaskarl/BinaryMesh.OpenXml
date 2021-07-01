using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace BinaryMesh.OpenXml.Helpers
{
    internal abstract class BaseEnumerableList<TEnum, TSelect> : IReadOnlyList<TSelect>, IReadOnlyCollection<TSelect>, IEnumerable<TSelect>, IEnumerable
    {
        public BaseEnumerableList()
        {
        }

        protected abstract IEnumerable<TEnum> Enumerable { get; }

        protected abstract Func<TEnum, TSelect> Selector { get; }

        public TSelect this[int index] => this.Selector(this.Enumerable.Skip(index).First());

        public int Count => this.Enumerable.Count();

        public IEnumerator<TSelect> GetEnumerator()
        {
            return Enumerable.Select(Selector).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }

    internal sealed class EnumerableList<TEnum, TSelect> : BaseEnumerableList<TEnum, TSelect>
    {
        private readonly IEnumerable<TEnum> enumerable;

        private readonly Func<TEnum, TSelect> selector;

        public EnumerableList(IEnumerable<TEnum> enumerable, Func<TEnum, TSelect> selector)
        {
            this.enumerable = enumerable;
            this.selector = selector;
        }

        protected override IEnumerable<TEnum> Enumerable => this.enumerable;

        protected override Func<TEnum, TSelect> Selector => this.selector;
    }
}
