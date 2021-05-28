using System;
using System.Collections.Immutable;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml
{
    public struct OpenXmlPresetGeometry
    {
        private readonly ShapeTypeValues shapeType;

        private readonly ImmutableArray<AdjustValue> adjustValues;

        public OpenXmlPresetGeometry(ShapeTypeValues shapeType)
        {
            this.shapeType = shapeType;
            this.adjustValues = ImmutableArray<AdjustValue>.Empty;
        }

        public OpenXmlPresetGeometry(ShapeTypeValues shapeType, ImmutableArray<AdjustValue> adjustValues)
        {
            this.shapeType = shapeType;
            this.adjustValues = adjustValues;
        }

        public static OpenXmlPresetGeometry BuildChevron(OpenXmlUnit adjust)
        {
            ImmutableArray<AdjustValue>.Builder adjustValues = ImmutableArray.CreateBuilder<AdjustValue>();

            return new OpenXmlPresetGeometry(ShapeTypeValues.Chevron, ImmutableArray.Create<AdjustValue>(new AdjustValue("adj", $"val {(long)adjust}")));
        }

        public ShapeTypeValues ShapeType => this.shapeType;

        public ImmutableArray<AdjustValue> AdjustValues => this.adjustValues;

        public struct AdjustValue
        {
            private readonly string name;

            private readonly string formula;

            public AdjustValue(string name, string formula)
            {
                this.name = name;
                this.formula = formula;
            }

            public string Name => this.name;

            public string Formula => this.formula;
        }
    }
}
