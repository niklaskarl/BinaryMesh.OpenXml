using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Shared;

namespace BinaryMesh.OpenXml.Styles.Internal
{
    internal class OpenXmlStrokeStyle<TFluent> : IStrokeStyle<TFluent>
    {
        protected readonly TFluent result;

        private readonly ElementGenerator<Outline> outlineGenerator;

        public OpenXmlStrokeStyle(TFluent result, ElementGenerator<Outline> outlineGenerator)
        {
            this.result = result;
            this.outlineGenerator = outlineGenerator;
        }

        public TFluent SetStroke(OpenXmlColor color)
        {
            Outline outline = this.outlineGenerator(true);
            this.InitializeOutline(outline);

            outline.RemoveAllChildren<NoFill>();
            outline.RemoveAllChildren<SolidFill>();
            outline.RemoveAllChildren<GradientFill>();
            outline.RemoveAllChildren<BlipFill>();
            outline.RemoveAllChildren<PatternFill>();
            outline.RemoveAllChildren<GroupFill>();

            outline.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));

            return this.result;
        }

        public TFluent SetStrokeWidth(double pt)
        {
            Outline outline = this.outlineGenerator(true);
            this.InitializeOutline(outline);
    
            outline.Width = (int)(pt * 12700);

            return this.result;
        }

        public TFluent RemoveStrokeDash()
        {
            Outline outline = this.outlineGenerator(true);
            this.InitializeOutline(outline);
    
            outline.RemoveAllChildren<PresetDash>();
            outline.RemoveAllChildren<CustomDash>();

            return this.result;
        }

        public TFluent SetStrokeDash(PresetLineDashValues value)
        {
            Outline outline = this.outlineGenerator(true);
            this.InitializeOutline(outline);
    
            outline.RemoveAllChildren<PresetDash>();
            outline.RemoveAllChildren<CustomDash>();

            outline.AppendChild(new PresetDash()
            {
                Val = value
            });

            return this.result;
        }

        private void InitializeOutline(Outline outline)
        {
            if (!(outline.Width?.HasValue ?? false))
            {
                outline.Width = 12700;
            }
        }
    }
}
