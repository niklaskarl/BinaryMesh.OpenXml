using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Styles.Internal.Mixins;

namespace BinaryMesh.OpenXml.Styles.Internal
{
    internal class OpenXmlVisualStyle<TElement, TFluent> : IVisualStyle<TFluent>
        where TElement : IOpenXmlShapeElement
    {
        protected readonly TElement element;

        protected readonly TFluent result;

        public OpenXmlVisualStyle(TElement element, TFluent result)
        {
            this.element = element;
            this.result = result;
        }

        public TFluent SetFill(OpenXmlColor color)
        {
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            shapeProperties.RemoveAllChildren<NoFill>();
            shapeProperties.RemoveAllChildren<SolidFill>();
            shapeProperties.RemoveAllChildren<GradientFill>();
            shapeProperties.RemoveAllChildren<BlipFill>();
            shapeProperties.RemoveAllChildren<PatternFill>();
            shapeProperties.RemoveAllChildren<GroupFill>();

            shapeProperties.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));

            return this.result;
        }

        public TFluent SetNoFill()
        {
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            shapeProperties.RemoveAllChildren<NoFill>();
            shapeProperties.RemoveAllChildren<SolidFill>();
            shapeProperties.RemoveAllChildren<GradientFill>();
            shapeProperties.RemoveAllChildren<BlipFill>();
            shapeProperties.RemoveAllChildren<PatternFill>();
            shapeProperties.RemoveAllChildren<GroupFill>();

            shapeProperties.AppendChild(new NoFill());

            return this.result;
        }

        public TFluent SetStroke(OpenXmlColor color)
        {
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            Outline outline = shapeProperties.GetFirstChild<Outline>() ?? shapeProperties.AppendChild(new Outline() { Width = 12700 });
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
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            Outline outline = shapeProperties.GetFirstChild<Outline>() ?? shapeProperties.AppendChild(new Outline());
            outline.Width = (int)(pt * 12700);

            return this.result;
        }

        public TFluent RemoveStrokeDash()
        {
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            Outline outline = shapeProperties.GetFirstChild<Outline>() ?? shapeProperties.AppendChild(new Outline() { Width = 12700 });

            outline.RemoveAllChildren<PresetDash>();
            outline.RemoveAllChildren<CustomDash>();

            return this.result;
        }

        public TFluent SetStrokeDash(PresetLineDashValues value)
        {
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            Outline outline = shapeProperties.GetFirstChild<Outline>() ?? shapeProperties.AppendChild(new Outline() { Width = 12700 });

            outline.RemoveAllChildren<PresetDash>();
            outline.RemoveAllChildren<CustomDash>();

            outline.AppendChild(new PresetDash()
            {
                Val = value
            });

            return this.result;
        }

        public TFluent SetPresetGeometry(OpenXmlPresetGeometry geometry)
        {
            OpenXmlElement shapeProperties = this.element.GetOrCreateShapeProperties();
            PresetGeometry presetGeometry = shapeProperties.GetFirstChild<PresetGeometry>() ?? shapeProperties.AppendChild(new PresetGeometry());

            presetGeometry.RemoveAllChildren();
            presetGeometry.Preset = geometry.ShapeType;
            if (!geometry.AdjustValues.IsDefault)
            {
                presetGeometry.AppendChild(new AdjustValueList().AppendFluent(geometry.AdjustValues.Select(av => new ShapeGuide() { Name = av.Name, Formula = av.Formula })));
            }

            return this.result;
        }
    }
}
