using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations.Internal.Mixins
{
    internal class OpenXmlVisualStyle<TElement, TFluent> : IVisualStyle<TFluent>
        where TElement : IOpenXmlShapeElement, TFluent
    {
        protected readonly TElement element;

        public OpenXmlVisualStyle(TElement element)
        {
            this.element = element;
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

            return this.element;
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

            return this.element;
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

            return this.element;
        }
    }
}
