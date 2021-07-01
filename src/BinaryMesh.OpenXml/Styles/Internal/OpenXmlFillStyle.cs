using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

using BinaryMesh.OpenXml.Shared;

namespace BinaryMesh.OpenXml.Styles.Internal
{
    internal class OpenXmlFillStyle<TFluent> : IFillStyle<TFluent>
    {
        protected readonly TFluent result;

        private readonly ElementGenerator<OpenXmlElement> fillGenerator;

        public OpenXmlFillStyle(TFluent result, ElementGenerator<OpenXmlElement> fillGenerator)
        {
            this.result = result;
            this.fillGenerator = fillGenerator;
        }

        public TFluent SetFill(OpenXmlColor color)
        {
            OpenXmlElement fill = this.fillGenerator(true);
            fill.RemoveAllChildren<NoFill>();
            fill.RemoveAllChildren<SolidFill>();
            fill.RemoveAllChildren<GradientFill>();
            fill.RemoveAllChildren<BlipFill>();
            fill.RemoveAllChildren<PatternFill>();
            fill.RemoveAllChildren<GroupFill>();

            fill.AppendChild(new SolidFill().AppendChildFluent(color.CreateColorElement()));

            return this.result;
        }

        public TFluent SetNoFill()
        {
            OpenXmlElement fill = this.fillGenerator(true);
            fill.RemoveAllChildren<NoFill>();
            fill.RemoveAllChildren<SolidFill>();
            fill.RemoveAllChildren<GradientFill>();
            fill.RemoveAllChildren<BlipFill>();
            fill.RemoveAllChildren<PatternFill>();
            fill.RemoveAllChildren<GroupFill>();

            fill.AppendChild(new NoFill());

            return this.result;
        }
    }
}
