using System;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace BinaryMesh.OpenXml.Presentations.Internal
{
    internal sealed class ChartSpaceRef
    {
        private readonly ChartSpace chartSpace;

        public ChartSpaceRef()
        {
            this.chartSpace.GetFirstChild<Chart>().PlotArea.GetType();
        }

        public void AddPieChart()
        {
            this.chartSpace.GetFirstChild<Chart>().PlotArea
                .AppendChildFluent(
                    new PieChart()
                    {
                        VaryColors = new VaryColors() { Val = true }
                    }
                        .AppendChildFluent(
                            new DataLabels()
                            {
                                
                            }
                        )
                );
        }
    }
}
