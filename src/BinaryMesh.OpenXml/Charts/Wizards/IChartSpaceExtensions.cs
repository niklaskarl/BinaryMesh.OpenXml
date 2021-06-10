using System;

namespace BinaryMesh.OpenXml.Charts.Wizards
{
    public static class IChartSpaceExtensions
    {
        public static BurndownChartWizard BuildBurndownChart(this IChartSpace chartSpace, BurndownChartData data)
        {
            return BurndownChartWizard.BuildBurndownChart(chartSpace, data);
        }
    }
}
