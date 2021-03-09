using System;
using System.Collections.Generic;

namespace BinaryMesh.OpenXml.Presentations
{
    public interface IVisualContainer
    {
        KeyedReadOnlyList<string, IVisual> Visuals { get; }

        IShapeVisual PrependShapeVisual(string name);

        IShapeVisual InsertShapeVisual(string name, int index);

        IShapeVisual AppendShapeVisual(string name);

        IChartVisual PrependChartVisual(string name);

        IChartVisual InsertChartVisual(string name, int index);

        IChartVisual AppendChartVisual(string name);

        ITableVisual PrependTableVisual(string name);

        ITableVisual InsertTableVisual(string name, int index);

        ITableVisual AppendTableVisual(string name);
    }
}
