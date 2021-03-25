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

        IConnectionVisual PrependConnectionVisual(string name);

        IConnectionVisual InsertConnectionVisual(string name, int index);

        IConnectionVisual AppendConnectionVisual(string name);

        IChartVisual PrependChartVisual(string name);

        IChartVisual InsertChartVisual(string name, int index);

        IChartVisual AppendChartVisual(string name);

        ITableVisual PrependTableVisual(string name);

        ITableVisual InsertTableVisual(string name, int index);

        ITableVisual AppendTableVisual(string name);
    }
}
