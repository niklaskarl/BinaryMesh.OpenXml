using System;
using System.IO;
using BinaryMesh.OpenXml.Presentations.Internal;

namespace BinaryMesh.OpenXml.Presentations
{
    public static class PresentationFactory
    {
        public static IPresentation CreatePresentation()
        {
            return new PresentationRef(null);
        }

        public static IPresentation CreatePresentation(string template)
        {
            using (Stream templateStream = new FileStream(template, FileMode.Open, FileAccess.Read))
            {
                return new PresentationRef(templateStream);
            }
        }

        public static IPresentation CreatePresentation(Stream templateStream)
        {
            return new PresentationRef(templateStream);
        }
    }
}
