using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;

namespace BinaryMesh.OpenXml
{
    internal static class OpenXmlPartContainerExtensions
    {
        private static Regex relationshipIdPattern = new Regex("rId[0-9]+");

        public static TPart AddNewPartDefaultId<TPart>(this OpenXmlPartContainer parent) where TPart : OpenXmlPart, IFixedContentTypePart
        {
            return parent.AddNewPart<TPart>(parent.GetNextRelationshipId());
        }

        public static TPart AddNewPartDefaultId<TPart>(this OpenXmlPartContainer parent, out string id) where TPart : OpenXmlPart, IFixedContentTypePart
        {
            id = parent.GetNextRelationshipId();
            return parent.AddNewPart<TPart>(id);
        }

        public static string CreateRelationshipToPartDefaultId(this OpenXmlPartContainer parent, OpenXmlPart targetPart)
        {
            return parent.CreateRelationshipToPart(targetPart, parent.GetNextRelationshipId());
        }

        public static string GetNextRelationshipId(this OpenXmlPartContainer parent)
        {
            int id = parent.Parts
                .Select(p => p.RelationshipId.StartsWith("rId") ? (int.TryParse(p.RelationshipId.Substring(3), out int number) ? number : 0) : 0)
                .DefaultIfEmpty(0)
                .Max() + 1;

            return $"rId{id}";
        }
    }
}
