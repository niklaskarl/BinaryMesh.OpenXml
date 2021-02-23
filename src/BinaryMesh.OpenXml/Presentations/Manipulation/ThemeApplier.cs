using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace BinaryMesh.OpenXml.Presentations
{
    internal static class ThemeApplier
    {
        internal static void ApplyTheme(PresentationPart presentationPart, Drawing.Theme theme)
        {
            string themeId = null;
            if (presentationPart.ThemePart != null)
            {
                themeId = presentationPart.GetIdOfPart(presentationPart.ThemePart);
                presentationPart.DeletePart(themeId);
            }
            else
            {
                themeId = presentationPart.GetNextRelationshipId();
            }

            ThemePart themePart = presentationPart.AddNewPart<ThemePart>(themeId);
            theme.Save(themePart);
        }

        internal static void ApplyTableStyles(PresentationPart presentationPart, Drawing.TableStyleList tableStyleList)
        {
            string tableStylesId = null;
            if (presentationPart.TableStylesPart != null)
            {
                tableStylesId = presentationPart.GetIdOfPart(presentationPart.TableStylesPart);
                presentationPart.DeletePart(tableStylesId);
            }
            else
            {
                tableStylesId = presentationPart.GetNextRelationshipId();
            }

            TableStylesPart tableStylesPart = presentationPart.AddNewPart<TableStylesPart>(tableStylesId);
            tableStyleList.Save(tableStylesPart);
        }

        internal static void AppendSlideMaster(PresentationPart presentationPart, SlideMaster slideMaster, IEnumerable<SlideLayout> slideLayouts)
        {
            // append a new SlideMasterPart
            SlideMasterPart slideMasterPart = presentationPart.AddNewPartDefaultId<SlideMasterPart>();

            // clear SlideLayoutIdList and rebuild it
            slideMaster.SlideLayoutIdList = new SlideLayoutIdList();

            foreach (SlideLayout slideLayout in slideLayouts)
            {
                SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPartDefaultId<SlideLayoutPart>(out string slideLayoutPartId);
                slideLayoutPart.CreateRelationshipToPartDefaultId(slideMasterPart);
                slideLayout.Save(slideLayoutPart);

                // add to SlideLayoutIdList
                slideMaster.SlideLayoutIdList.AppendChild(new SlideLayoutId()
                {
                    Id = slideMaster.SlideLayoutIdList.Elements<SlideLayoutId>().Select(sl => sl.Id.Value).DefaultIfEmpty(2147483648u).Max() + 1,
                    RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
                });
            }

            slideMasterPart.CreateRelationshipToPartDefaultId(presentationPart.ThemePart);
            slideMaster.Save(slideMasterPart);

            AddSlideMasterToSlideMasterIdList(presentationPart, slideMasterPart);
        }

        private static void AddSlideMasterToSlideMasterIdList(PresentationPart presentationPart, SlideMasterPart slideMasterPart)
        {
            Presentation presentation = presentationPart.Presentation;
            SlideMasterIdList slideMasterIdList = presentationPart.Presentation.SlideMasterIdList;
            slideMasterIdList.AppendChild(new SlideMasterId()
            {
                Id = slideMasterIdList.Elements<SlideMasterId>().Select(sm => sm.Id.Value).DefaultIfEmpty(2147483647u).Max() + 1,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

            presentation.Save(presentationPart);
        }

        internal static void AppendSlideMasterFromPart(PresentationPart presentationPart, PresentationPart sourcePresentationPart, SlideMasterPart themeSlideMasterPart)
        {
            SlideMasterPart slideMasterPart = presentationPart.AddPart(themeSlideMasterPart, presentationPart.GetNextRelationshipId());
            AddSlideMasterToSlideMasterIdList(presentationPart, slideMasterPart);
            /*IDictionary<OpenXmlPart, OpenXmlPart> mapping = new Dictionary<OpenXmlPart, OpenXmlPart>()
            {
                { sourcePresentationPart, presentationPart }
            };

            CopyPartRecursive(presentationPart, sourcePresentationPart.ThemePart, presentationPart.GetNextRelationshipId(), mapping);*/
        }

        private static void CopyPartRecursive(OpenXmlPart destinationParent, OpenXmlPart part, string id, IDictionary<OpenXmlPart, OpenXmlPart> mapping)
        {
            if (part is IFixedContentTypePart)
            {
                OpenXmlPart parent = part.GetParentParts().Select(p => mapping.TryGetValue(p, out OpenXmlPart mappedParent) ? mappedParent : null).FirstOrDefault(p => p != null);

                if (mapping.TryGetValue(part, out OpenXmlPart existingPart))
                {
                    destinationParent.CreateRelationshipToPart(existingPart);
                }
                else
                {
                    MethodInfo method = typeof(OpenXmlPart).GetTypeInfo().GetMethod(nameof(OpenXmlPart.AddNewPart), new Type[] { typeof(string) }).MakeGenericMethod(part.GetType());
                    OpenXmlPart newPart = method.Invoke(parent, new object[] { id }) as OpenXmlPart;
                    mapping.Add(part, newPart);

                    using (Stream data = part.GetStream())
                    {
                        newPart.FeedData(data);
                    }

                    foreach (IdPartPair relationship in part.Parts)
                    {
                        CopyPartRecursive(newPart, relationship.OpenXmlPart, relationship.RelationshipId, mapping);
                    }
                }
            }
        }

        internal static SlideLayout[] ExtractSlideLayouts(SlideMasterPart slideMasterPart)
        {
            return slideMasterPart.SlideMaster.SlideLayoutIdList
                .Elements<SlideLayoutId>()
                .Select(id => (slideMasterPart.GetPartById(id.RelationshipId) as SlideLayoutPart).SlideLayout)
                .ToArray();
        }
    }
}
