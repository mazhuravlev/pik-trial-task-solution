using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Xml;
using System.Windows.Forms;
using System.Linq;
using Autodesk.Revit.Attributes;

namespace TestAddIn
{
    [Transaction(TransactionMode.ReadOnly)]
    public class TestAddIn : IExternalCommand
    {
        private const string DialogTitle = "TestAddIn";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            using (var xmlWriter = GetXmlWriter())
            {
                if (xmlWriter == null)
                {
                    return Result.Cancelled;
                }
                WriteFamilyObectListToXml(xmlWriter,GetFamilyObjectsFromDocument(commandData.Application.ActiveUIDocument.Document));
            }
            TaskDialog.Show(DialogTitle, "Export completed");
            return Result.Succeeded;
        }

        private static void WriteFamilyObectListToXml(XmlWriter xmlWriter, IEnumerable<FamilyObject> familyObjects)
        {
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement("families");
            foreach (var familyObject in familyObjects)
            {
                xmlWriter.WriteStartElement("family");
                xmlWriter.WriteElementString("name", familyObject.Name);
                xmlWriter.WriteStartElement("types");
                foreach (var type in familyObject.Types)
                {
                    xmlWriter.WriteStartElement("type");
                    xmlWriter.WriteElementString("name", type.Name);
                    xmlWriter.WriteElementString("category", type.Category.Name);
                    xmlWriter.WriteEndElement();
                }
                xmlWriter.WriteEndElement();
                xmlWriter.WriteEndElement();
            }
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndDocument();
        }

        private static IEnumerable<FamilyObject> GetFamilyObjectsFromDocument(Document document)
        {
            var familyInstanceElements = new FilteredElementCollector(document).OfClass(typeof(FamilyInstance)).ToElements();
            var familyDict = new Dictionary<string, Dictionary<string, TypeObject>>();
            foreach (var element in familyInstanceElements)
            {
                var fInstance = element as FamilyInstance;
                if (fInstance == null) continue;
                var familySymbol = fInstance.Symbol;
                if (!familyDict.ContainsKey(familySymbol.FamilyName))
                {
                    familyDict[familySymbol.FamilyName] = new Dictionary<string, TypeObject>();
                }
                if (!familyDict[familySymbol.FamilyName].ContainsKey(familySymbol.Name))
                {
                    familyDict[familySymbol.FamilyName][familySymbol.Name] = (new TypeObject
                    {
                        Name = familySymbol.Name,
                        Category = familySymbol.Category
                    });
                }
            }
            return familyDict.Select(k => new FamilyObject { Name = k.Key, Types = k.Value.Select(t => t.Value).ToList() }).ToList();
        }

        private static XmlWriter GetXmlWriter()
        {
            var saveFileDialog = new SaveFileDialog() {
                Filter = "XML files (*.xml)|*.xml",
                FilterIndex = 1,
                RestoreDirectory = true
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                return new XmlTextWriter(saveFileDialog.OpenFile(), null)
                {
                    Formatting = Formatting.Indented
                }; 
            }
            return null;
        }
    }
}