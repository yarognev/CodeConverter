using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.VisualBasic;

namespace CodeConverter.Tests
{
    public class VbProjectHelpers
    {
        private static readonly Type[] TypesToLoadAssembliesFor = {
            typeof(Encoding),
            typeof(Enumerable),
            typeof(BrowsableAttribute),
            typeof(DynamicObject),
            typeof(DataRow),
            typeof(HttpClient),
            typeof(HttpUtility),
            typeof(XmlElement),
            typeof(XElement),
            typeof(Microsoft.VisualBasic.Constants)
        };

        public static IReadOnlyCollection<PortableExecutableReference> NetStandard2 => GetRefs(TypesToLoadAssembliesFor).ToArray();

        private static IEnumerable<PortableExecutableReference> GetRefs(IReadOnlyCollection<Type> types)
        {
            return types.Select(type => MetadataReference.CreateFromFile(type.Assembly.Location));
        }


        public static Document GetVbProjectWithDocument(string validInputVb, AdhocWorkspace workspace)
        {
            var tree = SyntaxFactory.ParseSyntaxTree(validInputVb, encoding: Encoding.UTF8);
            var compilationOptions = new VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
                .WithGlobalImports(GlobalImport.Parse(
                    "System",
                    "System.Collections",
                    "System.Collections.Generic",
                    "System.Diagnostics",
                    "System.Globalization",
                    "System.IO",
                    "System.Linq",
                    "System.Reflection",
                    "System.Runtime.CompilerServices",
                    "System.Runtime.InteropServices",
                    "System.Security",
                    "System.Text",
                    "System.Threading.Tasks",
                    "Microsoft.VisualBasic"))
                .WithOptionExplicit(true)
                .WithOptionCompareText(false)
                .WithOptionStrict(OptionStrict.Off)
                .WithOptionInfer(true);

            var singleDocumentAssemblyName = "ProjectToBeConverted";
            ProjectId projectId = ProjectId.CreateNewId();
            var solution = workspace.CurrentSolution.AddProject(projectId, singleDocumentAssemblyName,
                singleDocumentAssemblyName, LanguageNames.VisualBasic);

            var project = solution.GetProject(projectId)
                .WithCompilationOptions(compilationOptions)
                .WithParseOptions(VisualBasicParseOptions.Default)
                .WithMetadataReferences(VbProjectHelpers.NetStandard2);
            var projWithDoc = project.AddDocument("CodeToConvert", tree.GetRoot(),
                filePath: Path.Combine(Directory.GetCurrentDirectory(), "TempCodeToConvert.txt"));
            return projWithDoc;
        }
    }
}