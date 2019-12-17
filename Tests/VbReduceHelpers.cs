using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Simplification;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using Xunit;

namespace CodeConverter.Tests
{
    public class VbReduceHelpers
    {
        public static async Task<string> ReduceVbAsync(string validInputVb)
        {
            using (var workspace = new AdhocWorkspace()) {
                Document docInProject = VbProjectHelpers.GetVbProjectWithDocument(validInputVb, workspace);
                await AssertNoCompileWarnings(docInProject);
                var withExpandedRoot = await ReduceVbInternal(docInProject);
                await AssertNoCompileWarnings(withExpandedRoot);
                return (await withExpandedRoot.GetSyntaxRootAsync()).ToFullString();
            }
        }

        private static async Task AssertNoCompileWarnings(Document docInProject)
        {
            var compilation = await docInProject.Project.GetCompilationAsync();
            Assert.Empty(compilation.GetDiagnostics().Where(d => d.Severity >= DiagnosticSeverity.Warning));
        }

        private static async Task<Document> ReduceVbInternal(Document convertedDocument)
        {
            var originalRoot = await convertedDocument.GetSyntaxRootAsync();

            var toSimplify = originalRoot
                .DescendantNodes(n => !(n is ExpressionSyntax));
            var newRoot = originalRoot.ReplaceNodes(toSimplify, (orig, rewritten) =>
                rewritten.WithAdditionalAnnotations(Simplifier.Annotation)
            );

            var document = await WithReducedRootAsync(convertedDocument,
                newRoot.WithAdditionalAnnotations(Simplifier.Annotation));
            return document;
        }

        private static async Task<Document> WithReducedRootAsync(Document doc, SyntaxNode syntaxRoot = null)
        {
            var root = syntaxRoot ?? await doc.GetSyntaxRootAsync();
            var withSyntaxRoot = doc.WithSyntaxRoot(root);
            try {
                return await Simplifier.ReduceAsync(withSyntaxRoot);
            } catch {
                return doc;
            }
        }

    }
}