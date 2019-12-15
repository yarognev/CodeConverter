using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Simplification;
using Microsoft.CodeAnalysis.VisualBasic;

namespace CodeConverter.Tests
{
    public class VbExpandHelpers
    {
        public static async Task<string> ExpandVb(string validInputVb)
        {
            using (var workspace = new AdhocWorkspace()) {
                Document docInProject = VbProjectHelpers.GetVbProjectWithDocument(validInputVb, workspace);
                var withExpandedRoot = await ExpandVbAsync(docInProject, ShouldExpandVbNode);
                return (await withExpandedRoot.GetSyntaxRootAsync()).ToFullString();
            }
        }

        private static async Task<Document> ExpandVbAsync(Document document, Func<SemanticModel, SyntaxNode, bool> shouldExpand)
        {
            var semanticModel = await document.GetSemanticModelAsync();
            var workspace = document.Project.Solution.Workspace;
            var root = (VisualBasicSyntaxNode)await document.GetSyntaxRootAsync();
            try {
                var newRoot = root.ReplaceNodes(root.DescendantNodes(n => !shouldExpand(semanticModel, n)).Where(n => shouldExpand(semanticModel, n)),
                    (node, rewrittenNode) => TryExpandNode(node, semanticModel, workspace)
                );
                return document.WithSyntaxRoot(newRoot);
            } catch (Exception) {
                return document.WithSyntaxRoot(root);
            }
        }

        private static SyntaxNode TryExpandNode(SyntaxNode node, SemanticModel semanticModel, Workspace workspace)
        {
            try {
                return Simplifier.Expand(node, semanticModel, workspace);
            } catch (Exception) {
                return node;
            }
        }

        private static bool ShouldExpandVbNode(SemanticModel semanticModel, SyntaxNode node)
        {
            return node is Microsoft.CodeAnalysis.VisualBasic.Syntax.NameSyntax || node is Microsoft.CodeAnalysis.VisualBasic.Syntax.InvocationExpressionSyntax && !IsReducedTypeParameterMethod(semanticModel.GetSymbolInfo(node).Symbol);
        }

        private static bool IsReducedTypeParameterMethod(ISymbol symbol)
        {
            return symbol is IMethodSymbol ms && ms.ReducedFrom?.TypeParameters.Count() > ms.TypeParameters.Count();
        }
    }
}