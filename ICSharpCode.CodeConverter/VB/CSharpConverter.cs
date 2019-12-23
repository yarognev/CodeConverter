using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Threading.Tasks;
using ICSharpCode.CodeConverter.Util;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Editing;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Rename;
using Microsoft.CodeAnalysis.VisualBasic;
using CS = Microsoft.CodeAnalysis.CSharp;
using CSS = Microsoft.CodeAnalysis.CSharp.Syntax;

namespace ICSharpCode.CodeConverter.VB
{
    internal class CSharpConverter
    {
        public static async Task<SyntaxNode> ConvertCompilationTree(Document document,
            VisualBasicCompilation vbViewOfCsSymbols, Project vbReferenceProject)
        {
            var solution = await RenameClashingSymbols(document);
            document = solution.GetDocument(document.Id);

            var compilation = await document.Project.GetCompilationAsync();
            var tree = await document.GetSyntaxTreeAsync();
            var semanticModel = compilation.GetSemanticModel(tree, true);
            var root = await document.GetSyntaxRootAsync() as CS.CSharpSyntaxNode ??
                       throw new InvalidOperationException(NullRootError(document));

            var vbSyntaxGenerator = SyntaxGenerator.GetGenerator(vbReferenceProject);
            var visualBasicSyntaxVisitor = new NodesVisitor(document, (CS.CSharpCompilation)compilation, semanticModel, vbViewOfCsSymbols, vbSyntaxGenerator);
            return root.Accept(visualBasicSyntaxVisitor.TriviaConvertingVisitor);
        }

        private static string NullRootError(Document document)
        {
            var initial = document.Project.Language != LanguageNames.CSharp
                ? "Document cannot be converted because it's not within a C# project."
                : "Could not find valid C# within document.";
            return initial + " For best results, convert a c# document from within a C# project which compiles successfully.";
        }

        private static async Task<Solution> RenameClashingSymbols(Document document)
        {
            var compilation = await document.Project.GetCompilationAsync();
            var tree = await document.GetSyntaxTreeAsync();
            var semanticModel = compilation.GetSemanticModel(tree, true);
            var root = (CS.CSharpSyntaxNode) await document.GetSyntaxRootAsync();

            Solution solution = await RenameClashingSymbols(document, semanticModel, root);
            return solution;
        }

        private static async Task<Solution> RenameClashingSymbols(Document document, SemanticModel semanticModel, CS.CSharpSyntaxNode root)
        {
            var toRename = root
                .DescendantNodes(n => !(n is CSS.ExpressionSyntax)).Where(n => n is CSS.BaseTypeDeclarationSyntax)
                .Select(id =>
                {
                    var idSymbol = (ITypeSymbol) semanticModel.GetDeclaredSymbol(id);
                    var allSymbols = GetCsSymbolsDeclaredByType(semanticModel, idSymbol).Select(s => s.OriginalDefinition).ToList();
                    var potentialClashes = idSymbol != null ? GetPotentialClashes(semanticModel, id, idSymbol) : null;
                    return (idSymbol, potentialClashes);
                }).ToLookup(x => x.idSymbol, x => x.potentialClashes)
                .Select(x => (Declaration: x.Key, Clashes: x.SelectMany(y => y)))
                .Where(x => x.Declaration != null && x.Clashes.Any())
                .ToArray();
            var solution = document.Project.Solution;
            //TODO Only rename where qualification can't solve the issue
            //TODO Rename all but one symbol
            //TODO Prefer renaming locals first
            foreach (var (declaration, clashes) in toRename) {
                var compilation = await solution.GetProject(document.Project.Id).GetCompilationAsync();
                ISymbol currentDeclaration = SymbolFinder.FindSimilarSymbols(declaration, compilation).First();
                var model = compilation.GetSemanticModel(currentDeclaration.DeclaringSyntaxReferences.First().SyntaxTree);
                string declarationName = await GenerateUniqueCaseInsensitiveName(model, currentDeclaration);
                solution = await Renamer.RenameSymbolAsync(solution, currentDeclaration, declarationName,
                    await document.GetOptionsAsync());
            }

            return solution;
        }

        public static IEnumerable<ISymbol> GetCsSymbolsDeclaredByType(SemanticModel semanticModel, ITypeSymbol typeSymbol)
        {
            var members = typeSymbol.GetMembers();
            if (!members.Any())
                return DeclarationWhereNotNull(typeSymbol, (CSS.BaseTypeDeclarationSyntax r) => r.OpenBraceToken.Span.End)
                    .SelectMany(r => semanticModel.LookupSymbols(r, typeSymbol));
            return members.SelectMany(x => {
                switch (x) {
                    case IMethodSymbol methodSymbol:
                        return GetCsSymbolsDeclaredByMethod(semanticModel, methodSymbol);
                    case IPropertySymbol propertySymbol:
                        return GetCsSymbolsDeclaredByProperty(semanticModel, propertySymbol);
                    case ITypeSymbol nestedTypeSymbol:
                        return GetCsSymbolsDeclaredByType(semanticModel, nestedTypeSymbol);
                    case IEventSymbol eventSymbol:
                        return GetCsSymbolsDeclaredByEvent(semanticModel, eventSymbol);
                    case IFieldSymbol fieldSymbol:
                        return GetCsSymbolsDeclaredByField(semanticModel, fieldSymbol);
                    default:
                        return new ISymbol[0];
                }
            });
        }

        public static IEnumerable<ISymbol> GetCsSymbolsDeclaredByMethod(SemanticModel semanticModel, IMethodSymbol methodSymbol)
        {
            if (methodSymbol == null) return new ISymbol[0];
            var symbols =
                methodSymbol.TypeParameters
                    .Concat<ISymbol>(methodSymbol.Parameters);
            var bodies = DeclarationWhereNotNull(methodSymbol, (CSS.BaseMethodDeclarationSyntax b) => (CS.CSharpSyntaxNode)b.ExpressionBody ?? b.Body);
            foreach(var body in bodies) {
                symbols = symbols.Concat(semanticModel.LookupSymbols(body.SpanStart, methodSymbol.ContainingType));
                var descendantNodes = body.DescendantNodes().OfType<CSS.BlockSyntax>();
                symbols = symbols.Concat(descendantNodes.SelectMany(d => semanticModel.LookupSymbols(d.SpanStart, methodSymbol.ContainingType)));
            }

            return symbols;
        }

        private static IEnumerable<TResult> DeclarationWhereNotNull<TNode, TResult>(ISymbol symbol, Func<TNode, TResult> selectWhereNotNull)
        {
            return symbol.DeclaringSyntaxReferences.Select(d => d.GetSyntax()).OfType<TNode>().Select(selectWhereNotNull).Where(x => x != null);
        }

        private static IEnumerable<TResult> DeclarationWhereManyNotNull<TNode, TResult>(ISymbol symbol, Func<TNode, IEnumerable<TResult>> selectManyWhereNotNull)
        {
            return symbol.DeclaringSyntaxReferences.Select(d => d.GetSyntax()).OfType<TNode>().SelectMany(selectManyWhereNotNull).Where(x => x != null);
        }

        public static IEnumerable<ISymbol> GetCsSymbolsDeclaredByProperty(SemanticModel semanticModel, IPropertySymbol propertySymbol)
        {
            return GetCsSymbolsDeclaredByMethod(semanticModel, propertySymbol.GetMethod)
                .Concat(GetCsSymbolsDeclaredByMethod(semanticModel, propertySymbol.SetMethod));
        }

        public static IEnumerable<ISymbol> GetCsSymbolsDeclaredByField(SemanticModel semanticModel, IFieldSymbol fieldSymbol)
        {
            return DeclarationWhereManyNotNull(fieldSymbol,
                (CSS.BaseFieldDeclarationSyntax f) => f.Declaration.Variables.Select(v => v.Initializer?.Value))
                .SelectMany(i => semanticModel.LookupSymbols(i.SpanStart, fieldSymbol.ContainingType));
        }

        public static IEnumerable<ISymbol> GetCsSymbolsDeclaredByEvent(SemanticModel semanticModel, IEventSymbol propertySymbol)
        {
            return GetCsSymbolsDeclaredByMethod(semanticModel, propertySymbol.AddMethod)
                .Concat(GetCsSymbolsDeclaredByMethod(semanticModel, propertySymbol.RemoveMethod));
        }

        private static async Task<string> GenerateUniqueCaseInsensitiveName(SemanticModel model, ISymbol declaration)
        {
            string prefix = declaration.Kind.ToString().ToLowerInvariant()[0].ToString();
            string baseName = prefix + declaration.Name.Substring(0, 1).ToUpperInvariant() + declaration.Name.Substring(1);
            HashSet<string> generatedNames = new HashSet<string>();
            //TODO Make looking for clashes case insensitive as above
            return NameGenerator.GetUniqueVariableNameInScope(model, generatedNames, (CS.CSharpSyntaxNode) await declaration.DeclaringSyntaxReferences.First().GetSyntaxAsync(), baseName);
        }

        private static bool IsSymbolDeclaration(SyntaxNode n)
        {
            if (IsLocalSymbolDeclaration(n)) return true;

            switch (n) {
                case CSS.BaseTypeDeclarationSyntax _:
                case CSS.BaseFieldDeclarationSyntax _:
                case CSS.BaseMethodDeclarationSyntax _:
                case CSS.BasePropertyDeclarationSyntax _:
                case CSS.NamespaceDeclarationSyntax _:
                    return true;
            }

            return false;
        }

        private static bool IsLocalSymbolDeclaration(SyntaxNode n)
        {
            switch (n) {
                case CSS.ArgumentSyntax _:
                case CSS.AnonymousObjectMemberDeclaratorSyntax _:
                case CSS.AnonymousObjectCreationExpressionSyntax _:
                case CSS.JoinIntoClauseSyntax _:
                case CSS.QueryContinuationSyntax _:
                case CSS.VariableDeclaratorSyntax _:
                case CSS.LabeledStatementSyntax _:
                case CSS.ForEachStatementSyntax _:
                case CSS.SwitchLabelSyntax _:
                case CSS.CatchDeclarationSyntax _:
                case CSS.LocalFunctionStatementSyntax _:
                case CSS.NamespaceDeclarationSyntax _:
                case CSS.UsingDirectiveSyntax _:
                case CSS.ParameterSyntax _:
                case CSS.TypeParameterSyntax _:
                case CSS.TupleElementSyntax _:
                case CSS.TupleExpressionSyntax _:
                case CSS.SingleVariableDesignationSyntax _:
                    return true;
            }

            return false;
        }

        private static IEnumerable<ISymbol> GetPotentialClashes(SemanticModel model, SyntaxNode id, ISymbol idSymbol)
        {
            return model.LookupSymbols(id.SpanStart).Where(s =>
                s.Name.Equals(idSymbol.Name, StringComparison.OrdinalIgnoreCase) && !s.Equals(idSymbol));
        }
    }
}
