using System;
using System.Collections.Generic;
using System.Text;
using ICSharpCode.CodeConverter.Util;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using VBSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax;
using VBasic = Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.CSharp;
using System.Linq;
using System.Threading.Tasks;

namespace ICSharpCode.CodeConverter.CSharp
{
    public class ByRefParameterVisitor : VBasic.VisualBasicSyntaxVisitor<Task<SyntaxList<StatementSyntax>>>
    {
        private readonly VBasic.VisualBasicSyntaxVisitor<Task<SyntaxList<StatementSyntax>>> _wrappedVisitor;
        private readonly AdditionalLocals _additionalLocals;
        private readonly SemanticModel _semanticModel;
        private readonly HashSet<string> _generatedNames;

        public ByRefParameterVisitor(VBasic.VisualBasicSyntaxVisitor<Task<SyntaxList<StatementSyntax>>> wrappedVisitor, AdditionalLocals additionalLocals,
            SemanticModel semanticModel, HashSet<string> generatedNames)
        {
            _wrappedVisitor = wrappedVisitor;
            _additionalLocals = additionalLocals;
            _semanticModel = semanticModel;
            _generatedNames = generatedNames;
        }

        public override Task<SyntaxList<StatementSyntax>> DefaultVisit(SyntaxNode node)
        {
            if (node is VBasic.VisualBasicSyntaxNode vbsn) return AddLocalVariables(vbsn);

            throw new NotImplementedException($"Conversion for {VBasic.VisualBasicExtensions.Kind(node)} not implemented, please report this issue")
                .WithNodeInformation(node);
        }

        private async Task<SyntaxList<StatementSyntax>> AddLocalVariables(VBasic.VisualBasicSyntaxNode node)
        {
            _additionalLocals.PushScope();
            IEnumerable<SyntaxNode> csNodes;
            List<StatementSyntax> additionalDeclarations;
            try {
                (csNodes, additionalDeclarations) = await CreateLocals(node);
            } finally {
                _additionalLocals.PopScope();
            }

            return SyntaxFactory.List(additionalDeclarations.Concat(csNodes));
        }

        private async Task<(IEnumerable<SyntaxNode> csNodes, List<StatementSyntax> additionalDeclarations)> CreateLocals(VBasic.VisualBasicSyntaxNode node)
        {
            IEnumerable<SyntaxNode> csNodes = await _wrappedVisitor.Visit(node);

            var additionalDeclarations = new List<StatementSyntax>();

            if (_additionalLocals.Count() > 0)
            {
                var newNames = new Dictionary<string, string>();
                csNodes = csNodes.Select(csNode => csNode.ReplaceNodes(csNode.GetAnnotatedNodes(AdditionalLocals.Annotation),
                    (an, _) =>
                    {
                        var id = ((IdentifierNameSyntax) an).Identifier.ValueText;
                        newNames[id] = NameGenerator.GetUniqueVariableNameInScope(_semanticModel, _generatedNames, node,
                            _additionalLocals[id].Prefix);
                        return SyntaxFactory.IdentifierName(newNames[id]);
                    })).ToList();

                foreach (var additionalLocal in _additionalLocals)
                {
                    var decl = CommonConversions.CreateVariableDeclarationAndAssignment(newNames[additionalLocal.Key],
                        additionalLocal.Value.Initializer, additionalLocal.Value.Type);
                    additionalDeclarations.Add(SyntaxFactory.LocalDeclarationStatement(decl));
                }
            }

            return (csNodes, additionalDeclarations);
        }
    }
}
