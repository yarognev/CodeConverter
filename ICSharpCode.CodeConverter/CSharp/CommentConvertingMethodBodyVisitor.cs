using System;
using System.Linq;
using System.Threading.Tasks;
using ICSharpCode.CodeConverter.Shared;
using ICSharpCode.CodeConverter.Util;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using CS = Microsoft.CodeAnalysis.CSharp;
using CSSyntax = Microsoft.CodeAnalysis.CSharp.Syntax;
using SyntaxFactory = Microsoft.CodeAnalysis.CSharp.SyntaxFactory;
using SyntaxNodeExtensions = ICSharpCode.CodeConverter.Util.SyntaxNodeExtensions;

namespace ICSharpCode.CodeConverter.CSharp
{
    public class CommentConvertingMethodBodyVisitor : VisualBasicSyntaxVisitor<Task<SyntaxList<CSSyntax.StatementSyntax>>>
    {
        private readonly VisualBasicSyntaxVisitor<Task<SyntaxList<CSSyntax.StatementSyntax>>> _wrappedVisitor;
        private readonly TriviaConverter _triviaConverter;

        public CommentConvertingMethodBodyVisitor(VisualBasicSyntaxVisitor<Task<SyntaxList<CSSyntax.StatementSyntax>>> wrappedVisitor, TriviaConverter triviaConverter)
        {
            this._wrappedVisitor = wrappedVisitor;
            this._triviaConverter = triviaConverter;
        }

        public override async Task<SyntaxList<CSSyntax.StatementSyntax>> DefaultVisit(SyntaxNode node)
        {
            try {
                return await ConvertWithTrivia(node);
            } catch (Exception e) {
                var withTrailingErrorComment = SyntaxFactory.EmptyStatement()
                    .WithCsTrailingErrorComment<CSSyntax.StatementSyntax>((VisualBasicSyntaxNode) node, e);
                return SyntaxFactory.SingletonList(withTrailingErrorComment);
            }
        }

        private async Task<SyntaxList<CSSyntax.StatementSyntax>> ConvertWithTrivia(SyntaxNode node)
        {
            var convertedNodes = await _wrappedVisitor.Visit(node);
            if (!convertedNodes.Any()) return convertedNodes;
            // Port trivia to the last statement in the list
            var lastWithConvertedTrivia = _triviaConverter.PortConvertedTrivia(node, convertedNodes.LastOrDefault());
            return convertedNodes.Replace(convertedNodes.LastOrDefault(), lastWithConvertedTrivia);
        }

        public override async Task<SyntaxList<CSSyntax.StatementSyntax>> VisitTryBlock(TryBlockSyntax node)
        {
            var cSharpSyntaxNodes = await _wrappedVisitor.Visit(node);
            var tryStatementCs = (CSSyntax.TryStatementSyntax)cSharpSyntaxNodes.Single();
            var tryTokenCs = tryStatementCs.TryKeyword;
            var tryStatementWithTryTrivia = tryStatementCs.ReplaceToken(tryTokenCs, tryTokenCs.WithConvertedTriviaFrom(node.TryStatement));
            var tryStatementWithAllTrivia = _triviaConverter.PortConvertedTrivia(node, tryStatementWithTryTrivia);
            return cSharpSyntaxNodes.Replace(tryStatementCs, tryStatementWithAllTrivia);
        }

        public override async Task<SyntaxList<CSSyntax.StatementSyntax>> VisitMultiLineIfBlock(MultiLineIfBlockSyntax node)
        {
            return await CopyConvertedTrivia<CSSyntax.IfStatementSyntax>(node, 
                    (node.IfStatement.ThenKeyword, cs => cs.CloseParenToken),
                    (node.EndIfStatement.EndKeyword, cs => cs.GetLastToken())
            );
        }

        public override async Task<SyntaxList<CSSyntax.StatementSyntax>> VisitElseIfBlock(ElseIfBlockSyntax node)
        {
            return await CopyConvertedTrivia<CSSyntax.IfStatementSyntax>(node,
                (node.ElseIfStatement.ThenKeyword, cs => cs.CloseParenToken)
            );
        }

        public async Task<CSSyntax.ElseClauseSyntax> WithTrivia(ElseBlockSyntax node, CSSyntax.ElseClauseSyntax csNode)
        {
            return csNode.ReplaceToken(csNode.ElseKeyword, csNode.ElseKeyword.WithConvertedTrailingTriviaFrom(node.ElseStatement.ElseKeyword));
        }

        private async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> CopyConvertedTrivia<TCs>(SyntaxNode node, params (SyntaxToken vbToken, Func<TCs, SyntaxToken> getCsToken)[] replacements) where TCs : CSSyntax.StatementSyntax
        {
            var csStatements = await _wrappedVisitor.Visit(node);
            var csStatement = csStatements.Single() as TCs;
            var updatedCsStatement = csStatement;
            foreach (var (vbToken, getCsToken) in replacements) {
                var csToken = getCsToken(updatedCsStatement);
                var updatedCsToken = csToken.WithConvertedTrailingTriviaFrom(vbToken);
                updatedCsStatement = updatedCsStatement.ReplaceToken(csToken, updatedCsToken);
            }
            return csStatements.Replace(csStatement, updatedCsStatement);
        }
    }
}