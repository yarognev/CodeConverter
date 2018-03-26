using System;
using System.Linq;
using ICSharpCode.CodeConverter.Shared;
using ICSharpCode.CodeConverter.Util;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using VbSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax;
using CsSyntax = Microsoft.CodeAnalysis.CSharp.Syntax;
using SyntaxFactory = Microsoft.CodeAnalysis.CSharp.SyntaxFactory;
using SyntaxKind = Microsoft.CodeAnalysis.CSharp.SyntaxKind;

namespace ICSharpCode.CodeConverter.CSharp
{
    public class CommentConvertingNodesVisitor : VisualBasicSyntaxVisitor<CSharpSyntaxNode>
    {
        public TriviaConverter TriviaConverter { get; }
        private readonly VisualBasicSyntaxVisitor<CSharpSyntaxNode> _wrappedVisitor;

        public CommentConvertingNodesVisitor(VisualBasicSyntaxVisitor<CSharpSyntaxNode> wrappedVisitor)
        {
            TriviaConverter = new TriviaConverter();
            this._wrappedVisitor = wrappedVisitor;
        }
        public override CSharpSyntaxNode DefaultVisit(SyntaxNode node)
        {
            var convertedNode = TriviaConverter.PortConvertedTrivia(node, _wrappedVisitor.Visit(node))
                .WithPrependedLeadingTrivia(ConvertStructuredTrivia(node.GetLeadingTrivia()))
                .WithAppendedTrailingTrivia(ConvertStructuredTrivia(node.GetTrailingTrivia()));
            return convertedNode;
        }

        private SyntaxTrivia[] ConvertStructuredTrivia(SyntaxTriviaList triviaList)
        {
            return triviaList.Where(t => t.HasStructure)
                .Select(ConvertStructuredTrivia).ToArray();
        }

        private SyntaxTrivia ConvertStructuredTrivia(SyntaxTrivia t)
        {
            var convertedStructure = (CsSyntax.StructuredTriviaSyntax) Visit(t.GetStructure());
            return SyntaxFactory.Trivia(convertedStructure);
        }

        public override CSharpSyntaxNode VisitIfDirectiveTrivia(VbSyntax.IfDirectiveTriviaSyntax node)
        {
            return SyntaxFactory.IfDirectiveTrivia((CsSyntax.ExpressionSyntax) node.Condition.Accept(this), true, true, true);
                    
        }

        public override CSharpSyntaxNode VisitElseDirectiveTrivia(VbSyntax.ElseDirectiveTriviaSyntax node)
        {
            return SyntaxFactory.ElseDirectiveTrivia(true, true);
        }

        public override CSharpSyntaxNode VisitEndIfDirectiveTrivia(VbSyntax.EndIfDirectiveTriviaSyntax node)
        {
            return SyntaxFactory.EndIfDirectiveTrivia(true);
        }

        public override CSharpSyntaxNode VisitModuleBlock(VbSyntax.ModuleBlockSyntax node)
        {
            return WithPortedTrivia<VbSyntax.TypeBlockSyntax, CsSyntax.BaseTypeDeclarationSyntax>(node, WithTypeBlockTrivia);
        }

        public override CSharpSyntaxNode VisitStructureBlock(VbSyntax.StructureBlockSyntax node)
        {
            return WithPortedTrivia<VbSyntax.TypeBlockSyntax, CsSyntax.BaseTypeDeclarationSyntax>(node, WithTypeBlockTrivia);
        }

        public override CSharpSyntaxNode VisitInterfaceBlock(VbSyntax.InterfaceBlockSyntax node)
        {
            return WithPortedTrivia<VbSyntax.TypeBlockSyntax, CsSyntax.BaseTypeDeclarationSyntax>(node, WithTypeBlockTrivia);
        }

        public override CSharpSyntaxNode VisitClassBlock(VbSyntax.ClassBlockSyntax node)
        {
            return WithPortedTrivia<VbSyntax.TypeBlockSyntax, CsSyntax.BaseTypeDeclarationSyntax>(node, WithTypeBlockTrivia);
        }

        public override CSharpSyntaxNode VisitCompilationUnit(VbSyntax.CompilationUnitSyntax node)
        {
            var cSharpSyntaxNode = (CsSyntax.CompilationUnitSyntax) base.VisitCompilationUnit(node);
            cSharpSyntaxNode = cSharpSyntaxNode.WithEndOfFileToken(
                cSharpSyntaxNode.EndOfFileToken.WithConvertedLeadingTriviaFrom(node.EndOfFileToken));

            return TriviaConverter.IsAllTriviaConverted() 
                ? cSharpSyntaxNode 
                : cSharpSyntaxNode.WithAppendedTrailingTrivia(SyntaxFactory.Comment("/* Some trivia (e.g. comments) could not be converted */"));
        }

        private TDest WithPortedTrivia<TSource, TDest>(SyntaxNode node, Func<TSource, TDest, TDest> portExtraTrivia) where TSource : SyntaxNode where TDest : CSharpSyntaxNode
        {
            var cSharpSyntaxNode = portExtraTrivia((TSource)node, (TDest)_wrappedVisitor.Visit(node));
            return TriviaConverter.PortConvertedTrivia(node, cSharpSyntaxNode);
        }

        private CsSyntax.BaseTypeDeclarationSyntax WithTypeBlockTrivia(VbSyntax.TypeBlockSyntax sourceNode, CsSyntax.BaseTypeDeclarationSyntax destNode)
        {
            var beforeOpenBrace = destNode.OpenBraceToken.GetPreviousToken();
            var withAnnotation = TriviaConverter.WithDelegateToParentAnnotation(sourceNode.BlockStatement, beforeOpenBrace);
            withAnnotation = TriviaConverter.WithDelegateToParentAnnotation(sourceNode.Inherits, withAnnotation);
            withAnnotation = TriviaConverter.WithDelegateToParentAnnotation(sourceNode.Implements, withAnnotation);
            return destNode.ReplaceToken(beforeOpenBrace, withAnnotation);
        }
    }
}