using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ICSharpCode.CodeConverter.Shared;
using ICSharpCode.CodeConverter.Util;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Operations;
using SyntaxFactory = Microsoft.CodeAnalysis.CSharp.SyntaxFactory;
using VBasic = Microsoft.CodeAnalysis.VisualBasic;
using VBSyntax = Microsoft.CodeAnalysis.VisualBasic.Syntax;
using static ICSharpCode.CodeConverter.CSharp.SyntaxKindExtensions;
using Microsoft.CodeAnalysis.Text;

namespace ICSharpCode.CodeConverter.CSharp
{
    /// <summary>
    /// Executable statements - which includes executable blocks such as if statements
    /// Maintains state relevant to the called method-like object. A fresh one must be used for each method, and the same one must be reused for statements in the same method
    /// </summary>
    internal class MethodBodyExecutableStatementVisitor : VBasic.VisualBasicSyntaxVisitor<Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>>>
    {
        private readonly VBasic.VisualBasicSyntaxNode _methodNode;
        private readonly SemanticModel _semanticModel;
        private readonly CommentConvertingVisitorWrapper<CSharpSyntaxNode> _expressionVisitor;
        private readonly Stack<Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax> _withBlockLhs;
        private readonly HashSet<string> _extraUsingDirectives;
        private readonly MethodsWithHandles _methodsWithHandles;
        private readonly HashSet<string> _generatedNames = new HashSet<string>();
        private INamedTypeSymbol _vbBooleanTypeSymbol;

        public bool IsIterator { get; set; }
        public Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax ReturnVariable { get; set; }
        public bool HasReturnVariable => ReturnVariable != null;
        public CommentConvertingMethodBodyVisitor CommentConvertingVisitor { get; }

        private CommonConversions CommonConversions { get; }

        public MethodBodyExecutableStatementVisitor(VBasic.VisualBasicSyntaxNode methodNode, SemanticModel semanticModel,
            CommentConvertingVisitorWrapper<CSharpSyntaxNode> expressionVisitor, CommonConversions commonConversions,
            Stack<Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax> withBlockLhs, HashSet<string> extraUsingDirectives,
            AdditionalLocals additionalLocals, MethodsWithHandles methodsWithHandles,
            TriviaConverter triviaConverter)
        {
            _methodNode = methodNode;
            _semanticModel = semanticModel;
            _expressionVisitor = expressionVisitor;
            CommonConversions = commonConversions;
            _withBlockLhs = withBlockLhs;
            _extraUsingDirectives = extraUsingDirectives;
            _methodsWithHandles = methodsWithHandles;
            var byRefParameterVisitor = new ByRefParameterVisitor(this, additionalLocals, semanticModel, _generatedNames);
            CommentConvertingVisitor = new CommentConvertingMethodBodyVisitor(byRefParameterVisitor, triviaConverter);
            _vbBooleanTypeSymbol = _semanticModel.Compilation.GetTypeByMetadataName("System.Boolean");
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> DefaultVisit(SyntaxNode node)
        {
            throw new NotImplementedException($"Conversion for {VBasic.VisualBasicExtensions.Kind(node)} not implemented, please report this issue")
                .WithNodeInformation(node);
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitStopOrEndStatement(VBSyntax.StopOrEndStatementSyntax node)
        {
            return SingleStatement(SyntaxFactory.ParseStatement(ConvertStopOrEndToCSharpStatementText(node)));
        }

        private string ConvertStopOrEndToCSharpStatementText(VBSyntax.StopOrEndStatementSyntax node)
        {
            switch (VBasic.VisualBasicExtensions.Kind(node.StopOrEndKeyword)) {
                case VBasic.SyntaxKind.StopKeyword:
                    _extraUsingDirectives.Add("System.Diagnostics");
                    return "Debugger.Break();";
                case VBasic.SyntaxKind.EndKeyword:
                    _extraUsingDirectives.Add("System");
                    return "Environment.Exit(0);";
                default:
                    throw new NotImplementedException(node.StopOrEndKeyword.Kind() + " not implemented!");
            }
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitLocalDeclarationStatement(VBSyntax.LocalDeclarationStatementSyntax node)
        {
            var modifiers = CommonConversions.ConvertModifiers(node.Declarators[0].Names[0], node.Modifiers, TokenContext.Local);
            var isConst = modifiers.Any(a => a.Kind() == SyntaxKind.ConstKeyword);

            var declarations = new List<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>();

            foreach (var declarator in node.Declarators) {
                var splitVariableDeclarations = await CommonConversions.SplitVariableDeclarations(declarator, preferExplicitType: isConst);
                var localDeclarationStatementSyntaxs = splitVariableDeclarations.Variables.Select(decl => SyntaxFactory.LocalDeclarationStatement(modifiers, decl));
                declarations.AddRange(localDeclarationStatementSyntaxs);
                var localFunctions = splitVariableDeclarations.Methods.Cast<LocalFunctionStatementSyntax>();
                declarations.AddRange(localFunctions);
            }

            return SyntaxFactory.List(declarations);
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitAddRemoveHandlerStatement(VBSyntax.AddRemoveHandlerStatementSyntax node)
        {
            var syntaxKind = ConvertAddRemoveHandlerToCSharpSyntaxKind(node);
            return SingleStatement(SyntaxFactory.AssignmentExpression(syntaxKind,
                (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.EventExpression.AcceptAsync(_expressionVisitor),
                (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.DelegateExpression.AcceptAsync(_expressionVisitor)));
        }

        private static SyntaxKind ConvertAddRemoveHandlerToCSharpSyntaxKind(VBSyntax.AddRemoveHandlerStatementSyntax node)
        {
            switch (node.Kind()) {
                case VBasic.SyntaxKind.AddHandlerStatement:
                    return SyntaxKind.AddAssignmentExpression;
                case VBasic.SyntaxKind.RemoveHandlerStatement:
                    return SyntaxKind.SubtractAssignmentExpression;
                default:
                    throw new NotImplementedException(node.Kind() + " not implemented!");
            }
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitExpressionStatement(VBSyntax.ExpressionStatementSyntax node)
        {
            if (node.Expression is VBSyntax.InvocationExpressionSyntax invoke && invoke.Expression is VBSyntax.MemberAccessExpressionSyntax access && access.Expression is VBSyntax.MyBaseExpressionSyntax && access.Name.Identifier.ValueText.Equals("Finalize", StringComparison.OrdinalIgnoreCase)) {
                return new SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>();
            }

            return SingleStatement((Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.Expression.AcceptAsync(_expressionVisitor));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitAssignmentStatement(VBSyntax.AssignmentStatementSyntax node)
        {
            var lhs = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.Left.AcceptAsync(_expressionVisitor);
            var lOperation = _semanticModel.GetOperation(node.Left);

            //Already dealt with by call to the same method in VisitInvocationExpression
            var (parameterizedPropertyAccessMethod, _) = await CommonConversions.GetParameterizedPropertyAccessMethod(lOperation);
            if (parameterizedPropertyAccessMethod != null) return SingleStatement(lhs);
            var rhs = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.Right.AcceptAsync(_expressionVisitor);

            if (node.Left is VBSyntax.IdentifierNameSyntax id &&
                _methodNode is VBSyntax.MethodBlockSyntax mb &&
                HasReturnVariable &&
                id.Identifier.ValueText.Equals(mb.SubOrFunctionStatement.Identifier.ValueText, StringComparison.OrdinalIgnoreCase)) {
                lhs = ReturnVariable;
            }

            if (node.IsKind(VBasic.SyntaxKind.ExponentiateAssignmentStatement)) {
                rhs = SyntaxFactory.InvocationExpression(
                    SyntaxFactory.ParseExpression($"{nameof(Math)}.{nameof(Math.Pow)}"),
                    ExpressionSyntaxExtensions.CreateArgList(lhs, rhs));
            }
            var kind = node.Kind().ConvertToken(TokenContext.Local);

            rhs = CommonConversions.TypeConversionAnalyzer.AddExplicitConversion(node.Right, rhs);
            var assignment = SyntaxFactory.AssignmentExpression(kind, lhs, rhs);

            var postAssignment = GetPostAssignmentStatements(node);
            return postAssignment.Insert(0, SyntaxFactory.ExpressionStatement(assignment));
        }

        private SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax> GetPostAssignmentStatements(VBSyntax.AssignmentStatementSyntax node)
        {
            var potentialPropertySymbol = _semanticModel.GetSymbolInfo(node.Left).ExtractBestMatch();
            return _methodsWithHandles.GetPostAssignmentStatements(node, potentialPropertySymbol);
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitEraseStatement(VBSyntax.EraseStatementSyntax node)
        {
            var eraseStatements = await node.Expressions.SelectAsync<VBSyntax.ExpressionSyntax, Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>(async arrayExpression => {
                var lhs = await arrayExpression.AcceptAsync(_expressionVisitor);
                var rhs = SyntaxFactory.LiteralExpression(SyntaxKind.NullLiteralExpression);
                var assignmentExpressionSyntax =
                    SyntaxFactory.AssignmentExpression(SyntaxKind.SimpleAssignmentExpression, (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)lhs,
                        rhs);
                return SyntaxFactory.ExpressionStatement(assignmentExpressionSyntax);
            });
            return SyntaxFactory.List(eraseStatements);
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitReDimStatement(VBSyntax.ReDimStatementSyntax node)
        {
            return SyntaxFactory.List(await node.Clauses.SelectManyAsync(async arrayExpression => (IEnumerable<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>) await ConvertRedimClause(arrayExpression)));
        }

        /// <remarks>
        /// RedimClauseSyntax isn't an executable statement, therefore this isn't a "Visit" method.
        /// Since it returns multiple statements it's easiest for it to be here in the current architecture.
        /// </remarks>
        private async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> ConvertRedimClause(VBSyntax.RedimClauseSyntax node)
        {
            bool preserve = node.Parent is VBSyntax.ReDimStatementSyntax rdss && rdss.PreserveKeyword.IsKind(VBasic.SyntaxKind.PreserveKeyword);

            var csTargetArrayExpression = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.Expression.AcceptAsync(_expressionVisitor);
            var convertedBounds = (await CommonConversions.ConvertArrayBounds(node.ArrayBounds)).Sizes.ToList();

            var newArrayAssignment = CreateNewArrayAssignment(node.Expression, csTargetArrayExpression, convertedBounds, node.SpanStart);
            if (!preserve) return SingleStatement(newArrayAssignment);

            var lastIdentifierText = node.Expression.DescendantNodesAndSelf().OfType<VBSyntax.IdentifierNameSyntax>().Last().Identifier.Text;
            var (oldTargetExpression, stmts) = await GetReusableExpression(node.Expression, "old" + lastIdentifierText.ToPascalCase(), true);
            var arrayCopyIfNotNull = CreateConditionalArrayCopy(node, (Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax) oldTargetExpression, csTargetArrayExpression, convertedBounds);

            return stmts.AddRange(new Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax[] {newArrayAssignment, arrayCopyIfNotNull});
        }

        /// <summary>
        /// Cut down version of Microsoft.VisualBasic.CompilerServices.Utils.CopyArray
        /// </summary>
        private Microsoft.CodeAnalysis.CSharp.Syntax.IfStatementSyntax CreateConditionalArrayCopy(VBasic.VisualBasicSyntaxNode originalVbNode,
            Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax sourceArrayExpression,
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax targetArrayExpression,
            List<Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax> convertedBounds)
        {
            var sourceLength = SyntaxFactory.MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression, sourceArrayExpression, SyntaxFactory.IdentifierName("Length"));
            var arrayCopyStatement = convertedBounds.Count == 1
                ? CreateArrayCopyWithMinOfLengths(sourceArrayExpression, sourceLength, targetArrayExpression, convertedBounds.Single())
                : CreateArrayCopy(originalVbNode, sourceArrayExpression, sourceLength, targetArrayExpression, convertedBounds);

            var oldTargetNotEqualToNull = SyntaxFactory.BinaryExpression(SyntaxKind.NotEqualsExpression, sourceArrayExpression,
                SyntaxFactory.LiteralExpression(SyntaxKind.NullLiteralExpression));
            return SyntaxFactory.IfStatement(oldTargetNotEqualToNull, arrayCopyStatement);
        }

        /// <summary>
        /// Array copy for multiple array dimensions represented by <paramref name="convertedBounds"/>
        /// </summary>
        /// <remarks>
        /// Exception cases will sometimes silently succeed in the converted code,
        ///  but existing VB code relying on the exception thrown from a multidimensional redim preserve on
        ///  different rank arrays is hopefully rare enough that it's worth saving a few lines of code
        /// </remarks>
        private Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax CreateArrayCopy(VBasic.VisualBasicSyntaxNode originalVbNode,
            Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax sourceArrayExpression,
            Microsoft.CodeAnalysis.CSharp.Syntax.MemberAccessExpressionSyntax sourceLength,
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax targetArrayExpression, ICollection convertedBounds)
        {
            var lastSourceLengthArgs = ExpressionSyntaxExtensions.CreateArgList(CommonConversions.Literal(convertedBounds.Count - 1));
            var sourceLastRankLength = SyntaxFactory.InvocationExpression(
                SyntaxFactory.ParseExpression($"{sourceArrayExpression.Identifier}.GetLength"), lastSourceLengthArgs);
            var targetLastRankLength =
                SyntaxFactory.InvocationExpression(SyntaxFactory.ParseExpression($"{targetArrayExpression}.GetLength"),
                    lastSourceLengthArgs);
            var length = SyntaxFactory.InvocationExpression(SyntaxFactory.ParseExpression("Math.Min"), ExpressionSyntaxExtensions.CreateArgList(sourceLastRankLength, targetLastRankLength));

            var loopVariableName = GetUniqueVariableNameInScope(originalVbNode, "i");
            var loopVariableIdentifier = SyntaxFactory.IdentifierName(loopVariableName);
            var sourceStartForThisIteration =
                SyntaxFactory.BinaryExpression(SyntaxKind.MultiplyExpression, loopVariableIdentifier, sourceLastRankLength);
            var targetStartForThisIteration =
                SyntaxFactory.BinaryExpression(SyntaxKind.MultiplyExpression, loopVariableIdentifier, targetLastRankLength);

            var arrayCopy = CreateArrayCopyWithStartingPoints(sourceArrayExpression, sourceStartForThisIteration, targetArrayExpression,
                targetStartForThisIteration, length);

            var sourceArrayCount = SyntaxFactory.BinaryExpression(SyntaxKind.SubtractExpression,
                SyntaxFactory.BinaryExpression(SyntaxKind.DivideExpression, sourceLength, sourceLastRankLength), CommonConversions.Literal(1));

            return CreateForZeroToValueLoop(loopVariableIdentifier, arrayCopy, sourceArrayCount);
        }

        private Microsoft.CodeAnalysis.CSharp.Syntax.ForStatementSyntax CreateForZeroToValueLoop(Microsoft.CodeAnalysis.CSharp.Syntax.SimpleNameSyntax loopVariableIdentifier, Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax loopStatement, Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax inclusiveLoopUpperBound)
        {
            var loopVariableAssignment = CommonConversions.CreateVariableDeclarationAndAssignment(loopVariableIdentifier.Identifier.Text, CommonConversions.Literal(0));
            var lessThanSourceBounds = SyntaxFactory.BinaryExpression(SyntaxKind.LessThanOrEqualExpression,
                loopVariableIdentifier, inclusiveLoopUpperBound);
            var incrementors = SyntaxFactory.SingletonSeparatedList<Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax>(
                SyntaxFactory.PrefixUnaryExpression(SyntaxKind.PreIncrementExpression, loopVariableIdentifier));
            var forStatementSyntax = SyntaxFactory.ForStatement(loopVariableAssignment,
                SyntaxFactory.SeparatedList<Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax>(),
                lessThanSourceBounds, incrementors, loopStatement);
            return forStatementSyntax;
        }

        private static Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionStatementSyntax CreateArrayCopyWithMinOfLengths(
            Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax sourceExpression, Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax sourceLength,
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax targetExpression, Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax targetLength)
        {
            var minLength = SyntaxFactory.InvocationExpression(SyntaxFactory.ParseExpression("Math.Min"), ExpressionSyntaxExtensions.CreateArgList(targetLength, sourceLength));
            var copyArgList = ExpressionSyntaxExtensions.CreateArgList(sourceExpression, targetExpression, minLength);
            var arrayCopy = SyntaxFactory.InvocationExpression(SyntaxFactory.ParseExpression("Array.Copy"), copyArgList);
            return SyntaxFactory.ExpressionStatement(arrayCopy);
        }

        private static Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionStatementSyntax CreateArrayCopyWithStartingPoints(
            Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax sourceExpression, Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax sourceStart,
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax targetExpression, Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax targetStart, Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax length)
        {
            var copyArgList = ExpressionSyntaxExtensions.CreateArgList(sourceExpression, sourceStart, targetExpression, targetStart, length);
            var arrayCopy = SyntaxFactory.InvocationExpression(SyntaxFactory.ParseExpression("Array.Copy"), copyArgList);
            return SyntaxFactory.ExpressionStatement(arrayCopy);
        }

        private Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionStatementSyntax CreateNewArrayAssignment(VBSyntax.ExpressionSyntax vbArrayExpression,
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax csArrayExpression, List<Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax> convertedBounds,
            int nodeSpanStart)
        {
            var convertedType = (IArrayTypeSymbol) _semanticModel.GetTypeInfo(vbArrayExpression).ConvertedType;
            var arrayRankSpecifierSyntax = SyntaxFactory.ArrayRankSpecifier(SyntaxFactory.SeparatedList(convertedBounds));
            var rankSpecifiers = SyntaxFactory.SingletonList(arrayRankSpecifierSyntax);
            while (convertedType.ElementType is IArrayTypeSymbol ats) {
                convertedType = ats;
                rankSpecifiers = rankSpecifiers.Add(SyntaxFactory.ArrayRankSpecifier(SyntaxFactory.SingletonSeparatedList<Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax>(SyntaxFactory.OmittedArraySizeExpression())));
            };
            var typeSyntax = CommonConversions.GetTypeSyntax(convertedType.ElementType);
            var arrayCreation =
                SyntaxFactory.ArrayCreationExpression(SyntaxFactory.ArrayType(typeSyntax, rankSpecifiers));
            var assignmentExpressionSyntax =
                SyntaxFactory.AssignmentExpression(SyntaxKind.SimpleAssignmentExpression, csArrayExpression, arrayCreation);
            var newArrayAssignment = SyntaxFactory.ExpressionStatement(assignmentExpressionSyntax);
            return newArrayAssignment;
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitThrowStatement(VBSyntax.ThrowStatementSyntax node)
        {
            return SingleStatement(SyntaxFactory.ThrowStatement((Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.Expression.AcceptAsync(_expressionVisitor)));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitReturnStatement(VBSyntax.ReturnStatementSyntax node)
        {
            if (IsIterator)
                return SingleStatement(SyntaxFactory.YieldStatement(SyntaxKind.YieldBreakStatement));

            var csExpression = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)await node.Expression.AcceptAsync(_expressionVisitor);
            csExpression = CommonConversions.TypeConversionAnalyzer.AddExplicitConversion(node.Expression, csExpression);
            return SingleStatement(SyntaxFactory.ReturnStatement(csExpression));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitContinueStatement(VBSyntax.ContinueStatementSyntax node)
        {
            return SingleStatement(SyntaxFactory.ContinueStatement());
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitYieldStatement(VBSyntax.YieldStatementSyntax node)
        {
            return SingleStatement(SyntaxFactory.YieldStatement(SyntaxKind.YieldReturnStatement, (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.Expression.AcceptAsync(_expressionVisitor)));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitExitStatement(VBSyntax.ExitStatementSyntax node)
        {
            switch (VBasic.VisualBasicExtensions.Kind(node.BlockKeyword)) {
                case VBasic.SyntaxKind.SubKeyword:
                case VBasic.SyntaxKind.PropertyKeyword when node.GetAncestor<VBSyntax.AccessorBlockSyntax>()?.IsKind(VBasic.SyntaxKind.GetAccessorBlock) != true:
                    return SingleStatement(SyntaxFactory.ReturnStatement());
                case VBasic.SyntaxKind.FunctionKeyword:
                case VBasic.SyntaxKind.PropertyKeyword when node.GetAncestor<VBSyntax.AccessorBlockSyntax>()?.IsKind(VBasic.SyntaxKind.GetAccessorBlock) == true:
                    VBasic.VisualBasicSyntaxNode typeContainer = node.GetAncestor<VBSyntax.LambdaExpressionSyntax>()
                                                                 ?? (VBasic.VisualBasicSyntaxNode)node.GetAncestor<VBSyntax.MethodBlockSyntax>()
                                                                 ?? node.GetAncestor<VBSyntax.AccessorBlockSyntax>();
                    var enclosingMethodInfo = await typeContainer.TypeSwitch(
                        async (VBSyntax.LambdaExpressionSyntax e) => _semanticModel.GetSymbolInfo(e).Symbol,
                        async (VBSyntax.MethodBlockSyntax e) => _semanticModel.GetDeclaredSymbol(e),
                        async (VBSyntax.AccessorBlockSyntax e) => _semanticModel.GetDeclaredSymbol(e));
                    var info = enclosingMethodInfo?.GetReturnType();
                    Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax expr;
                    if (HasReturnVariable)
                        expr = ReturnVariable;
                    else if (info == null)
                        expr = null;
                    else if (IsIterator)
                        return SingleStatement(SyntaxFactory.YieldStatement(SyntaxKind.YieldBreakStatement));
                    else if (info.IsReferenceType)
                        expr = SyntaxFactory.LiteralExpression(SyntaxKind.NullLiteralExpression);
                    else if (info.CanBeReferencedByName)
                        expr = SyntaxFactory.DefaultExpression(CommonConversions.GetTypeSyntax(info));
                    else
                        throw new NotSupportedException();
                    return SingleStatement(SyntaxFactory.ReturnStatement(expr));
                default:
                    return SingleStatement(SyntaxFactory.BreakStatement());
            }
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitRaiseEventStatement(VBSyntax.RaiseEventStatementSyntax node)
        {
            var argumentListSyntax = (Microsoft.CodeAnalysis.CSharp.Syntax.ArgumentListSyntax) await node.ArgumentList.AcceptAsync(_expressionVisitor) ?? SyntaxFactory.ArgumentList();

            var symbolInfo = _semanticModel.GetSymbolInfo(node.Name).ExtractBestMatch() as IEventSymbol;
            if (symbolInfo?.RaiseMethod != null) {
                return SingleStatement(SyntaxFactory.InvocationExpression(
                    SyntaxFactory.IdentifierName($"On{symbolInfo.Name}"),
                    argumentListSyntax));
            }

            var memberBindingExpressionSyntax = SyntaxFactory.MemberBindingExpression(SyntaxFactory.IdentifierName("Invoke"));
            var conditionalAccessExpressionSyntax = SyntaxFactory.ConditionalAccessExpression(
                (Microsoft.CodeAnalysis.CSharp.Syntax.NameSyntax) await node.Name.AcceptAsync(_expressionVisitor),
                SyntaxFactory.InvocationExpression(memberBindingExpressionSyntax, argumentListSyntax)
            );
            return SingleStatement(
                conditionalAccessExpressionSyntax
            );
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitSingleLineIfStatement(VBSyntax.SingleLineIfStatementSyntax node)
        {
            var condition = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.Condition.AcceptAsync(_expressionVisitor);
            condition = CommonConversions.TypeConversionAnalyzer.AddExplicitConversion(node.Condition, condition, forceTargetType: _vbBooleanTypeSymbol);
            var block = SyntaxFactory.Block(await ConvertStatements(node.Statements));
            ElseClauseSyntax elseClause = null;

            if (node.ElseClause != null) {
                var elseBlock = SyntaxFactory.Block(await ConvertStatements(node.ElseClause.Statements));
                elseClause = SyntaxFactory.ElseClause(elseBlock.UnpackNonNestedBlock());
            }
            return SingleStatement(SyntaxFactory.IfStatement(condition, block.UnpackNonNestedBlock(), elseClause));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitMultiLineIfBlock(VBSyntax.MultiLineIfBlockSyntax node)
        {
            var condition = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.IfStatement.Condition.AcceptAsync(_expressionVisitor);
            condition = CommonConversions.TypeConversionAnalyzer.AddExplicitConversion(node.IfStatement.Condition, condition, forceTargetType: _vbBooleanTypeSymbol);
            var block = SyntaxFactory.Block(await ConvertStatements(node.Statements));
            ElseClauseSyntax elseClause = null;

            if (node.ElseBlock != null) {
                var elseBlock = SyntaxFactory.Block(await ConvertStatements(node.ElseBlock.Statements));
                // so that you get a neat "else if" at the end
                elseClause = await CommentConvertingVisitor.WithTrivia(node.ElseBlock, SyntaxFactory.ElseClause(elseBlock));
            }

            foreach (var elseIf in node.ElseIfBlocks.Reverse()) {

                var ifStmt = (Microsoft.CodeAnalysis.CSharp.Syntax.IfStatementSyntax) (await elseIf.Accept(CommentConvertingVisitor)).Single();
                elseClause = SyntaxFactory.ElseClause(ifStmt.WithElse(elseClause));
            }

            return SingleStatement(SyntaxFactory.IfStatement(condition, block, elseClause));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitElseIfBlock(VBSyntax.ElseIfBlockSyntax node)
        {
            var elseBlock = SyntaxFactory.Block(await ConvertStatements(node.Statements));
            var elseIfCondition = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)await node.ElseIfStatement.Condition.AcceptAsync(_expressionVisitor);
            elseIfCondition = CommonConversions.TypeConversionAnalyzer.AddExplicitConversion(node.ElseIfStatement.Condition, elseIfCondition, forceTargetType: _vbBooleanTypeSymbol);
            var ifStmt = SyntaxFactory.IfStatement(elseIfCondition, elseBlock);
            return SingleStatement(ifStmt);
        }

        private async Task<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax[]> ConvertStatements(SyntaxList<VBSyntax.StatementSyntax> statementSyntaxs)
        {
            return await statementSyntaxs.SelectManyAsync(async s => (IEnumerable<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>)await s.Accept(CommentConvertingVisitor));
        }

        /// <summary>
        /// See https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/for-next-statement#BKMK_Counter
        /// </summary>
        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitForBlock(VBSyntax.ForBlockSyntax node)
        {
            var stmt = node.ForStatement;
            var startValue = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await stmt.FromValue.AcceptAsync(_expressionVisitor);
            VariableDeclarationSyntax declaration = null;
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax id;
            var controlVarOp = _semanticModel.GetOperation(stmt.ControlVariable) as IVariableDeclaratorOperation;
            var controlVarType = controlVarOp?.Symbol.Type;
            var initializers = new List<Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax>();
            if (stmt.ControlVariable is VBSyntax.VariableDeclaratorSyntax) {
                var v = (VBSyntax.VariableDeclaratorSyntax)stmt.ControlVariable;
                declaration = (await CommonConversions.SplitVariableDeclarations(v)).Variables.Single();
                declaration = declaration.WithVariables(SyntaxFactory.SingletonSeparatedList(declaration.Variables[0].WithInitializer(SyntaxFactory.EqualsValueClause(startValue))));
                id = SyntaxFactory.IdentifierName(declaration.Variables[0].Identifier);
            } else {
                id = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await stmt.ControlVariable.AcceptAsync(_expressionVisitor);
                var controlVarSymbol = controlVarOp?.Symbol;
                if (controlVarSymbol != null && controlVarSymbol.DeclaringSyntaxReferences.Any(r => r.Span.OverlapsWith(stmt.ControlVariable.Span))) {
                    declaration = CommonConversions.CreateVariableDeclarationAndAssignment(controlVarSymbol.Name, startValue, CommonConversions.GetTypeSyntax(controlVarType));
                } else {
                    startValue = SyntaxFactory.AssignmentExpression(SyntaxKind.SimpleAssignmentExpression, id, startValue);
                    initializers.Add(startValue);
                }
            }

            var step = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await (stmt.StepClause?.StepValue).AcceptAsync(_expressionVisitor);
            PrefixUnaryExpressionSyntax value = step.SkipParens() as PrefixUnaryExpressionSyntax;
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax condition;

            // In Visual Basic, the To expression is only evaluated once, but in C# will be evaluated every loop.
            // If it could evaluate differently or has side effects, it must be extracted as a variable
            var preLoopStatements = new List<SyntaxNode>();
            var csToValue = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await stmt.ToValue.AcceptAsync(_expressionVisitor);
            if (!_semanticModel.GetConstantValue(stmt.ToValue).HasValue) {
                var loopToVariableName = GetUniqueVariableNameInScope(node, "loopTo");
                var toValueType = _semanticModel.GetTypeInfo(stmt.ToValue).ConvertedType;
                var toVariableId = SyntaxFactory.IdentifierName(loopToVariableName);
                if (controlVarType?.Equals(toValueType) == true && declaration != null) {
                    var loopToAssignment = CommonConversions.CreateVariableDeclarator(loopToVariableName, csToValue);
                    declaration = declaration.AddVariables(loopToAssignment);
                } else {
                    var loopEndDeclaration = SyntaxFactory.LocalDeclarationStatement(
                        CommonConversions.CreateVariableDeclarationAndAssignment(loopToVariableName, csToValue));
                    // Does not do anything about porting newline trivia upwards to maintain spacing above the loop
                    preLoopStatements.Add(loopEndDeclaration);
                }

                csToValue = toVariableId;
            };

            if (value == null) {
                condition = SyntaxFactory.BinaryExpression(SyntaxKind.LessThanOrEqualExpression, id, csToValue);
            } else {
                condition = SyntaxFactory.BinaryExpression(SyntaxKind.GreaterThanOrEqualExpression, id, csToValue);
            }
            if (step == null)
                step = SyntaxFactory.PostfixUnaryExpression(SyntaxKind.PostIncrementExpression, id);
            else
                step = SyntaxFactory.AssignmentExpression(SyntaxKind.AddAssignmentExpression, id, step);
            var block = SyntaxFactory.Block(await ConvertStatements(node.Statements));
            var forStatementSyntax = SyntaxFactory.ForStatement(
                declaration,
                SyntaxFactory.SeparatedList(initializers),
                condition,
                SyntaxFactory.SingletonSeparatedList(step),
                block.UnpackNonNestedBlock());
            return SyntaxFactory.List(preLoopStatements.Concat(new[] {forStatementSyntax}));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitForEachBlock(VBSyntax.ForEachBlockSyntax node)
        {
            var stmt = node.ForEachStatement;

            Microsoft.CodeAnalysis.CSharp.Syntax.TypeSyntax type;
            SyntaxToken id;
            if (stmt.ControlVariable is VBSyntax.VariableDeclaratorSyntax vds) {
                var declaration = (await CommonConversions.SplitVariableDeclarations(vds)).Variables.Single();
                type = declaration.Type;
                id = declaration.Variables.Single().Identifier;
            } else {
                var v = (Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax) await stmt.ControlVariable.AcceptAsync(_expressionVisitor);
                id = v.Identifier;
                type = SyntaxFactory.ParseTypeName("var");
            }

            var block = SyntaxFactory.Block(await ConvertStatements(node.Statements));
            var csExpression = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)await stmt.Expression.AcceptAsync(_expressionVisitor);
            return SingleStatement(SyntaxFactory.ForEachStatement(
                type,
                id,
                CommonConversions.TypeConversionAnalyzer.AddExplicitConversion(stmt.Expression, csExpression),
                block.UnpackNonNestedBlock()
            ));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitLabelStatement(VBSyntax.LabelStatementSyntax node)
        {
            return SingleStatement(SyntaxFactory.LabeledStatement(node.LabelToken.Text, SyntaxFactory.EmptyStatement()));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitGoToStatement(VBSyntax.GoToStatementSyntax node)
        {
            return SingleStatement(SyntaxFactory.GotoStatement(SyntaxKind.GotoStatement,
                SyntaxFactory.IdentifierName(node.Label.LabelToken.Text)));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitSelectBlock(VBSyntax.SelectBlockSyntax node)
        {
            var vbExpr = node.SelectStatement.Expression;
            var (reusableExpr, stmts) = await GetReusableExpression(vbExpr, "switchExpr");

            var usedConstantValues = new HashSet<object>();
            var sections = new List<SwitchSectionSyntax>();
            foreach (var block in node.CaseBlocks) {
                var labels = new List<SwitchLabelSyntax>();
                foreach (var c in block.CaseStatement.Cases) {
                    if (c is VBSyntax.SimpleCaseClauseSyntax s) {
                        var originalExpressionSyntax = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)await s.Value.AcceptAsync(_expressionVisitor);
                        // CSharp requires an explicit cast from the base type (e.g. int) in most cases switching on an enum
                        var typeConversionKind = CommonConversions.TypeConversionAnalyzer.AnalyzeConversion(s.Value);
                        var expressionSyntax = CommonConversions.TypeConversionAnalyzer.AddExplicitConversion(s.Value, originalExpressionSyntax, typeConversionKind, true);
                        SwitchLabelSyntax caseSwitchLabelSyntax = SyntaxFactory.CaseSwitchLabel(expressionSyntax);
                        var constantValue = _semanticModel.GetConstantValue(s.Value);
                        var isRepeatedConstantValue = constantValue.HasValue && !usedConstantValues.Add(constantValue);
                        if (!constantValue.HasValue || isRepeatedConstantValue ||
                            (typeConversionKind != TypeConversionAnalyzer.TypeConversionKind.NonDestructiveCast &&
                             typeConversionKind != TypeConversionAnalyzer.TypeConversionKind.Identity)) {
                            caseSwitchLabelSyntax =
                                WrapInCasePatternSwitchLabelSyntax(node, expressionSyntax);
                        }
                        labels.Add(caseSwitchLabelSyntax);
                    } else if (c is VBSyntax.ElseCaseClauseSyntax) {
                        labels.Add(SyntaxFactory.DefaultSwitchLabel());
                    } else if (c is VBSyntax.RelationalCaseClauseSyntax relational) {
                        var operatorKind = VBasic.VisualBasicExtensions.Kind(relational);
                        var cSharpSyntaxNode = SyntaxFactory.BinaryExpression(operatorKind.ConvertToken(TokenContext.Local), reusableExpr, (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)await relational.Value.AcceptAsync(_expressionVisitor));
                        labels.Add(WrapInCasePatternSwitchLabelSyntax(node, cSharpSyntaxNode, treatAsBoolean: true));
                    } else if (c is VBSyntax.RangeCaseClauseSyntax range) {
                        var lowerBoundCheck = SyntaxFactory.BinaryExpression(SyntaxKind.LessThanOrEqualExpression, (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)await range.LowerBound.AcceptAsync(_expressionVisitor), reusableExpr);
                        var upperBoundCheck = SyntaxFactory.BinaryExpression(SyntaxKind.LessThanOrEqualExpression, reusableExpr, (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)await range.UpperBound.AcceptAsync(_expressionVisitor));
                        var withinBounds = SyntaxFactory.BinaryExpression(SyntaxKind.LogicalAndExpression, lowerBoundCheck, upperBoundCheck);
                        labels.Add(WrapInCasePatternSwitchLabelSyntax(node, withinBounds, treatAsBoolean: true));
                    } else throw new NotSupportedException(c.Kind().ToString());
                }

                var csBlockStatements = (await ConvertStatements(block.Statements)).ToList();
                if (csBlockStatements.LastOrDefault()
                        ?.IsKind(SyntaxKind.ReturnStatement, SyntaxKind.BreakStatement) != true) {
                    csBlockStatements.Add(SyntaxFactory.BreakStatement());
                }
                var list = SingleStatement(SyntaxFactory.Block(csBlockStatements));
                sections.Add(SyntaxFactory.SwitchSection(SyntaxFactory.List(labels), list));
            }

            var switchStatementSyntax = ValidSyntaxFactory.SwitchStatement(reusableExpr, sections);
            return stmts.Add(switchStatementSyntax);
        }

        private async Task<(Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax, SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>)> GetReusableExpression(VBSyntax.ExpressionSyntax vbExpr, string variableNameBase, bool forceVariable = false)
        {
            var expr = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax)await vbExpr.AcceptAsync(_expressionVisitor);
            SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax> stmts = SyntaxFactory.List<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>();
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax reusableExpr;
            if (forceVariable || !await CanEvaluateMultipleTimesAsync(vbExpr)) {
                var contextNode = vbExpr.GetAncestor<VBSyntax.MethodBlockBaseSyntax>() ?? (VBasic.VisualBasicSyntaxNode) vbExpr.Parent;
                var varName = GetUniqueVariableNameInScope(contextNode, variableNameBase);
                var stmt = CreateLocalVariableDeclarationAndAssignment(varName, expr);
                stmts = stmts.Add(stmt);
                reusableExpr = SyntaxFactory.IdentifierName(varName);
            } else {
                reusableExpr = expr.WithoutTrivia().WithoutAnnotations();
            }

            return (reusableExpr, stmts);
        }

        private async Task<bool> CanEvaluateMultipleTimesAsync(VBSyntax.ExpressionSyntax vbExpr)
        {
            return _semanticModel.GetConstantValue(vbExpr).HasValue || vbExpr.SkipParens() is VBSyntax.NameSyntax ns && await IsNeverMutatedAsync(ns);
        }

        private async Task<bool> IsNeverMutatedAsync(VBSyntax.NameSyntax ns)
        {
            var allowedLocation = Location.Create(ns.SyntaxTree, TextSpan.FromBounds(ns.GetAncestor<VBSyntax.MethodBlockBaseSyntax>().SpanStart, ns.Span.End));
            return await CommonConversions.Document.Project.Solution.IsNeverWritten(_semanticModel.GetSymbolInfo(ns).Symbol, allowedLocation);
        }

        private CasePatternSwitchLabelSyntax WrapInCasePatternSwitchLabelSyntax(VBSyntax.SelectBlockSyntax node, Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax cSharpSyntaxNode, bool treatAsBoolean = false)
        {
            var typeInfo = _semanticModel.GetTypeInfo(node.SelectStatement.Expression);

            DeclarationPatternSyntax patternMatch;
            if (typeInfo.ConvertedType.SpecialType == SpecialType.System_Boolean || treatAsBoolean) {
                patternMatch = SyntaxFactory.DeclarationPattern(
                    SyntaxFactory.PredefinedType(SyntaxFactory.Token(SyntaxKind.ObjectKeyword)),
                    SyntaxFactory.DiscardDesignation());
            } else {
                var varName = CommonConversions.ConvertIdentifier(SyntaxFactory.Identifier(GetUniqueVariableNameInScope(node, "case"))).ValueText;
                patternMatch = SyntaxFactory.DeclarationPattern(
                    SyntaxFactory.ParseTypeName("var"), SyntaxFactory.SingleVariableDesignation(SyntaxFactory.Identifier(varName)));
                cSharpSyntaxNode = SyntaxFactory.BinaryExpression(SyntaxKind.EqualsExpression, SyntaxFactory.IdentifierName(varName), cSharpSyntaxNode);
            }

            var casePatternSwitchLabelSyntax = SyntaxFactory.CasePatternSwitchLabel(patternMatch,
                SyntaxFactory.WhenClause(cSharpSyntaxNode), SyntaxFactory.Token(SyntaxKind.ColonToken));
            return casePatternSwitchLabelSyntax;
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitWithBlock(VBSyntax.WithBlockSyntax node)
        {
            var (lhsExpression, prefixDeclarations) = await GetReusableExpression(node.WithStatement.Expression, "withBlock");

            _withBlockLhs.Push(lhsExpression);
            try {
                var statements = await ConvertStatements(node.Statements);

                var statementSyntaxs = SyntaxFactory.List(prefixDeclarations.Concat(statements));
                return prefixDeclarations.Any()
                    ? SingleStatement(SyntaxFactory.Block(statementSyntaxs))
                    : statementSyntaxs;
            } finally {
                _withBlockLhs.Pop();
            }
        }

        private Microsoft.CodeAnalysis.CSharp.Syntax.LocalDeclarationStatementSyntax CreateLocalVariableDeclarationAndAssignment(string variableName, Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax initValue)
        {
            return SyntaxFactory.LocalDeclarationStatement(CommonConversions.CreateVariableDeclarationAndAssignment(variableName, initValue));
        }

        private string GetUniqueVariableNameInScope(VBasic.VisualBasicSyntaxNode node, string variableNameBase)
        {
            return NameGenerator.GetUniqueVariableNameInScope(_semanticModel, _generatedNames, node, variableNameBase);
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitTryBlock(VBSyntax.TryBlockSyntax node)
        {
            var block = SyntaxFactory.Block(await ConvertStatements(node.Statements));
            return SingleStatement(
                SyntaxFactory.TryStatement(
                    block,
                    SyntaxFactory.List(await node.CatchBlocks.SelectAsync(async c => (CatchClauseSyntax) await c.AcceptAsync(_expressionVisitor))),
                    (FinallyClauseSyntax) await node.FinallyBlock.AcceptAsync(_expressionVisitor)
                )
            );
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitSyncLockBlock(VBSyntax.SyncLockBlockSyntax node)
        {
            return SingleStatement(SyntaxFactory.LockStatement(
                (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.SyncLockStatement.Expression.AcceptAsync(_expressionVisitor),
                SyntaxFactory.Block(await ConvertStatements(node.Statements)).UnpackNonNestedBlock()
            ));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitUsingBlock(VBSyntax.UsingBlockSyntax node)
        {
            var statementSyntax = SyntaxFactory.Block(await ConvertStatements(node.Statements));
            if (node.UsingStatement.Expression == null) {
                Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax stmt = statementSyntax;
                foreach (var v in node.UsingStatement.Variables.Reverse())
                foreach (var declaration in (await CommonConversions.SplitVariableDeclarations(v)).Variables.Reverse())
                    stmt = SyntaxFactory.UsingStatement(declaration, null, stmt);
                return SingleStatement(stmt);
            }

            var expr = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.UsingStatement.Expression.AcceptAsync(_expressionVisitor);
            var unpackPossiblyNestedBlock = statementSyntax.UnpackPossiblyNestedBlock(); // Allow reduced indentation for multiple usings in a row
            return SingleStatement(SyntaxFactory.UsingStatement(null, expr, unpackPossiblyNestedBlock));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitWhileBlock(VBSyntax.WhileBlockSyntax node)
        {
            return SingleStatement(SyntaxFactory.WhileStatement(
                (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.WhileStatement.Condition.AcceptAsync(_expressionVisitor),
                SyntaxFactory.Block(await ConvertStatements(node.Statements)).UnpackNonNestedBlock()
            ));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitDoLoopBlock(VBSyntax.DoLoopBlockSyntax node)
        {
            var statements = SyntaxFactory.Block(await ConvertStatements(node.Statements)).UnpackNonNestedBlock();

            if (node.DoStatement.WhileOrUntilClause != null) {
                var stmt = node.DoStatement.WhileOrUntilClause;
                if (SyntaxTokenExtensions.IsKind(stmt.WhileOrUntilKeyword, VBasic.SyntaxKind.WhileKeyword))
                    return SingleStatement(SyntaxFactory.WhileStatement(
                        (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await stmt.Condition.AcceptAsync(_expressionVisitor),
                        statements
                    ));
                return SingleStatement(SyntaxFactory.WhileStatement(
                    ((Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await stmt.Condition.AcceptAsync(_expressionVisitor)).InvertCondition(),
                    statements
                ));
            }

            var whileOrUntilStmt = node.LoopStatement.WhileOrUntilClause;
            Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax conditionExpression;
            bool isUntilStmt;
            if (whileOrUntilStmt != null) {
                conditionExpression = (Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await whileOrUntilStmt.Condition.AcceptAsync(_expressionVisitor);
                isUntilStmt = SyntaxTokenExtensions.IsKind(whileOrUntilStmt.WhileOrUntilKeyword, VBasic.SyntaxKind.UntilKeyword);
            } else {
                conditionExpression = SyntaxFactory.LiteralExpression(SyntaxKind.TrueLiteralExpression);
                isUntilStmt = false;
            }

            if (isUntilStmt) {
                conditionExpression = conditionExpression.InvertCondition();
            }

            return SingleStatement(SyntaxFactory.DoStatement(statements, conditionExpression));
        }

        public override async Task<SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>> VisitCallStatement(VBSyntax.CallStatementSyntax node)
        {
            return SingleStatement((Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax) await node.Invocation.AcceptAsync(_expressionVisitor));
        }

        private SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax> SingleStatement(Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax statement)
        {
            return SyntaxFactory.SingletonList(statement);
        }

        private SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax> SingleStatement(Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax expression)
        {
            return SyntaxFactory.SingletonList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>(SyntaxFactory.ExpressionStatement(expression));
        }
    }

    internal static class Extensions
    {
        /// <summary>
        /// Returns the single statement in a block if it has no nested statements.
        /// If it has nested statements, and the surrounding block was removed, it could be ambiguous,
        /// e.g. if (...) { if (...) return null; } else return "";
        /// Unbundling the middle if statement would bind the else to it, rather than the outer if statement
        /// </summary>
        public static Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax UnpackNonNestedBlock(this BlockSyntax block)
        {
            return block.Statements.Count == 1 && !block.ContainsNestedStatements() ? block.Statements[0] : block;
        }

        /// <summary>
        /// Returns the single statement in a block
        /// </summary>
        public static bool TryUnpackSingleStatement(this IReadOnlyCollection<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax> statements, out Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax singleStatement)
        {
            singleStatement = statements.Count == 1 ? statements.Single() : null;
            if (singleStatement is BlockSyntax block && TryUnpackSingleStatement(block.Statements, out var s)) {
                singleStatement = s;
            }

            return singleStatement != null;
        }

        /// <summary>
        /// Returns the single expression in a statement
        /// </summary>
        public static bool TryUnpackSingleExpressionFromStatement(this Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax statement, out Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionSyntax singleExpression)
        {
            switch(statement){
                case BlockSyntax blockSyntax:
                    singleExpression = null;
                    return TryUnpackSingleStatement(blockSyntax.Statements, out var nestedStmt) &&
                           TryUnpackSingleExpressionFromStatement(nestedStmt, out singleExpression);
                case Microsoft.CodeAnalysis.CSharp.Syntax.ExpressionStatementSyntax expressionStatementSyntax:
                    singleExpression = expressionStatementSyntax.Expression;
                    return singleExpression != null;
                case Microsoft.CodeAnalysis.CSharp.Syntax.ReturnStatementSyntax returnStatementSyntax:
                    singleExpression = returnStatementSyntax.Expression;
                    return singleExpression != null;
                default:
                    singleExpression = null;
                    return false;
            }
        }

        /// <summary>
        /// Only use this over <see cref="UnpackNonNestedBlock"/> in special cases where it will display more neatly and where you're sure nested statements don't introduce ambiguity
        /// </summary>
        public static Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax UnpackPossiblyNestedBlock(this BlockSyntax block)
        {
            SyntaxList<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax> statementSyntaxs = block.Statements;
            return statementSyntaxs.Count == 1 ? statementSyntaxs[0] : block;
        }

        private static bool ContainsNestedStatements(this BlockSyntax block)
        {
            return block.Statements.Any(HasDescendantCSharpStatement);
        }

        private static bool HasDescendantCSharpStatement(this Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax c)
        {
            return c.DescendantNodes().OfType<Microsoft.CodeAnalysis.CSharp.Syntax.StatementSyntax>().Any();
        }
    }
}
