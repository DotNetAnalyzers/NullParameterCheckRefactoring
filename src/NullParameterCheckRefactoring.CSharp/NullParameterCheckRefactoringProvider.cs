using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeRefactorings;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Formatting;
using Microsoft.CodeAnalysis.Simplification;
using System;
using System.Collections.Generic;
using System.Composition;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace NullParameterCheckRefactoring
{
    [ExportCodeRefactoringProvider(RefactoringId, LanguageNames.CSharp), Shared]
    public class NullParameterCheckRefactoringProvider : CodeRefactoringProvider
    {
        internal const string RefactoringId = "TR0001";

        public override async Task ComputeRefactoringsAsync(CodeRefactoringContext context)
        {
            SyntaxNode root = await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(false);
            SyntaxNode node = root.FindNode(context.Span);
            ParameterSyntax parameterSyntax = node as ParameterSyntax;

            if(parameterSyntax != null)
            {
                TypeSyntax paramTypeName = parameterSyntax.Type;
                SemanticModel semanticModel = await context.Document.GetSemanticModelAsync(context.CancellationToken);

                // make sure ArgumentNullException is available
                INamedTypeSymbol argumentNullExceptionSymbol = semanticModel.Compilation.GetTypeByMetadataName(typeof(ArgumentNullException).FullName);
                if (argumentNullExceptionSymbol == null)
                {
                    return;
                }

                ITypeSymbol type = semanticModel.GetTypeInfo(paramTypeName).ConvertedType;

                BaseMethodDeclarationSyntax methodDeclaration = parameterSyntax.Parent.Parent as BaseMethodDeclarationSyntax;
                IEnumerable<IfStatementSyntax> availableIfStatements = methodDeclaration.Body.ChildNodes().OfType<IfStatementSyntax>();

                if (type.IsReferenceType)
                {
                    // check if the null check already exists.
                    bool isNullCheckAlreadyPresent = availableIfStatements.Any(ifStatement =>
                    {
                        return ifStatement.ChildNodes()
                            .OfType<BinaryExpressionSyntax>()
                            .Where(x => x.IsKind(SyntaxKind.EqualsExpression))
                            .Any(expression =>
                            {
                                bool result = false;
                                bool isNullCheck = expression.Right.IsKind(SyntaxKind.NullLiteralExpression);
                                if (isNullCheck)
                                {
                                    IdentifierNameSyntax identifierSyntaxt = expression.ChildNodes().OfType<IdentifierNameSyntax>().FirstOrDefault();
                                    if (identifierSyntaxt != null)
                                    {
                                        string identifierText = identifierSyntaxt.Identifier.Text;
                                        string paramText = parameterSyntax.Identifier.Text;

                                        if (identifierText.Equals(paramText, StringComparison.Ordinal))
                                        {
                                            // There is already a null check for this parameter. Skip it.
                                            result = true;
                                        }
                                    }
                                }
                                return result;
                            });
                    });

                    if(isNullCheckAlreadyPresent == false)
                    {
                        CodeAction action = CodeAction.Create(
                            "Check parameter for null",
                            ct => AddParameterNullCheckAsync(context.Document, parameterSyntax, methodDeclaration, ct));

                        context.RegisterRefactoring(action);
                    }
                }
            }
        }

        private async Task<Document> AddParameterNullCheckAsync(Document document, ParameterSyntax parameter, BaseMethodDeclarationSyntax methodDeclaration, CancellationToken cancellationToken)
        {
            BinaryExpressionSyntax binaryExpression = SyntaxFactory.BinaryExpression(
                SyntaxKind.EqualsExpression,
                SyntaxFactory.IdentifierName(parameter.Identifier),
                SyntaxFactory.LiteralExpression(SyntaxKind.NullLiteralExpression));

            ExpressionSyntax parameterNameExpression;
            if((document.Project.ParseOptions as CSharpParseOptions)?.LanguageVersion >= LanguageVersion.CSharp6)
            {
                parameterNameExpression = SyntaxFactory.NameOfExpression(
                    "nameof",
                    SyntaxFactory.ParseExpression(parameter.Identifier.Text));
            }
            else
            {
                parameterNameExpression = SyntaxFactory.LiteralExpression(
                    SyntaxKind.StringLiteralExpression,
                    SyntaxFactory.Literal(parameter.Identifier.ValueText));
            }

            ObjectCreationExpressionSyntax objectCreationExpression = SyntaxFactory.ObjectCreationExpression(
                SyntaxFactory.ParseTypeName("global::" + typeof(ArgumentNullException).FullName),
                SyntaxFactory.ArgumentList(SyntaxFactory.SeparatedList(new[] { SyntaxFactory.Argument(parameterNameExpression) })),
                null);

            BlockSyntax syntaxBlock = SyntaxFactory.Block(
                SyntaxFactory.Token(SyntaxKind.OpenBraceToken),
                new SyntaxList<StatementSyntax>().Add(SyntaxFactory.ThrowStatement(objectCreationExpression)),
                SyntaxFactory.Token(SyntaxKind.CloseBraceToken));

            IfStatementSyntax nullCheckIfStatement = SyntaxFactory.IfStatement(
                    SyntaxFactory.Token(SyntaxKind.IfKeyword),
                    SyntaxFactory.Token(SyntaxKind.OpenParenToken),
                    binaryExpression, 
                    SyntaxFactory.Token(SyntaxKind.CloseParenToken), 
                    syntaxBlock, null).WithAdditionalAnnotations(Formatter.Annotation, Simplifier.Annotation);

            SyntaxList<SyntaxNode> newStatements = methodDeclaration.Body.Statements.Insert(0, nullCheckIfStatement);
            BlockSyntax newBlock = methodDeclaration.Body.WithStatements(newStatements);
            SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken);
            SyntaxNode newRoot = root.ReplaceNode(methodDeclaration.Body, newBlock);

            return document.WithSyntaxRoot(newRoot);
        }
    }
}
