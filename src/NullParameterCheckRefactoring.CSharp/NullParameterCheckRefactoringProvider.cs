using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CodeActions;
using Microsoft.CodeAnalysis.CodeRefactorings;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Extensions;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Formatting;
using System;
using System.Collections.Generic;
using System.Composition;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using RoslynExts.CS;


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

      if (parameterSyntax != null)
      {
        TypeSyntax paramTypeName = parameterSyntax.Type;
        SemanticModel semanticModel = await context.Document.GetSemanticModelAsync(context.CancellationToken);
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
                                                                      if (expression.Right.IsKind(SyntaxKind.NullLiteralExpression))
                                                                      {
                                                                        var identifierSyntaxt = expression.ChildNodes().OfType<IdentifierNameSyntax>().FirstOrDefault();
                                                                        return (identifierSyntaxt != null) &&
                                                                               (identifierSyntaxt.Identifier.Text.Equals(parameterSyntax.Identifier.Text, StringComparison.Ordinal));
                                                                      }
                                                                      return false;
                                                                    });
                                           });

          if (isNullCheckAlreadyPresent == false)
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
      
      var nullCheckIfStatement =
        "if( \{parameter.Identifier} == null ) { throw new ArgumentNullException( nameof( \{parameter.Identifier.Text} )); };".ToSExpr<IfStatementSyntax>();
      SyntaxList<SyntaxNode> newStatements = methodDeclaration.Body.Statements.Insert(0, nullCheckIfStatement);
      BlockSyntax newBlock = SyntaxFactory.Block(newStatements).WithAdditionalAnnotations(Formatter.Annotation);
      SyntaxNode root = await document.GetSyntaxRootAsync(cancellationToken);
      SyntaxNode newRoot = root.ReplaceNode(methodDeclaration.Body, newBlock);

      return document.WithSyntaxRoot(newRoot);
    }
  }
}