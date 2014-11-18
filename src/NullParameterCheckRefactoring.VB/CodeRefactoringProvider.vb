<ExportCodeRefactoringProvider(NullCheck_CodeRefactoringCodeRefactoringProvider.RefactoringId, LanguageNames.VisualBasic), [Shared]>
Friend Class NullCheck_CodeRefactoringCodeRefactoringProvider
  Inherits CodeRefactoringProvider

  Public Const RefactoringId As String = "TR0001"

  Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
    Dim root = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)
    ' Find the node at the selection.
    Dim node = root.FindNode(context.Span)
    ' Only offer a refactoring if the selected node is a type statement node.
    Dim _parmeter_ = TryCast(node, ParameterSyntax)
    If _parmeter_ Is Nothing Then Return
    Dim _method_ = TryCast(_parmeter_.Parent.Parent.Parent, MethodBlockSyntax)
    If _method_ Is Nothing Then Exit Function
    Dim _Model_ = Await context.Document.GetSemanticModelAsync(context.CancellationToken)
    Dim ifStatements = _method_.Statements.Where(Function(s) (TypeOf s Is MultiLineIfBlockSyntax) OrElse (TypeOf s Is SingleLineIfStatementSyntax))
    Dim pinfo = _Model_.GetTypeInfo(_parmeter_.AsClause.Type, context.CancellationToken)
    If pinfo.ConvertedType.IsReferenceType = False Then Return 
    Dim IsNullCheckAlreadyPresent = ifStatements.Any(NullChecks(_parmeter_))
    If Not IsNullCheckAlreadyPresent Then context.RegisterRefactoring(CodeAction.Create("Check Parameter for null", Function(ct As CancellationToken) AddParameterNullCheckAsync(context.Document, _parmeter_, _method_, ct)))
  End Function

  Private Shared Function NullChecks(_parmeter_ As ParameterSyntax) As Func(Of StatementSyntax, Boolean)
    Return Function(s)
             If TypeOf s Is SingleLineIfStatementSyntax Then
               Dim isExpr = TryCast(DirectCast(s, SingleLineIfStatementSyntax).Condition, BinaryExpressionSyntax)
               If (isExpr Is Nothing) OrElse (Not isExpr.IsKind(SyntaxKind.IsExpression)) Then Return False
               Return CheckIfCondition(_parmeter_, isExpr)
             ElseIf TypeOf s Is MultiLineIfBlockSyntax Then
               Dim isExpr = TryCast(DirectCast(s, MultiLineIfBlockSyntax).IfStatement.Condition, BinaryExpressionSyntax)
               If (isExpr Is Nothing) OrElse (Not isExpr.IsKind(SyntaxKind.IsExpression)) Then Return False
               Return CheckIfCondition(_parmeter_, isExpr)
             End If
             Return False
           End Function
  End Function

  Private Shared Function CheckIfCondition(paramSyntax As ParameterSyntax, isExpr As BinaryExpressionSyntax) As Boolean
    Dim l = TryCast(isExpr.Left, IdentifierNameSyntax)
    Dim r = TryCast(isExpr.Right, LiteralExpressionSyntax)
    If l Is Nothing OrElse r Is Nothing Then Return False
    If r.IsKind(SyntaxKind.NothingLiteralExpression) = False Then Return False
    Return String.Compare(l.Identifier.Text, paramSyntax.Identifier.Identifier.Text, StringComparison.Ordinal) = 0
  End Function

  Private Async Function AddParameterNullCheckAsync(document As Document,
                                                    parameterStmt As ParameterSyntax,
                                                    method As MethodBlockSyntax,
                                                    cancellationToken As CancellationToken) As Task(Of Document)
    Dim _null_ = SyntaxFactory.NothingLiteralExpression(SyntaxFactory.Token(SyntaxKind.NothingKeyword))
    Dim _IsExpr_ = SyntaxFactory.IsExpression(
                     SyntaxFactory.IdentifierName(parameterStmt.Identifier.Identifier.Text),
                     _null_).WithAdditionalAnnotations(Formatting.Formatter.Annotation)
    ' Note: If I can find the nameof feature in VB.net, then I'll change this line to reflect that
    Dim _paramname_ = SyntaxFactory.StringLiteralExpression(
                        SyntaxFactory.StringLiteralToken("""" & parameterStmt.Identifier.Identifier.Text & """",
                                                         parameterStmt.Identifier.Identifier.Text))
    Dim st = SyntaxFactory.ObjectCreationExpression(Nothing,
                              SyntaxFactory.ParseTypeName(GetType(ArgumentNullException).FullName),
                              SyntaxFactory.ArgumentList().AddArguments(
                                SyntaxFactory.SimpleArgument(_paramname_)
                              ),
                              Nothing)

    Dim throwExpr = SyntaxFactory.ThrowStatement(st)

    Dim if_ = SyntaxFactory.SingleLineIfStatement(
                SyntaxFactory.Token(SyntaxKind.IfKeyword),
                _IsExpr_,
                SyntaxFactory.Token(SyntaxKind.ThenKeyword),
                New SyntaxList(Of StatementSyntax)().Add(throwExpr),
                Nothing).WithAdditionalAnnotations(Formatting.Formatter.Annotation)

    Dim newStatements = method.Statements.Insert(0, if_)
    Dim newBlock = method.WithStatements(newStatements)
    Return document.WithSyntaxRoot((Await document.GetSyntaxRootAsync(cancellationToken)).ReplaceNode(method, newBlock))
  End Function
End Class