Imports Microsoft.CodeAnalysis.VisualBasic.SyntaxFactory
'Imports Microsoft.CodeAnalysis.VisualBasic.SyntaxKind
Imports DotNetAnalyzers.RoslynExts.VB


<ExportCodeRefactoringProvider(NullCheck_CodeRefactoringCodeRefactoringProvider.RefactoringId, LanguageNames.VisualBasic), [Shared]>
Friend Class NullCheck_CodeRefactoringCodeRefactoringProvider
  Inherits CodeRefactoringProvider

  Public Const RefactoringId As String = "TR0001"

  Public NotOverridable Overrides Async Function ComputeRefactoringsAsync(context As CodeRefactoringContext) As Task
    Dim root = Await context.Document.GetSyntaxRootAsync(context.CancellationToken).ConfigureAwait(False)
    ' Find the node at the selection.
    Dim node = root.FindNode(context.Span)
    ' Only offer a refactoring if the selected node is a type statement node.
    Dim _Identifier_ = TryCast(node, ModifiedIdentifierSyntax)
    If _Identifier_ Is Nothing Then Return
    Dim _parmeter_ = TryCast(_Identifier_.Parent, ParameterSyntax)
    If _parmeter_ Is Nothing Then Return
    Dim _method_ = TryCast(_parmeter_.Parent.Parent.Parent, MethodBlockSyntax)
    If _method_ Is Nothing Then Exit Function

    Dim _Model_ = Await context.Document.GetSemanticModelAsync(context.CancellationToken)
    Dim ifStatements = _method_.Statements.Where(Function(s) (TypeOf s Is MultiLineIfBlockSyntax) OrElse (TypeOf s Is SingleLineIfStatementSyntax))
    Dim pinfo = _Model_.GetTypeInfo(_parmeter_.AsClause.Type, context.CancellationToken)
    If pinfo.ConvertedType.IsReferenceType = False Then Return
    Dim GuardStatements = ifStatements.Where(NullChecks(_parmeter_))

    Dim IsNullCheckAlreadyPresent = GuardStatements.Any
    If Not IsNullCheckAlreadyPresent Then context.RegisterRefactoring(CodeAction.Create("Check Parameter for null", Function(ct As CancellationToken) AddParameterNullCheckAsync(context.Document, _parmeter_, _method_, ct)))
  End Function

  Private Shared Function NullChecks(_parmeter_ As ParameterSyntax) As Func(Of StatementSyntax, Boolean)
    Return Function(s)
             If TypeOf s Is SingleLineIfStatementSyntax Then
               Dim singleIF = DirectCast(s, SingleLineIfStatementSyntax)
               Dim isExpr = TryCast(singleIF.Condition, BinaryExpressionSyntax)
               If (isExpr Is Nothing) OrElse (Not isExpr.IsKind(SyntaxKind.IsExpression)) Then Return False
               Dim res = CheckIfCondition(_parmeter_, isExpr)
               If Not res Then Return res
               Dim _sif_ = GetSingleIFstatement(_parmeter_).NormalizeWhitespace
               Return Not singleIF.IsEquivalentTo(_sif_)
             ElseIf TypeOf s Is MultiLineIfBlockSyntax Then
               Dim multiIF = DirectCast(s, MultiLineIfBlockSyntax)
               Dim isExpr = TryCast(multiIF.IfStatement.Condition, BinaryExpressionSyntax)
               If (isExpr Is Nothing) OrElse (Not isExpr.IsKind(SyntaxKind.IsExpression)) Then Return False
               Dim res = CheckIfCondition(_parmeter_, isExpr)
               If Not res Then Return res
               Dim _mif_ = GetMultiLineIFstatement(_parmeter_)
               Return Not multiIF.WithoutAnnotations().IsEquivalentTo(_mif_)
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

  Private Shared Function CheckIfCondition(isExpr As BinaryExpressionSyntax) As Boolean
    Dim l = TryCast(isExpr.Left, IdentifierNameSyntax)
    Dim r = TryCast(isExpr.Right, LiteralExpressionSyntax)
    If l Is Nothing OrElse r Is Nothing Then Return False
    Return r.IsKind(SyntaxKind.NothingLiteralExpression)
  End Function

  Private Async Function AddParameterNullCheckAsync(document As Document,
                                                    _parmeter_ As ParameterSyntax,
                                                    method As MethodBlockSyntax,
                                                    cancellationToken As CancellationToken) As Task(Of Document)

    Dim NewGuard As SingleLineIfStatementSyntax = GetSingleIFstatement(_parmeter_)

    Dim ifStatements = method.Statements.Where(Function(s) (TypeOf s Is MultiLineIfBlockSyntax) OrElse (TypeOf s Is SingleLineIfStatementSyntax))
    Dim ExistingGuards = ifStatements.Where(Function(s)
                                              If TypeOf s Is SingleLineIfStatementSyntax Then
                                                Dim singleIF = DirectCast(s, SingleLineIfStatementSyntax)
                                                Dim isExpr = TryCast(singleIF.Condition, BinaryExpressionSyntax)
                                                Return CheckIfCondition(isExpr)
                                              ElseIf TypeOf s Is MultiLineIfBlockSyntax Then
                                                Dim multiIF = DirectCast(s, MultiLineIfBlockSyntax)
                                                Dim isExpr = TryCast(multiIF.IfStatement.Condition, BinaryExpressionSyntax)
                                                Return CheckIfCondition(isExpr)
                                              Else
                                                Return False
                                              End If
                                            End Function).ToList()

    Dim ExistingGuardParameters = ExistingGuards.Select(Function(s)
                                                          If TypeOf s Is SingleLineIfStatementSyntax Then
                                                            Dim singleIF = DirectCast(s, SingleLineIfStatementSyntax)
                                                            Dim isExpr = TryCast(singleIF.Condition, BinaryExpressionSyntax)
                                                            Return DirectCast(isExpr.Left, IdentifierNameSyntax)
                                                          ElseIf TypeOf s Is MultiLineIfBlockSyntax Then
                                                            Dim multiIF = DirectCast(s, MultiLineIfBlockSyntax)
                                                            Dim isExpr = TryCast(multiIF.IfStatement.Condition, BinaryExpressionSyntax)
                                                            Return DirectCast(isExpr.Left, IdentifierNameSyntax)
                                                          Else
                                                            Return Nothing
                                                          End If
                                                        End Function).ToList()


    Dim parameters = method.Begin.ParameterList.Parameters
    Dim pCount = parameters.Count
    Dim _Model_ = Await document.GetSemanticModelAsync(cancellationToken)

    Dim Guards As New List(Of StatementSyntax)(Enumerable.Repeat(Of StatementSyntax)(Nothing, pCount))
    Dim parameterNames = parameters.Select(Function(p) p.Identifier).ToList

    For i = 0 To ExistingGuards.Count - 1
      Dim eg = ExistingGuards(i)
      Dim pinfo = _Model_.GetTypeInfo(_parmeter_.AsClause.Type, cancellationToken)
      If pinfo.ConvertedType.IsReferenceType = False Then Continue For
      Dim Id = ExistingGuardParameters(i)
      Dim index = FindGuardIndex(parameterNames, Id)
      If index.HasValue Then Guards(index.Value) = eg.WithSameTriviaAs(eg)''.WithTrailingTrivia(SyntaxTrivia(SyntaxKind.EndOfLineTrivia, ""))
    Next

    Dim NewGuardIndex = FindGuardIndex(parameterNames, _parmeter_.Identifier)
    If NewGuardIndex.HasValue Then Guards(NewGuardIndex.Value) = NewGuard.WithLeadingTrivia(SyntaxTrivia(SyntaxKind.EndOfLineTrivia,"")).
                                                                          WithTrailingTrivia(SyntaxTrivia(SyntaxKind.EndOfLineTrivia, ""))
    Dim newmethod = method.RemoveNodes(ExistingGuards, SyntaxRemoveOptions.KeepExteriorTrivia  And  SyntaxRemoveOptions.KeepEndOfLine)

    Dim NonNullGuards = Guards.WhereNonNull.ToList
    Dim newStatements = newmethod.Statements.InsertRange(0, NonNullGuards)
    newmethod = newmethod.WithStatements(newStatements)
    Dim newBlock = newmethod.WithAdditionalAnnotations(Formatting.Formatter.Annotation)
    Return document.WithSyntaxRoot((Await document.GetSyntaxRootAsync(cancellationToken)).ReplaceNode(method, newBlock))
  End Function

  Private Shared Function FindGuardIndex(ParametersNames As List(Of ModifiedIdentifierSyntax), Id As IdentifierNameSyntax) As Integer?
    For j = 0 To ParametersNames.Count - 1
      If String.Compare(ParametersNames(j).Identifier.Text, Id.Identifier.Text, StringComparison.Ordinal) = 0 Then Return New Integer?(j)
    Next j
    Return New Integer?()
  End Function

  Private Shared Function FindGuardIndex(ParametersNames As List(Of ModifiedIdentifierSyntax), Id As ModifiedIdentifierSyntax) As Integer?
    For j = 0 To ParametersNames.Count - 1
      If String.Compare(ParametersNames(j).Identifier.Text, Id.Identifier.Text, StringComparison.Ordinal) = 0 Then Return New Integer?(j)
    Next j
    Return New Integer?()
  End Function



  Private Shared Function GetSingleIFstatement(parameterStmt As ParameterSyntax) As SingleLineIfStatementSyntax
    Return String.Format("If {0} Then {1}", GetIsNothingExpr(parameterStmt), GetThrowStatementForParameter(parameterStmt)).ToSExpr(Of SingleLineIfStatementSyntax)
  End Function

  Private Shared Function GetThrowStatementForParameter(parameterStmt As ParameterSyntax) As ThrowStatementSyntax
    Return String.Format(" Throw New System.ArgumentNullException({0})", GetParameterName(parameterStmt)).ToSExpr(Of ThrowStatementSyntax)
  End Function

  Private Shared Function GetIsNothingExpr(forParameter As ParameterSyntax) As BinaryExpressionSyntax
    Return String.Format(" {0} Is Nothing ", forParameter.Identifier.Identifier.Text).ToExpr(Of BinaryExpressionSyntax)
  End Function

  Private Shared Function GetMultiLineIFstatement(parameterStmt As ParameterSyntax) As MultiLineIfBlockSyntax
    Return String.Format("If {0} Then
  {1}
End If",
    GetIsNothingExpr(parameterStmt),
    GetThrowStatementForParameter(parameterStmt)).ToSExpr(Of MultiLineIfBlockSyntax)

  End Function

  Private Shared Function GetParameterName(parameterStmt As ParameterSyntax) As LiteralExpressionSyntax
    ' Note: If I can find the nameof feature in VB.net, then I'll change this line to reflect that
    Return StringLiteralExpression(Literal(parameterStmt.Identifier.Identifier.Text))
  End Function

End Class

Public Module Exts

  <Runtime.CompilerServices.Extension>
  Public Function WhereNonNull(Of T As Class)(xs As IEnumerable(Of T)) As IEnumerable(Of T)
    Return xs.Where(Function(x) x IsNot Nothing)
  End Function
End Module