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
    Dim _Identifier_ = node.Try(Of ModifiedIdentifierSyntax) : If _Identifier_ Is Nothing Then Return
    Dim _parmeter_ = _Identifier_.Parent.Try(of ParameterSyntax) : If _parmeter_ Is Nothing Then Return
    Dim _method_ =  _parmeter_.Parent.Parent.Parent.Try(of MethodBlockSyntax) : If _method_ Is Nothing Then Return

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
               Dim singleIF = s.As(of SingleLineIfStatementSyntax)
               Dim isExpr = singleIF.Condition.Try(of BinaryExpressionSyntax)
               If (isExpr Is Nothing) OrElse (Not isExpr.IsKind(SyntaxKind.IsExpression)) Then Return False
               Dim res = CheckIfCondition(_parmeter_, isExpr)
               If Not res Then Return res
               Dim _sif_ = GetGuardStatement(_parmeter_).NormalizeWhitespace
               Return Not singleIF.IsEquivalentTo(_sif_)
             ElseIf TypeOf s Is MultiLineIfBlockSyntax Then
               Dim multiIF = s.As(Of MultiLineIfBlockSyntax)
               Dim isExpr = multiIF.IfStatement.Condition.Try(of BinaryExpressionSyntax)
               If (isExpr Is Nothing) OrElse (Not isExpr.IsKind(SyntaxKind.IsExpression)) Then Return False
               Dim res = CheckIfCondition(_parmeter_, isExpr)
               If Not res Then Return res
               Dim _mif_ = GetMultiLineGuardStatement(_parmeter_)
               Return Not multiIF.WithoutAnnotations().IsEquivalentTo(_mif_)
             End If
             Return False
           End Function
  End Function

  Private Shared Function CheckIfCondition(paramSyntax As ParameterSyntax, isExpr As BinaryExpressionSyntax) As Boolean
    Dim l = isExpr.Left.Try(of IdentifierNameSyntax)
    Dim r = isExpr.Right.Try(Of LiteralExpressionSyntax)
    If l Is Nothing OrElse r Is Nothing Then Return False
    If r.IsKind(SyntaxKind.NothingLiteralExpression) = False Then Return False
    Return String.Compare(l.Identifier.Text, paramSyntax.Identifier.Identifier.Text, StringComparison.Ordinal) = 0
  End Function

  Private Shared Function CheckIfCondition(isExpr As BinaryExpressionSyntax) As Boolean
    Dim l = isExpr.Left.Try(Of IdentifierNameSyntax)
    Dim r = isExpr.Right.Try(Of LiteralExpressionSyntax)
    If l Is Nothing OrElse r Is Nothing Then Return False
    Return r.IsKind(SyntaxKind.NothingLiteralExpression)
  End Function

  Private Async Function AddParameterNullCheckAsync(document As Document,
                                                    _parmeter_ As ParameterSyntax,
                                                    method As MethodBlockSyntax,
                                                    cancellationToken As CancellationToken) As Task(Of Document)

    Dim NewGuardStatement = GetGuardStatement(_parmeter_)

    Dim ifStatements = method.Statements.Where(Function(s) (TypeOf s Is MultiLineIfBlockSyntax) OrElse (TypeOf s Is SingleLineIfStatementSyntax))

    Dim ExistingGuards = ifStatements.Where(
      Function(s)
        If TypeOf s Is SingleLineIfStatementSyntax Then Return CheckIfCondition(s.As(Of SingleLineIfStatementSyntax).Condition.Try(Of BinaryExpressionSyntax))
        If TypeOf s Is MultiLineIfBlockSyntax Then Return CheckIfCondition(s.As(Of MultiLineIfBlockSyntax).IfStatement.Condition.Try(Of BinaryExpressionSyntax))
        Return False
      End Function)

    Dim ExistingGuardParameters = ExistingGuards.Select(
      Function(s)
        If TypeOf s Is SingleLineIfStatementSyntax Then Return s.As(Of SingleLineIfStatementSyntax).Condition.Try(Of BinaryExpressionSyntax)?.Left.As(Of IdentifierNameSyntax)
        If TypeOf s Is MultiLineIfBlockSyntax Then Return s.As(Of MultiLineIfBlockSyntax).IfStatement.Condition.Try(Of BinaryExpressionSyntax)?.Left.As(Of IdentifierNameSyntax)
        Return Nothing
      End Function)


    Dim parameters = method.Begin.ParameterList.Parameters
    Dim NumberOfParameters = parameters.Count


    Dim Guards = Enumerable.Repeat(Of StatementSyntax)(Nothing, NumberOfParameters).ToArray
    Dim parameterNames = parameters.Select(Function(p) p.Identifier).ToArray

    For i = 0 To ExistingGuards.Count - 1
      Dim eg = ExistingGuards(i)
      Dim Id = ExistingGuardParameters(i)
      Dim index = FindGuardIndex(parameterNames, Id)
      If index.HasValue Then Guards(index.Value) = eg '.WithSameTriviaAs(eg)
    Next

    Dim NewGuardIndex = FindGuardIndex(parameterNames, _parmeter_.Identifier)
    If NewGuardIndex.HasValue Then Guards(NewGuardIndex.Value) = NewGuardStatement.AddEOL

    ' Remove the existing guard statement
    Dim newmethod = method.RemoveNodes(ExistingGuards, SyntaxRemoveOptions.KeepNoTrivia)
    ' Get the Guards that will exist
    Dim NonNullGuards = Guards.WhereNonNull
    ' Insert them into the code
    Dim newStatements = newmethod.Statements.InsertRange(0, NonNullGuards)
    newmethod = newmethod.WithStatements(newStatements)
    Dim newBlock = newmethod.WithAdditionalAnnotations(Formatting.Formatter.Annotation)
    ' Get the new document with the applied changes
    Return document.WithSyntaxRoot((Await document.GetSyntaxRootAsync(cancellationToken)).ReplaceNode(method, newBlock))
  End Function

  Private Shared Function FindGuardIndex(ParametersNames() As ModifiedIdentifierSyntax, Id As IdentifierNameSyntax) As Integer?
    For j = 0 To ParametersNames.Count - 1
      If String.Compare(ParametersNames(j).Identifier.Text,
                        Id.Identifier.Text, StringComparison.Ordinal) = 0 Then Return New Integer?(j)
    Next j
    Return New Integer?()
  End Function

  Private Shared Function FindGuardIndex(ParametersNames() As ModifiedIdentifierSyntax, Id As ModifiedIdentifierSyntax) As Integer?
    For j = 0 To ParametersNames.Count - 1
      If String.Compare(ParametersNames(j).Identifier.Text,
                        Id.Identifier.Text, StringComparison.Ordinal) = 0 Then Return New Integer?(j)
    Next j
    Return New Integer?()
  End Function



  Private Shared Function GetGuardStatement(parameterStmt As ParameterSyntax) As SingleLineIfStatementSyntax
    Return String.Format("If {0} Then {1}", GetIsNothingExpr(parameterStmt), GetThrowStatementForParameter(parameterStmt)).ToSExpr(Of SingleLineIfStatementSyntax)
  End Function

  Private Shared Function GetThrowStatementForParameter(parameterStmt As ParameterSyntax) As ThrowStatementSyntax
    Return String.Format(" Throw New System.ArgumentNullException({0})", GetParameterName(parameterStmt)).ToSExpr(Of ThrowStatementSyntax)
  End Function

  Private Shared Function GetIsNothingExpr(forParameter As ParameterSyntax) As BinaryExpressionSyntax
    Return String.Format(" {0} Is Nothing ", forParameter.Identifier.Identifier.Text).ToExpr(Of BinaryExpressionSyntax)
  End Function

  Private Shared Function GetMultiLineGuardStatement(parameterStmt As ParameterSyntax) As MultiLineIfBlockSyntax
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
  Public Function [Try](Of T1 As Class)(expr As Object) As T1
    Return TryCast(expr, T1)
  End Function
  <Runtime.CompilerServices.Extension>
  Public Function [As](Of T1 As Class)(expr As Object) As T1
    Return DirectCast(expr, T1)
  End Function

  <Runtime.CompilerServices.Extension>
  Public Function WhereNonNull(Of T As Class)(xs As IEnumerable(Of T)) As IEnumerable(Of T)
    Return xs.Where(Function(x) x IsNot Nothing)
  End Function

  <Runtime.CompilerServices.Extension>
  Public Function OfType(Of xT, T1, T2)(xs As IEnumerable(Of xT)) As IEnumerable(Of xT)
    Return xs.Where(Function(x) (TypeOf xs Is T1) OrElse (TypeOf xs Is T2))
  End Function


  <Runtime.CompilerServices.Extension>
  Public Function AddEOL(Of T0 As SyntaxNode)(node As T0) As T0
    Return node.WithTrailingTrivia(SyntaxFactory.SyntaxTrivia(SyntaxKind.EndOfLineTrivia, Environment.NewLine))
  End Function
End Module