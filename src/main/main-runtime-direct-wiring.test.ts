import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import ts from 'typescript';

test('main process call sites bypass zero-logic pass-through wrappers', () => {
  const sourcePath = path.join(process.cwd(), 'src/main.ts');
  const source = fs.readFileSync(sourcePath, 'utf8');
  const sourceFile = ts.createSourceFile(
    sourcePath,
    source,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  const offenders: string[] = [];

  for (const statement of sourceFile.statements) {
    if (ts.isFunctionDeclaration(statement) && statement.name && statement.body) {
      if (isRuntimePassthrough(statement, statement.body)) {
        offenders.push(statement.name.text);
      }
      continue;
    }

    if (!ts.isVariableStatement(statement)) continue;
    for (const declaration of statement.declarationList.declarations) {
      if (
        ts.isIdentifier(declaration.name) &&
        declaration.initializer &&
        (ts.isArrowFunction(declaration.initializer) ||
          ts.isFunctionExpression(declaration.initializer)) &&
        isRuntimePassthrough(declaration.initializer, declaration.initializer.body)
      ) {
        offenders.push(declaration.name.text);
      }
    }
  }

  assert.deepEqual(offenders, []);
});

function isRuntimePassthrough(
  callable: ts.FunctionDeclaration | ts.ArrowFunction | ts.FunctionExpression,
  body: ts.ConciseBody,
): boolean {
  let expression: ts.Expression | undefined;
  if (!ts.isBlock(body)) {
    expression = body;
  } else {
    if (body.statements.length !== 1) return false;
    const statement = body.statements[0]!;
    if (ts.isExpressionStatement(statement)) expression = statement.expression;
    if (ts.isReturnStatement(statement)) expression = statement.expression;
  }
  if (!expression) return false;
  if (ts.isAwaitExpression(expression) || ts.isVoidExpression(expression)) {
    expression = expression.expression;
  }
  if (!ts.isCallExpression(expression)) return false;

  const root = getCalleeRoot(expression.expression);
  if (!root) return false;
  if (callable.parameters.length !== expression.arguments.length) return false;

  return callable.parameters.every((parameter, index) => {
    if (!ts.isIdentifier(parameter.name)) return false;
    const argument = expression.arguments[index];
    if (!argument) return false;
    if (parameter.dotDotDotToken) {
      if (!ts.isSpreadElement(argument)) return false;
      const spreadExpression = unwrapTransparentExpression(argument.expression);
      return ts.isIdentifier(spreadExpression) && spreadExpression.text === parameter.name.text;
    }
    const argumentExpression = unwrapTransparentExpression(argument);
    return ts.isIdentifier(argumentExpression) && argumentExpression.text === parameter.name.text;
  });
}

function unwrapTransparentExpression(expression: ts.Expression): ts.Expression {
  let current = expression;
  while (
    ts.isParenthesizedExpression(current) ||
    ts.isAsExpression(current) ||
    ts.isNonNullExpression(current)
  ) {
    current = current.expression;
  }
  return current;
}

function getCalleeRoot(expression: ts.LeftHandSideExpression): string {
  let current: ts.Expression = expression;
  while (ts.isPropertyAccessExpression(current) || ts.isCallExpression(current)) {
    current = current.expression;
  }
  return ts.isIdentifier(current) ? current.text : '';
}
