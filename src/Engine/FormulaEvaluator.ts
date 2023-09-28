import Cell from "./Cell";
import SheetMemory from "./SheetMemory";
import { ErrorMessages } from "./GlobalDefinitions";

type TokenType = string;
type FormulaType = TokenType[];

export class FormulaEvaluator {
  private _sheetMemory: SheetMemory;
  private _errorMessage: string = "";
  private _result: number = 0;

  constructor(memory: SheetMemory) {
    this._sheetMemory = memory;
  }

  evaluate(formula: FormulaType): void {
    // Reset error message and result
    this._errorMessage = "";
    this._result = 0;

    if (formula.length === 0) {
      this._errorMessage = ErrorMessages.emptyFormula;
      return;
    }

    try {
      const result = this.evaluateExpression(formula);
      this._result = result;
    } catch (error) {
      this._errorMessage = ErrorMessages.invalidFormula;
    }
  }

  private evaluateExpression(tokens: FormulaType): number {
    const values: number[] = [];
    const operators: string[] = [];

    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i];

      if (this.isNumber(token)) {
        values.push(Number(token));
      } else if (this.isCellReference(token)) {
        const [value, error] = this.getCellValue(token);
        if (error) {
          this._errorMessage = error;
          throw new Error("Invalid cell reference");
        }
        values.push(value);
      } else if (token === "(") {
        operators.push(token);
      } else if (token === ")") {
        while (operators.length > 0 && operators[operators.length - 1] !== "(") {
          this.applyOperator(operators.pop()!, values);
        }
        operators.pop(); // Remove the "("
      } else if (this.isOperator(token)) {
        while (
          operators.length > 0 &&
          this.getOperatorPrecedence(token) <= this.getOperatorPrecedence(operators[operators.length - 1])
        ) {
          this.applyOperator(operators.pop()!, values);
        }
        operators.push(token);
      }
    }

    while (operators.length > 0) {
      this.applyOperator(operators.pop()!, values);
    }

    if (values.length !== 1) {
      this._errorMessage = ErrorMessages.invalidFormula;
      throw new Error("Invalid formula");
    }

    return values.pop() || 0;
  }

  private applyOperator(operator: string, values: number[]): void {
    const operand2 = values.pop()!;
    const operand1 = values.pop()!;
    let result: number;

    switch (operator) {
      case "+":
        result = operand1 + operand2;
        break;
      case "-":
        result = operand1 - operand2;
        break;
      case "*":
        result = operand1 * operand2;
        break;
      case "/":
        if (operand2 === 0) {
          this._errorMessage = ErrorMessages.divideByZero;
          throw new Error("Division by zero");
        }
        result = operand1 / operand2;
        break;
      default:
        this._errorMessage = ErrorMessages.invalidOperator;
        throw new Error("Invalid operator");
    }

    values.push(result);
  }

  private isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  private isCellReference(token: TokenType): boolean {
    return Cell.isValidCellLabel(token);
  }

  private getCellValue(token: TokenType): [number, string] {
    const cell = this._sheetMemory.getCellByLabel(token);
    const formula = cell.getFormula();
    const error = cell.getError();

    if (error && error !== ErrorMessages.emptyFormula) {
      return [0, error];
    }

    if (formula.length === 0) {
      return [0, ErrorMessages.invalidCell];
    }

    const value = cell.getValue();
    return [value, ""];
  }

  private isOperator(token: TokenType): boolean {
    return ["+", "-", "*", "/"].includes(token);
  }

  private getOperatorPrecedence(operator: string): number {
    switch (operator) {
      case "+":
      case "-":
        return 1;
      case "*":
      case "/":
        return 2;
      default:
        return 0;
    }
  }

  public get error(): string {
    return this._errorMessage;
  }

  public get result(): number {
    return this._result;
  }
}

export default FormulaEvaluator;
