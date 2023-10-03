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

    // Validate initial formula conditions
    if (formula.length === 0) {
      this._errorMessage = ErrorMessages.emptyFormula;
      return;
    }

    // Evaluate the formula expression when there's only parenthesis
    if (formula.length === 2 && formula[0] === "(" && formula[1] === ")") {
      this._errorMessage = ErrorMessages.missingParentheses;
      return;
    }

    // Check for formula with the last character being an operator
    const lastCharacter = formula[formula.length - 1];
    if (this.isOperator(lastCharacter) && lastCharacter !== ")") {
      this._errorMessage = ErrorMessages.invalidFormula;
    }

    // Evaluate the formula expression
    try {
      const result = this.evaluateExpression(formula);
      this._result = result;
    } catch (error) {
      if (!this._errorMessage) {
        this._errorMessage = ErrorMessages.invalidFormula;
      }
    }
  }

  // Evaluate the formula expression
  private evaluateExpression(tokens: FormulaType): number {
    const values: number[] = [];
    const operators: string[] = [];

    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i];

      if (this.isNumber(token)) {
        values.push(Number(token));
        // Check if the number is followed by an open parenthesis without an operator in between
        if (i < tokens.length - 1 && tokens[i + 1] === "(") {
          this._errorMessage = ErrorMessages.invalidFormula;
          return Number(token);
        }
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
        while (
          operators.length > 0 &&
          operators[operators.length - 1] !== "("
        ) {
          this.applyOperator(operators.pop()!, values);
        }
        operators.pop(); // Remove the "("
      } else if (this.isOperator(token)) {
        if (i < tokens.length - 1 && this.isOperator(tokens[i + 1])) {
          this._errorMessage = ErrorMessages.invalidFormula;
          return values[values.length - 1] || 0; // Return the last value
        }
        while (
          operators.length > 0 &&
          this.getOperatorPrecedence(token) <=
            this.getOperatorPrecedence(operators[operators.length - 1])
        ) {
          this.applyOperator(operators.pop()!, values);
        }
        operators.push(token);
      }
    }

    // Check for trailing operators
    const lastToken = tokens[tokens.length - 1];
    if (this.isOperator(lastToken)) {
      this._errorMessage = ErrorMessages.invalidFormula;
      return values[values.length - 1] || 0; // Return the last value
    }

    // Check if the formula is incomplete (due to unmatched parentheses, etc.)
    if (
      operators.length > 0 &&
      (!this.isOperator(operators[operators.length - 1]) || values.length < 2)
    ) {
      this._errorMessage = ErrorMessages.invalidFormula;
      throw new Error("Invalid formula");
    }

    // Remaining operations
    while (operators.length > 0) {
      this.applyOperator(operators.pop()!, values);
    }

    return values.pop() || 0;
  }

  // Apply the operator to the operands
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
          this._result = Infinity;
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

  // Check if the token is a number
  private isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  // Check if the token is a cell reference
  private isCellReference(token: TokenType): boolean {
    return Cell.isValidCellLabel(token);
  }

  // Get the value of the cell
  private getCellValue(token: TokenType): [number, string] {
    const cell = this._sheetMemory.getCellByLabel(token);
    const formula = cell.getFormula();
    const error = cell.getError();

    // Check for circular reference
    if (error && error !== ErrorMessages.emptyFormula) {
      return [0, error];
    }

    // Check for empty formula
    if (formula.length === 0) {
      return [0, ErrorMessages.invalidCell];
    }

    // Evaluate the formula
    const value = cell.getValue();
    return [value, ""];
  }

  // Check if the token is an operator
  private isOperator(token: TokenType): boolean {
    return ["+", "-", "*", "/"].includes(token);
  }

  // Get the precedence of the operator
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
