import { Formula } from './docxTypes';
import { ensureArray } from './utils/array';
import { extractTextValue } from './utils/xml';

/**
 * Extract mathematical formulas from document elements
 */
export function extractFormulas(elements: any[]): Formula[] {
  const formulas: Formula[] = [];

  for (const element of elements) {
    if (element.type === 'paragraph') {
      const paraFormulas = extractFormulasFromParagraph(element.data);
      formulas.push(...paraFormulas);
    }
  }

  return formulas;
}

/**
 * Extract formulas from a paragraph
 */
function extractFormulasFromParagraph(para: any): Formula[] {
  const formulas: Formula[] = [];

  // Check for oMathPara (paragraph-level math)
  const oMathPara = para['m:oMathPara'];
  if (oMathPara) {
    const formula = extractFormulaFromOMathPara(oMathPara);
    if (formula) {
      formulas.push(formula);
    }
  }

  // Check for inline oMath
  const oMath = para['m:oMath'];
  if (oMath) {
    const formula = extractFormulaFromOMath(oMath);
    if (formula) {
      formulas.push(formula);
    }
  }

  return formulas;
}

/**
 * Extract formula from oMathPara element
 */
function extractFormulaFromOMathPara(oMathPara: any): Formula | null {
  // Get the oMath element inside oMathPara
  const oMath = oMathPara['m:oMath'];
  if (!oMath) return null;

  return extractFormulaFromOMath(oMath);
}

/**
 * Extract formula from oMath element
 */
function extractFormulaFromOMath(oMath: any): Formula | null {
  // Serialize the OMML back to XML string
  const omml = serializeOMML(oMath);

  // Extract plain text representation
  const text = extractTextFromOMML(oMath);

  return {
    omml,
    text: text || undefined
  };
}

/**
 * Serialize OMML object back to XML string
 */
function serializeOMML(oMath: any): string {
  // Simple serialization - just keep the OMML structure as a string representation
  // For production, you might want to use a proper XML serializer
  return JSON.stringify(oMath, null, 2);
}

// Maximum recursion depth for formula extraction to prevent stack overflow
const MAX_FORMULA_RECURSION_DEPTH = 50;

/**
 * Extract plain text from OMML recursively
 */
function extractTextFromOMML(obj: any, depth: number = 0): string {
  if (!obj || typeof obj !== 'object') return '';

  // Prevent stack overflow from deeply nested or circular formula structures
  if (depth > MAX_FORMULA_RECURSION_DEPTH) {
    console.warn(`Maximum formula recursion depth (${MAX_FORMULA_RECURSION_DEPTH}) reached, truncating extraction`);
    return '';
  }

  const textParts: string[] = [];

  // Look for m:t elements (math text)
  if (obj['m:t']) {
    const text = extractTextValue(obj['m:t']);
    if (text) textParts.push(text);
  }

  // Recursively search all properties with depth tracking
  for (const key in obj) {
    if (obj.hasOwnProperty(key) && key !== 'm:t') {
      const value = obj[key];
      if (Array.isArray(value)) {
        value.forEach(item => {
          const childText = extractTextFromOMML(item, depth + 1);
          if (childText) textParts.push(childText);
        });
      } else if (typeof value === 'object') {
        const childText = extractTextFromOMML(value, depth + 1);
        if (childText) textParts.push(childText);
      }
    }
  }

  return textParts.join(' ');
}
