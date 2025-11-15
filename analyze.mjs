import { stripComments, isInsideQuotes } from "./utils/helpers.js";
import fs from "fs";

const formulaRaw = fs.readFileSync("temp_formula.txt", "utf8");
const stripped = stripComments(formulaRaw) || "";
let formulaText = stripped.trim();
if (!formulaText.startsWith("=")) {
  formulaText = `=${formulaText}`;
}

const periodPattern = /(\w+|"[^"]*"|[A-Z]+\$?\d+\$?)\s*\./g;
let match;
const results = [];
while ((match = periodPattern.exec(formulaText)) !== null) {
  const tokenStartIndex = match.index;
  const dotIndex = match.index + match[0].length - 1;
  const skip = isInsideQuotes(formulaText, tokenStartIndex) || isInsideQuotes(formulaText, dotIndex);
  results.push({ match: match[0], tokenStartIndex, dotIndex, skip });
}
console.log(JSON.stringify(results, null, 2));
