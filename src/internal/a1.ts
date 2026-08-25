export function columnIndexToLetters(index: number): string {
  assertIndex(index, "columnIndex");
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

export function toA1Address(rowIndex: number, columnIndex: number): string {
  assertIndex(rowIndex, "rowIndex");
  return `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`;
}

export function fromA1Address(address: string): [number, number] {
  const match = /^([A-Za-z]+)([1-9]\d*)$/.exec(address.trim());
  if (!match?.[1] || !match[2]) {
    throw new TypeError(`Invalid A1 cell address: ${address}`);
  }
  let column = 0;
  for (const character of match[1].toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return [Number(match[2]) - 1, column - 1];
}

export function quoteSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

export function assertIndex(value: number, name: string): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative integer.`);
  }
}
