// =============================================
// 셀 참조 유틸리티
// Excel 스타일 열 문자 <-> 인덱스 변환
// =============================================

/**
 * 열 인덱스(0-based)를 Excel 열 문자로 변환
 * 예: 0 -> "A", 25 -> "Z", 26 -> "AA"
 * @param {number} colIndex - 0-based 열 인덱스
 * @returns {string} 열 문자 (예: "A", "AB")
 */
export function colToLetter(colIndex) {
  let result = '';
  let index = colIndex;
  while (index >= 0) {
    result = String.fromCharCode((index % 26) + 65) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
}

/**
 * Excel 열 문자를 0-based 인덱스로 변환
 * 예: "A" -> 0, "Z" -> 25, "AA" -> 26
 * @param {string} letter - 열 문자 (예: "A", "AB")
 * @returns {number} 0-based 열 인덱스
 */
export function letterToCol(letter) {
  const upper = letter.toUpperCase();
  let result = 0;
  for (let i = 0; i < upper.length; i++) {
    result = result * 26 + (upper.charCodeAt(i) - 64);
  }
  return result - 1;
}

/**
 * 행/열 인덱스를 Excel 셀 참조 문자열로 변환
 * 예: (0, 0) -> "A1", (8, 4) -> "E9"
 * @param {number} row - 0-based 행 인덱스
 * @param {number} col - 0-based 열 인덱스
 * @returns {string} 셀 참조 (예: "E9")
 */
export function toCellRef(row, col) {
  return `${colToLetter(col)}${row + 1}`;
}

/**
 * Excel 셀 참조 문자열을 행/열 인덱스로 파싱
 * 예: "E9" -> {row: 8, col: 4}, "AA1" -> {row: 0, col: 26}
 * @param {string} ref - 셀 참조 문자열 (예: "E9")
 * @returns {{row: number, col: number}} 0-based 행/열 인덱스
 */
export function parseCellRef(ref) {
  const match = ref.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return { row: 0, col: 0 };
  const col = letterToCol(match[1]);
  const row = parseInt(match[2], 10) - 1;
  return { row, col };
}
