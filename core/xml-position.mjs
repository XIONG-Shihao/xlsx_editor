/**
 * Where a parsed XML element sits in the text it was read from.
 *
 * Several writers change one element of a part and leave every other byte as
 * it was, so they need the element's place in the original text. The XML reader
 * reports a line and a column; this turns them into an offset. Lines are
 * counted the way the reader counts them: a line ends at "\r\n", at a lone
 * "\r" (workbooks from older Mac software have these) or at "\n".
 */
export function lineStarts(source) {
  const starts = [0],
    ending = /\r\n?|\n/g;
  for (let match = ending.exec(source); match; match = ending.exec(source)) starts.push(match.index + match[0].length);
  return starts;
}

/** The offset in `source` at which the reader found `node`. Callers still check the text there. */
export function nodeStart(source, node, starts = lineStarts(source)) {
  return starts[node.lineNumber - 1] + node.columnNumber - 1;
}
