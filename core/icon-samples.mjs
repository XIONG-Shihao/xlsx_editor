/** Three repository-authored vector assets. Saved models retain their closed
 * custom path and editable sRGB color; no bitmap or imported SVG execution. */
export const ICON_COLOR = '217346';
const icon = (id, label, points) =>
  Object.freeze({
    id,
    label,
    family: 'icon',
    geometry: 'custom',
    fill: ICON_COLOR,
    line: null,
    color: ICON_COLOR,
    fontSize: 14,
    align: 'ctr',
    textBox: false,
    asset: 'core/assets/icons/' + id + '.svg',
    points: Object.freeze(points.map((p) => Object.freeze(p))),
  });
export const ICON_SAMPLES = Object.freeze([
  icon('icon-check', 'Check', [
    [12, 50],
    [35, 73],
    [88, 20],
    [96, 28],
    [35, 89],
    [4, 58],
  ]),
  icon('icon-star', 'Star', [
    [50, 4],
    [61, 36],
    [96, 36],
    [68, 57],
    [79, 91],
    [50, 70],
    [21, 91],
    [32, 57],
    [4, 36],
    [39, 36],
  ]),
  icon('icon-plus', 'Plus', [
    [40, 5],
    [60, 5],
    [60, 40],
    [95, 40],
    [95, 60],
    [60, 60],
    [60, 95],
    [40, 95],
    [40, 60],
    [5, 60],
    [5, 40],
    [40, 40],
  ]),
]);
export function iconGeometryXml(sample) {
  const path = sample.points
    .map(
      ([x, y], i) =>
        '<a:' + (i ? 'lnTo' : 'moveTo') + '><a:pt x="' + x + '" y="' + y + '"/></a:' + (i ? 'lnTo' : 'moveTo') + '>'
    )
    .join('');
  return (
    '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="100" b="100"/><a:pathLst><a:path w="100" h="100" stroke="0" extrusionOk="0">' +
    path +
    '<a:close/></a:path></a:pathLst></a:custGeom>'
  );
}
