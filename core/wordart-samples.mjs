/** Repository-authored WordArt run/effect presets; actual text stays editable. */
const sample = (id, label, color, outline, shadow) =>
  Object.freeze({
    id,
    label,
    family: 'wordart',
    geometry: 'rect',
    fill: null,
    line: null,
    color,
    fontSize: 32,
    align: 'ctr',
    textBox: true,
    outline,
    shadow,
  });
export const WORDART_SAMPLES = Object.freeze([
  sample('wordart-blue', 'Blue Fill', '4472C4', null, false),
  sample('wordart-outline', 'Orange Outline', 'C65911', 'C65911', false),
  sample('wordart-shadow', 'Purple Shadow', '7030A0', null, true),
]);
export const WORDART_LIMITS = Object.freeze({ text: 256, paragraphs: 4 });
export function wordArtRunXml(sample) {
  const fill = (c) => (c ? '<a:solidFill><a:srgbClr val="' + c + '"/></a:solidFill>' : '<a:noFill/>'),
    line = '<a:ln w="12700">' + fill(sample.outline) + '</a:ln>',
    shadow = sample.shadow
      ? '<a:outerShdw blurRad="19050" dist="53882" dir="2700000" sx="100000" sy="100000" kx="0" ky="0" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>'
      : '';
  return (
    '<a:rPr lang="en-US" sz="3200" b="1" i="0">' +
    line +
    fill(sample.outline ? null : sample.color) +
    '<a:effectLst>' +
    shadow +
    '</a:effectLst><a:latin typeface="Arial"/></a:rPr>'
  );
}
