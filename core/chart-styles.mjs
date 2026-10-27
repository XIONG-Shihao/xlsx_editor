/** Three explicit series-format presets; no Office catalog or theme inference. */
import { refusal } from './refusal-codes.mjs';

const colorful = Object.freeze(['2563EB', 'E86125', '16A34A', '9333EA', 'CA8A04', '0891B2', 'DC2626', '64748B']);
const mono = Object.freeze(['217346', '42875F', '66A27F', '85B998', '175431', '327149', '50946A', '749F84']);
export const CHART_STYLES = Object.freeze(
  [
    { id: 'colorful', label: 'Colorful', colors: colorful, outline: false },
    { id: 'monochrome', label: 'Monochrome', colors: mono, outline: false },
    { id: 'outline', label: 'Outline', colors: colorful, outline: true },
  ].map(Object.freeze)
);
export function chartStyle(id) {
  const style = CHART_STYLES.find((s) => s.id === id);
  if (!style) throw refusal('chart-bad-input', 'Choose Colorful, Monochrome or Outline');
  return style;
}
export function chartSeriesAppearance(id, index) {
  if (!Number.isSafeInteger(index) || index < 0 || index > 7)
    throw refusal('chart-limit', 'Chart series index exceeds its bound');
  const style = id === undefined ? null : chartStyle(id),
    color = '#' + (style?.colors ?? colorful)[index];
  return {
    color,
    fill: style?.outline ? 'none' : color,
    stroke: style ? color : undefined,
    width: style ? 2 : undefined,
    dash: style?.outline ? '6 3' : undefined,
  };
}
export function chartSeriesPropertiesXml(id, index) {
  if (id === undefined) return '';
  const style = chartStyle(id),
    color = style.colors[index];
  if (!color) throw refusal('chart-limit', 'Chart series index exceeds its bound');
  return `<c:spPr>${style.outline ? '<a:noFill/>' : `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`}<a:ln w="19050"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:prstDash val="${style.outline ? 'dash' : 'solid'}"/></a:ln></c:spPr>`;
}
