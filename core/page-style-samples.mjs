import { FORMAT, formatTemplate } from './format.mjs';
/** Small application-owned samples, not the Office asset catalog. Fonts are
 * family references; this module embeds no font files or imported assets. */
import { refusal } from './refusal-codes.mjs';

const palettes = {
  neutral: [
    '17212B',
    'FFFFFF',
    '364152',
    'F3F5F7',
    '3366CC',
    'E87924',
    '27864B',
    'C83D45',
    '8053A5',
    '228B9A',
    '0057B8',
    '70458A',
  ],
  forest: [
    '182A20',
    'FFFFFF',
    '355642',
    'EDF5EF',
    '28744B',
    'A17828',
    '4879A1',
    '985552',
    '718145',
    '4D8E87',
    '226748',
    '806146',
  ],
  plum: [
    '2D2034',
    'FFFFFF',
    '5E4369',
    'F6EFF8',
    '86509C',
    'D17149',
    '4D8191',
    'A84F70',
    '7B8D49',
    'C59A39',
    '74448B',
    '935E78',
  ],
};
const xmlOrder = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];
export const pageThemeSamples = Object.freeze(
  [
    { id: 'neutral', label: 'Neutral', majorFont: 'Arial', minorFont: 'Arial' },
    { id: 'forest', label: 'Forest', majorFont: 'Georgia', minorFont: 'Verdana' },
    { id: 'plum', label: 'Plum', majorFont: 'Verdana', minorFont: 'Georgia' },
  ].map((sample) =>
    Object.freeze({
      ...sample,
      colors: Object.freeze(Object.fromEntries(xmlOrder.map((key, index) => [key, palettes[sample.id][index]]))),
    })
  )
);
export function pageThemeSample(id) {
  const sample = pageThemeSamples.find((item) => item.id === id);
  if (!sample) throw refusal('internal-unknown-reference', 'Unknown page theme sample');
  return sample;
}
const font = (tag, family) =>
  `<a:${tag}><a:latin typeface="${family}"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`;
const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
/** Closed catalog lookup keeps XML interpolation separate from user text. */
export function pageThemeXml(id) {
  const sample = pageThemeSample(id);
  return formatTemplate(
    [
      FORMAT.XML_00FFDBFDB5,
      FORMAT.XML_FB0728D55F,
      FORMAT.XML_AB2CD77488,
      FORMAT.XML_343AB1FB21,
      FORMAT.XML_AB2CD77488,
      FORMAT.FORMAT_E3B0C4,
      FORMAT.XML_88C522409D,
      FORMAT.XML_5CECD6BA0D,
      FORMAT.XML_5E79C936B2,
      FORMAT.XML_527B68AF2A,
      FORMAT.XML_C5B39A6AD6,
    ],
    [
      sample.label,
      sample.label,
      xmlOrder.map((key) => `<a:${key}><a:srgbClr val="${sample.colors[key]}"/></a:${key}>`).join(''),
      sample.label,
      font('majorFont', sample.majorFont),
      font('minorFont', sample.minorFont),
      fill.repeat(3),
      [6350, 12700, 19050]
        .map((width) => `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr">${fill}<a:prstDash val="solid"/></a:ln>`)
        .join(''),
      '<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3),
      fill.repeat(3),
    ]
  );
}

// Canonical 32×32, 8-bit RGB, noninterlaced tiles. Generated procedurally from
// the patterns named below; SHA identities are verified by the qualification.
export const worksheetBackgroundSamples = Object.freeze(
  [
    {
      id: 'graph-paper',
      label: 'Graph paper',
      width: 32,
      height: 32,
      mime: 'image/png',
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAAN0lEQVR42mN48OojTREDEP/68w8TUUt81IJRC6hjwWg+GLVg1ILRfDBqwagFo/lg1IJRC/78AwD+Nps27drLrwAAAABJRU5ErkJggg==',
      sha256: '2a170d83b9240a532a7b55fee1ba602ef49a980242c8c75f38a9a54e6868ccdf',
    },
    {
      id: 'soft-dots',
      label: 'Soft dots',
      width: 32,
      height: 32,
      mime: 'image/png',
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAANklEQVR42mP48+sfTRHDMLLg0YPXWFVQKM4wGgejFozmg9E4GLVgNB+MxsGoBaP5YDQOBokFAG+bzDa3/bn2AAAAAElFTkSuQmCC',
      sha256: 'a9aaa04469a39dba041129e95381251558d2f9862439b4424b242faffffcfe49',
    },
  ].map((sample) => Object.freeze(sample))
);
export function worksheetBackgroundSample(id) {
  const sample = worksheetBackgroundSamples.find((item) => item.id === id);
  if (!sample) throw refusal('internal-unknown-reference', 'Unknown worksheet background sample');
  return sample;
}
