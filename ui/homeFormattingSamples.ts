/** Finite Home samples. Fonts stay local: no embedded font assets or installation claim. */
export const homeFontSamples = Object.freeze(['Arial', 'Georgia', 'Verdana'] as const);
/** Each preset changes these fields only; values, number formats and alignment survive. */
export const homeStyleSamples = Object.freeze({
  good: Object.freeze({ bold: true, fontColor: '006100', fillColor: 'C6EFCE', border: 'none' as const }),
  bad: Object.freeze({ bold: true, fontColor: '9C0006', fillColor: 'FFC7CE', border: 'none' as const }),
  heading: Object.freeze({ bold: true, fontColor: 'FFFFFF', fillColor: '244062', border: 'none' as const }),
});
