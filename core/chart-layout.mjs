/** Shared finite geometry for the qualified column/bar/line grid and PDF views. */
export function chartPreviewGeometry(object) {
  const { categories, series, hidden } = object.data,
    kind = object.chart.kind,
    w = 600,
    h = 350,
    left = kind === 'bar' ? 110 : 52,
    right = 18,
    top = 35,
    bottom = 55,
    pw = w - left - right,
    ph = h - top - bottom;
  const visible = categories.map((_, i) => i).filter((i) => !hidden[i]),
    values = series.flatMap((s) => visible.flatMap((i) => (s.values[i] === null ? [] : [s.values[i]]))),
    min = Math.min(0, ...values),
    max = Math.max(0, ...values),
    normalizer = Math.max(1, Math.abs(min), Math.abs(max)),
    span = max / normalizer - min / normalizer || 1;
  const x = (value) => left + ((value / normalizer - min / normalizer) / span) * pw,
    y = (value) => top + ((max / normalizer - value / normalizer) / span) * ph;
  const slot = (kind === 'bar' ? ph : pw) / Math.max(1, visible.length),
    bars = [],
    segments = [],
    labels = [],
    lines =
      kind === 'bar'
        ? [
            [x(0), top, x(0), h - bottom],
            [left, h - bottom, w - right, h - bottom],
          ]
        : [
            [left, top, left, h - bottom],
            [left, y(0), w - right, y(0)],
          ];
  labels.push(
    ...(kind === 'bar'
      ? [
          { x: left, y: h - bottom + 18, anchor: 'middle', text: String(min) },
          { x: w - right, y: h - bottom + 18, anchor: 'middle', text: String(max) },
        ]
      : [
          { x: left - 5, y: top + 4, anchor: 'end', text: String(max) },
          { x: left - 5, y: h - bottom, anchor: 'end', text: String(min) },
        ])
  );
  series.forEach((item, s) => {
    if (kind === 'bar' || kind === 'column')
      visible.forEach((i, order) => {
        const value = item.values[i];
        if (value === null) return;
        const breadth = Math.max(0.2, (slot * 0.8) / series.length);
        const box =
          kind === 'bar'
            ? {
                x: Math.min(x(value), x(0)),
                y: h - bottom - (order + 1) * slot + slot * 0.1 + s * breadth,
                width: Math.max(0.5, Math.abs(x(value) - x(0))),
                height: breadth,
              }
            : {
                x: left + order * slot + slot * 0.1 + s * breadth,
                y: Math.min(y(value), y(0)),
                width: breadth,
                height: Math.max(0.5, Math.abs(y(value) - y(0))),
              };
        bars.push({
          ...box,
          series: s,
          title: `${categories[i]}: ${item.name} ${Object.is(value, -0) ? '-0' : value}`,
        });
      });
    else {
      let points = [];
      const emit = () => {
        if (points.length) segments.push({ series: s, points });
        points = [];
      };
      visible.forEach((i, order) => {
        const value = item.values[i];
        if (value === null) emit();
        else points.push([left + (order + 0.5) * slot, y(value)]);
      });
      emit();
    }
  });
  visible.forEach((i, order) => {
    if (visible.length <= 12 || order === 0 || order === visible.length - 1)
      labels.push(
        kind === 'bar'
          ? { x: left - 5, y: h - bottom - (order + 0.5) * slot + 4, anchor: 'end', text: categories[i].slice(0, 16) }
          : { x: left + (order + 0.5) * slot, y: h - bottom + 18, anchor: 'middle', text: categories[i].slice(0, 24) }
      );
  });
  return { width: w, height: h, left, plotWidth: pw, lines, bars, segments, labels };
}
