import { INK_PENS, type InkPenId } from '../core/ink-stroke.mjs';
export function WorkbookInkPalette({
  selected,
  disabled,
  reason,
  warning,
  onChoose,
  words = (s) => s,
}: {
  selected: InkPenId;
  disabled: boolean;
  reason?: string;
  warning?: string | null;
  onChoose: (pen: InkPenId) => void;
  words?: (s: string) => string;
}): JSX.Element {
  return (
    <div className="sheet-ink-palette" role="group" aria-label={words('Pen presets')}>
      <div className="sheet-ink-palette-options">
        {INK_PENS.map((pen) => (
          <button
            key={pen.id}
            type="button"
            data-ribbon-command="draw.preset"
            data-ink-preset={pen.id}
            aria-label={words(pen.label)}
            aria-pressed={selected === pen.id}
            disabled={disabled}
            title={words(
              reason ??
                pen.label +
                  ' — ' +
                  pen.widthEmu / 9525 +
                  ' px' +
                  (pen.alpha < 100000 ? ' · 30% opacity' : '') +
                  '; remembered privately in this app profile'
            )}
            onClick={() => onChoose(pen.id)}
          >
            <svg width="28" height="34" viewBox="0 0 28 34" aria-hidden="true">
              <path d="M8 2h12v17l-6 11-6-11Z" fill={'#' + pen.color} stroke="#444" />
              <path d="M8 18h12M14 30v2" stroke="#444" />
              <path
                d="M3 32h22"
                stroke={'#' + pen.color}
                strokeWidth={pen.id === 'yellow-highlighter' ? 4 : 2}
                strokeOpacity={pen.alpha / 100000}
              />
            </svg>
            <span>{words(pen.label)}</span>
            <small>
              {pen.widthEmu / 9525} px{pen.alpha < 100000 ? ' · 30%' : ''}
            </small>
          </button>
        ))}
      </div>
      {warning && (
        <span className="sheet-ink-preference-warning" role="status">
          {words(warning)}
        </span>
      )}
    </div>
  );
}
