import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import {
  pageThemeSamples,
  worksheetBackgroundSamples,
  pageThemeSample,
  worksheetBackgroundSample,
} from '../core/page-style-samples.mjs';
import { createPageStylesController, type PageStylesFrame } from './pageStylesController';
export function WorkbookPageStyles({
  frame,
  mode,
  words = (s: string) => s,
}: {
  frame: PageStylesFrame | null;
  mode: 'themes' | 'backgrounds';
  words?: (s: string) => string;
}) {
  const [controller] = useState(createPageStylesController),
    lifetime = useRef(0),
    snapshot = useSyncExternalStore(controller.subscribe, controller.getSnapshot);
  useLayoutEffect(() => controller.setFrame(frame));
  useEffect(() => {
    const serial = ++lifetime.current;
    return () => {
      queueMicrotask(() => {
        if (lifetime.current === serial) controller.dispose();
      });
    };
  }, [controller]);
  const [error, setError] = useState({ key: '', text: '' }),
    current = snapshot.key === frame?.key,
    disabled = !frame?.active || !frame.writable || !current || snapshot.pending;
  const theme = mode === 'themes' && snapshot.theme ? pageThemeSample(snapshot.theme) : null,
    background = mode === 'backgrounds' && snapshot.background ? worksheetBackgroundSample(snapshot.background) : null;
  const protection =
    mode === 'themes'
      ? !frame?.theme.canApply
        ? (frame?.theme.protectedReason ?? 'This theme is protected.')
        : undefined
      : frame?.background && !frame.background.canEdit
        ? (frame.background.protectedReason ?? 'This background is protected.')
        : undefined;
  const choose = (id: string) => {
    if (frame)
      try {
        controller.choose(frame, mode, id);
        setError({ key: '', text: '' });
      } catch (e) {
        setError({ key: frame.key, text: String(e instanceof Error ? e.message : e) });
      }
  };
  const run = async (action: () => Promise<unknown>) => {
    const key = frame?.key ?? '';
    setError({ key: '', text: '' });
    try {
      await action();
    } catch (e) {
      setError({ key, text: String(e instanceof Error ? e.message : e) });
    }
  };
  return (
    <section aria-label={words(mode === 'themes' ? 'Theme samples' : 'Worksheet background samples')}>
      <p>
        {words(
          mode === 'themes' ? 'Applies to the workbook. Literal cell fonts and colors are retained.' : 'Applies to'
        )}
        {mode === 'backgrounds' ? ' ' + frame?.sheet : ''}
      </p>
      <p>{words('Choose a sample to preview. Apply changes the workbook; Reset preview leaves it unchanged.')}</p>
      <div
        className="sheet-page-style-gallery"
        role="group"
        aria-label={words(mode === 'themes' ? 'Choose theme sample' : 'Choose background sample')}
      >
        {(mode === 'themes' ? pageThemeSamples : worksheetBackgroundSamples).map((sample) => (
          <button
            type="button"
            key={sample.id}
            aria-pressed={(mode === 'themes' ? snapshot.theme : snapshot.background) === sample.id}
            disabled={!frame?.active || snapshot.pending}
            onClick={() => choose(sample.id)}
          >
            {mode === 'backgrounds' && (
              <span
                aria-hidden="true"
                style={{
                  display: 'inline-block',
                  width: 32,
                  height: 32,
                  verticalAlign: 'middle',
                  marginRight: 8,
                  backgroundImage: `url(data:image/png;base64,${worksheetBackgroundSample(sample.id).base64})`,
                }}
              />
            )}
            {words(sample.label)}
          </button>
        ))}
      </div>
      {theme && (
        <div
          role="img"
          aria-label={words('Preview theme') + ' ' + words(theme.label)}
          style={{
            margin: '12px 0',
            padding: 12,
            border: '1px solid #bbb',
            background: '#' + theme.colors.lt1,
            color: '#' + theme.colors.dk1,
          }}
        >
          <strong style={{ fontFamily: theme.majorFont, fontSize: 20 }}>{words('Heading sample')}</strong>
          <p style={{ fontFamily: theme.minorFont }}>{words('Body text sample')} — 123.45</p>
          <div style={{ display: 'flex', gap: 4 }}>
            {['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((role) => (
              <span
                key={role}
                title={role + ' #' + theme.colors[role]}
                style={{ background: '#' + theme.colors[role], width: 24, height: 24, border: '1px solid #888' }}
              />
            ))}
          </div>
          <p>
            {theme.majorFont} / {theme.minorFont}
          </p>
        </div>
      )}
      {background && (
        <div
          role="img"
          aria-label={words('Preview background') + ' ' + words(background.label)}
          style={{
            margin: '12px 0',
            padding: 18,
            border: '1px solid #bbb',
            backgroundImage: `url(data:image/png;base64,${background.base64})`,
            backgroundRepeat: 'repeat',
          }}
        >
          {words('Worksheet sample')}
          <br />
          A1 · 123.45
        </div>
      )}
      {mode === 'themes' ? (
        <button
          type="button"
          disabled={disabled || !!protection || !theme}
          onClick={() => {
            if (frame) void run(() => controller.theme(frame));
          }}
        >
          {words('Apply theme')}
        </button>
      ) : (
        <>
          <button
            type="button"
            disabled={disabled || !!protection || !background}
            onClick={() => {
              if (frame) void run(() => controller.background(frame));
            }}
          >
            {words('Apply background')}
          </button>
          <button
            type="button"
            disabled={disabled || !!protection || !frame?.background?.sample}
            onClick={() => {
              if (frame) void run(() => controller.background(frame, true));
            }}
          >
            {words('Remove background')}
          </button>
        </>
      )}
      <button
        type="button"
        disabled={!frame?.active || snapshot.pending}
        onClick={() => {
          if (frame) controller.reset(frame);
        }}
      >
        {words('Reset preview')}
      </button>
      <p>
        {words(
          mode === 'themes'
            ? 'Three color/font samples. Fonts use installed faces or the disclosed worksheet substitution. These themes support printing. Imported and tinted theme formats stay protected.'
            : 'Two repeating background samples. Cell fills take precedence. Worksheet backgrounds do not print.'
        )}
      </p>
      {(frame?.reason || protection) && <p role="status">{words(frame?.reason ?? protection!)}</p>}
      {snapshot.pending && <p role="status">{words('Applying page style…')}</p>}
      {(snapshot.error || (error.key === frame?.key ? error.text : '')) && (
        <p role="alert">{words(snapshot.error || (error.key === frame?.key ? error.text : ''))}</p>
      )}
      {snapshot.notice && <p role="status">{words(snapshot.notice)}</p>}
    </section>
  );
}
