import type { JSX } from 'react';
const paths: Record<string, JSX.Element> = {
  chevron: <path d="M6 9l6 6 6-6" />,
  font: (
    <>
      <path d="m4 19 6-14 6 14M6 15h8M18 11h4m-2 0v8" />
    </>
  ),
  fontIncrease: (
    <>
      <path d="m3 19 5-13 5 13M5 15h6M18 4v7m-3-4h6" />
    </>
  ),
  fontDecrease: (
    <>
      <path d="m3 19 5-13 5 13M5 15h6M15 7h6" />
    </>
  ),
  number: (
    <>
      <path d="m9 4-3 16M17 4l-3 16M4 9h17M3 15h17" />
    </>
  ),
  decimalDecrease: (
    <>
      <circle cx="4" cy="8" r=".7" fill="currentColor" />
      <rect x="8" y="3" width="5" height="8" rx="2" />
      <path d="M5 18h14m-4-3 4 3-4 3" />
    </>
  ),
  decimalIncrease: (
    <>
      <circle cx="3" cy="8" r=".7" fill="currentColor" />
      <rect x="7" y="3" width="5" height="8" rx="2" />
      <rect x="16" y="3" width="5" height="8" rx="2" />
      <path d="M5 18h14m-10-3-4 3 4 3" />
    </>
  ),
  insertCells: (
    <>
      <path d="M14 3H3v18h18V10M3 9h9M3 15h18M9 9v12" />
      <path d="M18 2v8m-4-4h8" stroke="#268457" />
    </>
  ),
  deleteCells: (
    <>
      <path d="M14 3H3v18h18V10M3 9h9M3 15h18M9 9v12" />
      <path d="m15 2 7 7m0-7-7 7" stroke="#c35757" />
    </>
  ),
  cellSize: (
    <>
      <rect x="3" y="8" width="18" height="12" />
      <path d="M9 8v12M15 8v12M3 14h18M4 3h16M6 1 4 3l2 2m12-4 2 2-2 2" />
    </>
  ),
  fillDown: (
    <>
      <rect x="3" y="3" width="18" height="18" rx="1" />
      <path d="M12 6v11m-4-4 4 4 4-4" stroke="#3779ae" />
    </>
  ),
  sum: <path d="M19 4H5l8 8-8 8h14" />,
  sort: (
    <>
      <path d="M11 5h10" />
      <path d="M11 9h7" />
      <path d="M11 13h4" />
      <path d="M3 17l3 3 3-3" />
      <path d="M6 4v16" />
    </>
  ),
  search: (
    <>
      <circle cx="11" cy="11" r="7" />
      <path d="M21 21l-4.3-4.3" />
    </>
  ),
  undo: (
    <>
      <path d="M9 14 4 9l5-5" />
      <path d="M4 9h11a5 5 0 0 1 0 10H7" />
    </>
  ),
  copy: (
    <>
      <rect x="9" y="9" width="13" height="13" rx="2" />
      <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
    </>
  ),
  paste: (
    <>
      <path d="M8 4h2a2 2 0 0 1 4 0h2v3H8z" />
      <path d="M7 5H5a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h10" />
      <path d="M17 5h2a2 2 0 0 1 2 2v6" />
      <rect x="11" y="13" width="10" height="9" rx="1" />
    </>
  ),
  cut: (
    <>
      <circle cx="5" cy="18" r="2" />
      <circle cx="12" cy="18" r="2" />
      <path d="M7 16 19 3M10 16 4 3M14 12l5 7" />
    </>
  ),
  paint: (
    <>
      <path d="M3 5h14v5H3zM17 7h3v6h-8v3" />
      <path d="M12 16h4v6h-4z" />
    </>
  ),
  border: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M12 3v18M3 12h18" />
    </>
  ),
  fill: (
    <>
      <path d="m5 11 6-6 8 8-6 6zM3 21h18" />
      <path d="M15 21v-3" />
    </>
  ),
  alignLeft: (
    <>
      <path d="M3 4v16M7 6h13M7 11h9M7 16h13" />
    </>
  ),
  alignCenter: (
    <>
      <path d="M12 3v18M4 6h16M7 11h10M4 16h16" />
    </>
  ),
  alignRight: (
    <>
      <path d="M21 4v16M4 6h13M8 11h9M4 16h13" />
    </>
  ),
  alignTop: (
    <>
      <path d="M3 4h18M6 8h12M8 12h8M6 16h12" />
    </>
  ),
  alignMiddle: (
    <>
      <path d="M3 12h18M6 4h12M8 8h8M8 16h8M6 20h12" />
    </>
  ),
  alignBottom: (
    <>
      <path d="M3 20h18M6 8h12M8 12h8M6 16h12" />
    </>
  ),
  wrap: (
    <>
      <path d="M3 5h18M3 9h18M3 13h12a3 3 0 0 1 0 6h-3" />
      <path d="m15 16-3 3 3 3M3 19h5" />
    </>
  ),
  merge: (
    <>
      <rect x="2" y="4" width="20" height="16" />
      <path d="M8 4v16M16 4v16M4 12h16m-3-2 3 2-3 2m-10-4-3 2 3 2" />
    </>
  ),
  conditional: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="M3 9h18M9 3v18" />
      <path d="m12 15 2 2 4-5" />
    </>
  ),
  table: (
    <>
      <rect x="2" y="3" width="20" height="18" />
      <path d="M2 8h20M9 8v13M16 8v13" />
    </>
  ),
  cellStyle: (
    <>
      <rect x="3" y="3" width="18" height="18" />
      <path d="m7 16 4-8 4 8M8 14h6M17 6h2" />
    </>
  ),
  clear: (
    <>
      <path d="m4 14 10-10 7 7-9 9H8zM13 20h9" />
      <path d="m10 8 7 7" />
    </>
  ),
};
interface IconProps {
  name: keyof typeof paths;
  size?: number;
  className?: string;
}

export function Icon({ name, size = 16, className }: IconProps): JSX.Element {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      className={className}
    >
      {paths[name]}
    </svg>
  );
}
