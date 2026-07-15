import type { SVGProps } from 'react';

export type IconName =
  | 'arrow-left'
  | 'arrow-right'
  | 'check'
  | 'chevron-down'
  | 'diagnostics'
  | 'expand'
  | 'file'
  | 'fit'
  | 'info'
  | 'minus'
  | 'plus'
  | 'replace'
  | 'search'
  | 'slides'
  | 'upload'
  | 'warning'
  | 'x';

const paths: Record<IconName, React.ReactNode> = {
  'arrow-left': <path d="m14.5 5-7 7 7 7" />,
  'arrow-right': <path d="m9.5 5 7 7-7 7" />,
  check: <path d="m5 12 4 4L19 6" />,
  'chevron-down': <path d="m7 9.5 5 5 5-5" />,
  diagnostics: (
    <>
      <path d="M4 19V9m5 10V5m6 14v-7m5 7V3" />
      <path d="M2 19h20" />
    </>
  ),
  expand: (
    <>
      <path d="M8 3H3v5M16 3h5v5M8 21H3v-5m13 5h5v-5" />
      <path d="m3 8 6-6m12 6-6-6M3 16l6 6m12-6-6 6" />
    </>
  ),
  file: (
    <>
      <path d="M7 3h7l4 4v14H7z" />
      <path d="M14 3v5h4M10 12h5m-5 4h5" />
    </>
  ),
  fit: (
    <>
      <rect x="4" y="6" width="16" height="12" rx="1" />
      <path d="M8 10h8M8 14h8" />
    </>
  ),
  info: (
    <>
      <circle cx="12" cy="12" r="9" />
      <path d="M12 11v6m0-10h.01" />
    </>
  ),
  minus: <path d="M5 12h14" />,
  plus: <path d="M12 5v14M5 12h14" />,
  replace: (
    <>
      <path d="M20 11a8 8 0 1 0-2.34 5.66" />
      <path d="M20 5v6h-6" />
    </>
  ),
  search: (
    <>
      <circle cx="10.5" cy="10.5" r="6.5" />
      <path d="m15.5 15.5 5 5" />
    </>
  ),
  slides: (
    <>
      <rect x="3" y="4" width="14" height="12" rx="1" />
      <path d="M7 20h14V8" />
    </>
  ),
  upload: (
    <>
      <path d="M12 16V4m-5 5 5-5 5 5" />
      <path d="M4 15v5h16v-5" />
    </>
  ),
  warning: (
    <>
      <path d="m12 3 10 18H2z" />
      <path d="M12 9v5m0 3h.01" />
    </>
  ),
  x: <path d="m6 6 12 12M18 6 6 18" />,
};

export function Icon({ name, ...props }: { name: IconName } & SVGProps<SVGSVGElement>) {
  return (
    <svg
      aria-hidden="true"
      fill="none"
      viewBox="0 0 24 24"
      stroke="currentColor"
      strokeWidth="1.75"
      strokeLinecap="round"
      strokeLinejoin="round"
      {...props}
    >
      {paths[name]}
    </svg>
  );
}
