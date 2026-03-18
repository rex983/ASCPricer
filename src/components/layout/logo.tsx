/**
 * Big Buildings Direct logo — steel building silhouette with brand text.
 */
export function Logo({ size = "default" }: { size?: "small" | "default" | "large" }) {
  const iconSize = size === "small" ? "h-6 w-6" : size === "large" ? "h-10 w-10" : "h-7 w-7";
  const textSize = size === "small" ? "text-sm" : size === "large" ? "text-2xl" : "text-lg";

  return (
    <div className="flex items-center gap-2">
      <svg
        className={iconSize}
        viewBox="0 0 40 40"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
      >
        {/* Building body */}
        <rect x="4" y="16" width="32" height="20" rx="1" className="fill-blue-600" />
        {/* A-frame roof */}
        <path d="M2 17L20 4L38 17H2Z" className="fill-blue-700" />
        {/* Ridge line */}
        <path d="M20 4L20 17" className="stroke-blue-500" strokeWidth="1" />
        {/* Roll-up door */}
        <rect x="15" y="24" width="10" height="12" rx="0.5" className="fill-blue-800" />
        <line x1="15" y1="27" x2="25" y2="27" className="stroke-blue-600" strokeWidth="0.5" />
        <line x1="15" y1="30" x2="25" y2="30" className="stroke-blue-600" strokeWidth="0.5" />
        <line x1="15" y1="33" x2="25" y2="33" className="stroke-blue-600" strokeWidth="0.5" />
        {/* Side panels */}
        <line x1="10" y1="17" x2="10" y2="36" className="stroke-blue-800" strokeWidth="0.75" />
        <line x1="30" y1="17" x2="30" y2="36" className="stroke-blue-800" strokeWidth="0.75" />
        {/* Roof ribs */}
        <line x1="8" y1="12.5" x2="8" y2="17" className="stroke-blue-500" strokeWidth="0.5" />
        <line x1="14" y1="8.8" x2="14" y2="17" className="stroke-blue-500" strokeWidth="0.5" />
        <line x1="26" y1="8.8" x2="26" y2="17" className="stroke-blue-500" strokeWidth="0.5" />
        <line x1="32" y1="12.5" x2="32" y2="17" className="stroke-blue-500" strokeWidth="0.5" />
      </svg>
      <div className="flex flex-col leading-tight">
        <span className={`${textSize} font-bold tracking-tight`}>
          Big Buildings Direct
        </span>
      </div>
    </div>
  );
}

/** Icon-only version for compact spaces */
export function LogoIcon({ className = "h-6 w-6" }: { className?: string }) {
  return (
    <svg
      className={className}
      viewBox="0 0 40 40"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <rect x="4" y="16" width="32" height="20" rx="1" className="fill-blue-600" />
      <path d="M2 17L20 4L38 17H2Z" className="fill-blue-700" />
      <path d="M20 4L20 17" className="stroke-blue-500" strokeWidth="1" />
      <rect x="15" y="24" width="10" height="12" rx="0.5" className="fill-blue-800" />
      <line x1="15" y1="27" x2="25" y2="27" className="stroke-blue-600" strokeWidth="0.5" />
      <line x1="15" y1="30" x2="25" y2="30" className="stroke-blue-600" strokeWidth="0.5" />
      <line x1="15" y1="33" x2="25" y2="33" className="stroke-blue-600" strokeWidth="0.5" />
      <line x1="10" y1="17" x2="10" y2="36" className="stroke-blue-800" strokeWidth="0.75" />
      <line x1="30" y1="17" x2="30" y2="36" className="stroke-blue-800" strokeWidth="0.75" />
    </svg>
  );
}
