import Image from "next/image";

/**
 * American Steel Carports logo with brand text.
 */
export function Logo({ size = "default" }: { size?: "small" | "default" | "large" }) {
  const imgSize = size === "small" ? 32 : size === "large" ? 56 : 40;
  const textSize = size === "small" ? "text-sm" : size === "large" ? "text-2xl" : "text-lg";

  return (
    <div className="flex items-center gap-2">
      <Image
        src="/ASCI.png"
        alt="American Steel Carports"
        width={imgSize}
        height={imgSize}
        className="rounded-full"
      />
      <div className="flex flex-col leading-tight">
        <span className={`${textSize} font-bold tracking-tight`}>
          American Steel Carports
        </span>
      </div>
    </div>
  );
}

/** Icon-only version for compact spaces */
export function LogoIcon({ className = "h-6 w-6" }: { className?: string }) {
  return (
    <Image
      src="/ASCI.png"
      alt="ASC"
      width={24}
      height={24}
      className={`${className} rounded-full`}
    />
  );
}
