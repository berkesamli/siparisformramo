/* eslint-disable @next/next/no-img-element */
export default function Logo({ size = 1 }: { size?: number }) {
  return (
    <span
      style={{
        display: "inline-flex",
        flexDirection: "column",
        alignItems: "center",
        lineHeight: 1.1,
        gap: 6 * size,
      }}
    >
      <img
        src="/logo.png"
        alt="Olga Çerçeve"
        style={{ height: 52 * size, width: "auto" }}
      />
      <span
        style={{
          fontSize: 10.5 * size,
          fontWeight: 600,
          letterSpacing: "0.62em",
          marginRight: "-0.62em",
          color: "var(--brand)",
        }}
      >
        ÇERÇEVE
      </span>
    </span>
  );
}
