export default function Logo({ size = 1 }: { size?: number }) {
  return (
    <span
      style={{
        display: "inline-flex",
        flexDirection: "column",
        alignItems: "center",
        lineHeight: 1.1,
      }}
    >
      <span
        style={{
          fontSize: 34 * size,
          fontWeight: 300,
          letterSpacing: "0.42em",
          marginRight: "-0.42em",
          color: "var(--brand)",
        }}
      >
        OLGA
      </span>
      <span
        style={{
          fontSize: 10.5 * size,
          fontWeight: 600,
          letterSpacing: "0.62em",
          marginRight: "-0.62em",
          color: "var(--muted)",
        }}
      >
        ÇERÇEVE
      </span>
    </span>
  );
}
