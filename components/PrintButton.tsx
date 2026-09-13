"use client";

import Icon from "@/components/shell/Icon";

export default function PrintButton() {
  return (
    <button
      type="button"
      className="btn secondary no-print"
      onClick={() => window.print()}
    >
      <Icon name="printer" size={16} /> Yazdır / PDF
    </button>
  );
}
