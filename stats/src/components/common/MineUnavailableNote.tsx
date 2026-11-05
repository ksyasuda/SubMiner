/** Stands in for the mine buttons when a line can't be mined, and says why. */
export function MineUnavailableNote({ reason }: { reason: string }) {
  return <span className="text-xs text-ctp-peach">{reason}</span>;
}
