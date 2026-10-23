/**
 * Heartbeat and check periods for a hang limit: at least 4 heartbeats per
 * limit, so a healthy but busy process never looks blocked (a fixed 500 ms
 * beat with a 300 ms limit used to kill healthy processes under load).
 */
export const periods = (hangMs: number): { readonly beatMs: number; readonly checkMs: number } => ({
  beatMs: Math.max(1, Math.min(500, Math.floor(hangMs / 4))),
  checkMs: Math.max(1, Math.min(250, Math.floor(hangMs / 4)))
})
