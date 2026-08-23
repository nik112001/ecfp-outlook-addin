/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Office: any;

/**
 * No-op action placeholder.
 * Register on the global window so the manifest FunctionName can resolve it
 * when a button uses Action type="ExecuteFunction".
 */
export function action(event: { completed: () => void }): void {
  // No-op in M0 — real logic added in later milestones
  event.completed();
}

Office.onReady(() => {
  // Expose functions on the global window object so Office.js can invoke them
  // by name when the manifest specifies Action xsi:type="ExecuteFunction".
  (window as any).action = action;
});
