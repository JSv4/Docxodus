// Every demo loads the package built alongside this site. An engine override remains
// useful for integration tests and release comparisons; history is a shared host option.
const params = new URLSearchParams(location.search);
export const engineUrl = new URL(params.get('engine') ?? './embed.bundle.js', location.href);

export async function loadDemoEngine() {
  const engine = await import(engineUrl.href);
  return {
    ...engine,
    createRibbonEditor(container, source, options = {}) {
      const history = params.get('history') === '0' ? false : {
        storageName: 'docxodus-demo',
        // Ordinary editors resume their saved workspace; animation hosts opt out below.
        workspaceId: typeof source === 'string' ? `editor:${new URL(source, location.href).href}` : `editor:${location.pathname}:blank`,
      };
      return engine.createRibbonEditor(container, source, { history, ...options });
    },
  };
}
