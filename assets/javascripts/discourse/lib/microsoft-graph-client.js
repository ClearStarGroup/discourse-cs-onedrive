import "../../vendor/microsoft-graph-client";

const graphGlobal = window.MicrosoftGraph;

if (!graphGlobal?.Client) {
  throw new Error("Microsoft Graph global not available after loading bundle");
}

export const Client = graphGlobal.Client;

export default graphGlobal;
