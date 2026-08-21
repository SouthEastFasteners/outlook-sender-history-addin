import { createNestablePublicClientApplication, IPublicClientApplication } from "@azure/msal-browser";
import { msalConfig, graphScopes } from "./authConfig";

let pca: IPublicClientApplication | null = null;

/**
 * Nested App Authentication (NAA) lets the task pane silently reuse the
 * identity the user is already signed into Outlook/Windows with — no
 * separate login prompt, no on-behalf-of middle-tier server required.
 */
async function getPca(): Promise<IPublicClientApplication> {
  if (!pca) {
    pca = await createNestablePublicClientApplication(msalConfig);
  }
  return pca;
}

export async function getGraphToken(): Promise<string> {
  const client = await getPca();

  try {
    const result = await client.acquireTokenSilent({ scopes: graphScopes });
    return result.accessToken;
  } catch (silentError) {
    // Silent acquisition can fail on first run or if consent is needed —
    // fall back to an interactive popup.
    const result = await client.acquireTokenPopup({ scopes: graphScopes });
    return result.accessToken;
  }
}
