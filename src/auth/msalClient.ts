/**
 * MSAL PublicClientApplication configured for a single personal/dev Microsoft account.
 * Scope: Mail.ReadBasic (metadata only — no body content per spec §3.3).
 * Auth flow: silent token acquisition → popup fallback (no redirect, add-in context).
 *
 * Spec refs: §3.2 (Key Design Decisions — Auth row), §3.3 (Permissions & Scopes)
 */

import {
  PublicClientApplication,
  Configuration,
  AuthenticationResult,
  InteractionRequiredAuthError,
  AccountInfo,
} from "@azure/msal-browser";

// ---------------------------------------------------------------------------
// Scopes
// ---------------------------------------------------------------------------

/**
 * Delegated Graph scopes requested at sign-in.
 * `Mail.ReadBasic` grants access to message metadata (including size) but never
 * to message bodies — consistent with spec §3.3 privacy constraints.
 */
export const GRAPH_SCOPES = ["Mail.ReadBasic"];

// ---------------------------------------------------------------------------
// MSAL configuration
// ---------------------------------------------------------------------------

const msalConfig: Configuration = {
  auth: {
    /**
     * Replace with your Azure app registration client ID before first Graph call.
     * Register the app at https://portal.azure.com → Azure Active Directory →
     * App registrations → New registration (Accounts in any personal Microsoft
     * account only).
     */
    clientId: "ad6766ae-f77d-4bda-8071-b910c2fe5268",

    /**
     * "common" accepts both personal Microsoft accounts (MSA / Outlook.com)
     * and work/school accounts (Entra ID). Use "consumers" to restrict to
     * personal accounts only, or a specific tenant ID to lock to one org.
     */
    authority: "https://login.microsoftonline.com/common",

    /**
     * The add-in's task pane URL.  Must be registered as an allowed redirect URI
     * in the Azure portal under the app registration → Authentication → SPA platform.
     */
    redirectUri: "https://localhost:3000/taskpane.html",
  },
  cache: {
    /**
     * localStorage persists across browser sessions so the user doesn't have to
     * sign in every time the task pane is opened.
     */
    cacheLocation: "localStorage",

    /**
     * Cookies for auth state are not needed because the add-in runs in modern
     * browsers where localStorage is available and reliable.
     */
    storeAuthStateInCookie: false,
  },
};

// ---------------------------------------------------------------------------
// Lazy singleton
// ---------------------------------------------------------------------------

let _msalInstance: PublicClientApplication | null = null;

/**
 * Returns the shared MSAL `PublicClientApplication` instance, creating it on
 * the first call (lazy singleton).  Subsequent calls return the same object so
 * the in-memory token cache is shared across the module.
 */
export function getMsalInstance(): PublicClientApplication {
  if (!_msalInstance) {
    _msalInstance = new PublicClientApplication(msalConfig);
  }
  return _msalInstance;
}

// ---------------------------------------------------------------------------
// Token acquisition
// ---------------------------------------------------------------------------

/**
 * Returns a valid access token for `GRAPH_SCOPES`.
 *
 * Strategy (required in add-in context — no full-page redirect allowed):
 *  1. Attempt a silent token acquisition against the cached account.
 *  2. If the cache is empty or the token is expired and requires user
 *     interaction (`InteractionRequiredAuthError`), fall back to a popup.
 *  3. Store the returned account on the instance so subsequent silent calls
 *     succeed without another popup.
 *
 * @throws {Error} with a descriptive message so callers can surface a
 *   "Sign in required" prompt in the UI.
 */
export async function getAccessToken(): Promise<string> {
  const msalInstance = getMsalInstance();

  // Initialise the instance (resolves any pending redirects; safe to call
  // multiple times — MSAL guards against double-initialisation internally).
  await msalInstance.initialize();

  // Prefer the first cached account; users of this add-in are expected to
  // have at most one account signed in at a time.
  const accounts: AccountInfo[] = msalInstance.getAllAccounts();
  const account: AccountInfo | undefined = accounts[0];

  const tokenRequest = {
    scopes: GRAPH_SCOPES,
    account,
  };

  let result: AuthenticationResult;

  try {
    // 1. Silent path — uses the in-memory / localStorage cache.
    result = await msalInstance.acquireTokenSilent(tokenRequest);
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      // 2. Popup fallback — shown when:
      //    - No cached account (first sign-in).
      //    - Cached token expired and refresh token is also gone/invalid.
      //    - Consent is required for a new scope.
      try {
        result = await msalInstance.acquireTokenPopup(tokenRequest);
      } catch (popupErr) {
        const message =
          popupErr instanceof Error ? popupErr.message : String(popupErr);
        throw new Error(
          `Sign-in required but popup was blocked or dismissed: ${message}`
        );
      }
    } else {
      const message = err instanceof Error ? err.message : String(err);
      throw new Error(`Failed to acquire access token silently: ${message}`);
    }
  }

  // 3. Persist the returned account so the next silent call finds it.
  if (result.account) {
    msalInstance.setActiveAccount(result.account);
  }

  return result.accessToken;
}

// ---------------------------------------------------------------------------
// Account helpers
// ---------------------------------------------------------------------------

/**
 * Returns the email address of the currently signed-in Microsoft account, or
 * `null` if no account is present in the MSAL cache.
 *
 * Checks `username` (populated for personal accounts) then falls back to the
 * `login_hint` claim inside the ID token claims object.
 */
export function getSignedInEmail(): string | null {
  const msalInstance = getMsalInstance();
  const accounts = msalInstance.getAllAccounts();

  if (accounts.length === 0) {
    return null;
  }

  const account = accounts[0];

  // `username` is the UPN / email for personal Microsoft accounts.
  if (account.username) {
    return account.username;
  }

  // Fallback: extract from id-token claims when username is unavailable.
  const claims = account.idTokenClaims as Record<string, unknown> | undefined;
  if (claims) {
    const email =
      (claims["email"] as string | undefined) ??
      (claims["preferred_username"] as string | undefined) ??
      null;
    if (email) return email;
  }

  return null;
}

// ---------------------------------------------------------------------------
// Sign-out
// ---------------------------------------------------------------------------

/**
 * Clears all MSAL cache state for the current account.
 *
 * Intended for the "Sign out" button in the Settings pane (M4).  Uses
 * `logoutPopup` so the task-pane context is not disrupted by a full-page
 * navigation.
 */
export async function signOut(): Promise<void> {
  const msalInstance = getMsalInstance();

  await msalInstance.initialize();

  const account = msalInstance.getAllAccounts()[0];

  if (!account) {
    // Nothing to sign out; clear any residual cache state just in case.
    msalInstance.clearCache();
    return;
  }

  try {
    await msalInstance.logoutPopup({ account });
  } catch (err) {
    // If the popup is blocked or the user closes it, fall back to a local
    // cache-only clear so the UI can still transition to the signed-out state.
    msalInstance.clearCache();
    const message = err instanceof Error ? err.message : String(err);
    console.warn(
      `logoutPopup did not complete cleanly; local cache cleared. Reason: ${message}`
    );
  }
}
