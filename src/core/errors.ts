export const noSessionMessage =
  "No Outlook session. Run the login command first.";

export function expiredSessionMessage(reason: string) {
  return `Outlook session expired (${reason}). Run the login command again.`;
}

/**
 * Why a silent refresh failed is something Microsoft states, and expiry is only
 * one of the reasons. A session signed in under an earlier scope set is not
 * expired, it is merely too narrow, and telling someone their session expired
 * sends them looking for a clock instead of a consent prompt.
 */
const expiryPattern = /expired|AADSTS700082|invalid_grant/i;

export function isSessionExpiry(reason: string): boolean {
  return expiryPattern.test(reason);
}

/** A saved grant that does not cover what this process asks for. */
function reauthRequiredMessage(reason: string, scopes: string[]) {
  return `Outlook session needs re-authorisation: the saved grant does not cover what this server asks for (${scopes.join(", ")}). Run the login command again to grant them (${reason}).`;
}

/** The one place that decides which of the two wordings a failure deserves. */
export function sessionProblemMessage(reason: string, scopes: string[]) {
  return isSessionExpiry(reason)
    ? expiredSessionMessage(reason)
    : reauthRequiredMessage(reason, scopes);
}

/** Raised when a call needs a signed-in user; the message is the remediation. */
export class AuthRequiredError extends Error {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = "AuthRequiredError";
  }
}
