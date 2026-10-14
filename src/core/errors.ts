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

/** Microsoft's own words for "this grant is not enough, ask the user". */
const consentPattern = /AADSTS65001|interaction_required|consent_required/i;

export function isSessionExpiry(reason: string): boolean {
  return expiryPattern.test(reason);
}

/**
 * Three answers, and the middle one now has to be earned. Labelling every
 * non-expiry failure a consent problem sends an agent to re-consent on a machine
 * whose session is fine - a laptop asleep, a DNS failure, a Microsoft incident -
 * and a consent flow spends the device-code rate limit while it is at it. So
 * `needsConsent` requires Microsoft to have said so, and anything else is
 * reported as what it actually is: the check could not be completed.
 */
export type SessionProblem = "expired" | "needsConsent" | "unavailable";

export function sessionProblem(reason: string): SessionProblem {
  if (isSessionExpiry(reason)) return "expired";
  if (consentPattern.test(reason)) return "needsConsent";
  return "unavailable";
}

/** A saved grant that does not cover what this process asks for. */
function reauthRequiredMessage(reason: string, scopes: string[]) {
  return `Outlook session needs re-authorisation: the saved grant does not cover what this server asks for (${scopes.join(", ")}). Run the login command again to grant them (${reason}).`;
}

/** Nothing is wrong with the session that a retry might not fix by itself. */
function unavailableMessage(reason: string) {
  return `Could not confirm the Outlook session (${reason}). Nothing was changed and no re-consent is needed for this - try the call again, and if it persists check the network and login.microsoftonline.com.`;
}

/** The one place that decides which wording a failure deserves. */
export function sessionProblemMessage(reason: string, scopes: string[]) {
  const problem = sessionProblem(reason);
  if (problem === "expired") return expiredSessionMessage(reason);
  if (problem === "needsConsent") return reauthRequiredMessage(reason, scopes);
  return unavailableMessage(reason);
}

/** Raised when a call needs a signed-in user; the message is the remediation. */
export class AuthRequiredError extends Error {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = "AuthRequiredError";
  }
}
