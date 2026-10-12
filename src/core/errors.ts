export const noSessionMessage =
  "No Outlook session. Run the login command first.";

export function expiredSessionMessage(reason: string) {
  return `Outlook session expired (${reason}). Run the login command again.`;
}

/** Raised when a call needs a signed-in user; the message is the remediation. */
export class AuthRequiredError extends Error {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = "AuthRequiredError";
  }
}
