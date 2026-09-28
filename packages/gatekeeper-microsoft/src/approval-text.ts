// Untrusted text on its way into an approval prompt.
//
// Everything these two helpers touch is attacker-controlled — a mail subject, a Teams message, a
// list field a stranger submitted — so their whole job is to stop it forging the prompt around it:
// a title stays one bounded line, and a field body is fenced with a fence longer than any backtick
// run inside it.
//
// Shared by the Outlook mailbox, Teams and SharePoint list gatekeepers, which prompt the same way.

/** Longest single-line value echoed into an approval title. */
const MAX_APPROVAL_TITLE_CHARS = 200;

/** Fences an untrusted title into one bounded single line for an approval prompt. */
export function sanitizeApprovalTitle(value: string): string {
  return value.replace(/[\r\n]+/g, " ").slice(0, MAX_APPROVAL_TITLE_CHARS);
}

/** Fences an untrusted field value into a labelled code block that cannot forge approval Markdown. */
export function formatApprovalField(label: string, value: string): string {
  // Use a fence longer than any backtick run in the value, so untrusted text renders verbatim and
  // cannot forge surrounding approval Markdown.
  let fence = "```";
  while (value.includes(fence)) fence += "`";
  return `**${label}:**\n\n${fence}\n${value}\n${fence}`;
}
