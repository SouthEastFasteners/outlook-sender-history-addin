import * as React from "react";
import { useEffect, useState } from "react";
import { getRecentEmailsFromSender, EmailSummary } from "../graph";

/* global Office */

interface SenderInfo {
  name: string;
  email: string;
  itemId: string;
}

function getInitials(name: string): string {
  const parts = name.trim().split(/\s+/);
  const initials = parts.slice(0, 2).map((p) => p[0]?.toUpperCase() ?? "");
  return initials.join("") || "?";
}

function formatRelativeDate(iso: string): string {
  const date = new Date(iso);
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  if (diffDays === 0) {
    return date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  }
  if (diffDays === 1) return "Yesterday";
  if (diffDays < 7) return date.toLocaleDateString([], { weekday: "short" });
  return date.toLocaleDateString([], { month: "short", day: "numeric" });
}

export default function App() {
  const [sender, setSender] = useState<SenderInfo | null>(null);
  const [emails, setEmails] = useState<EmailSummary[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  async function loadForCurrentItem() {
    const item = Office.context.mailbox.item;
    if (!item || !item.from) {
      setError("No message is currently open.");
      setLoading(false);
      return;
    }

    const senderInfo: SenderInfo = {
      name: item.from.displayName || item.from.emailAddress,
      email: item.from.emailAddress,
      itemId: item.itemId ?? "",
    };
    setSender(senderInfo);
    setLoading(true);
    setError(null);

    try {
      const results = await getRecentEmailsFromSender(senderInfo.email, senderInfo.itemId, 6);
      setEmails(results);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Couldn't load email history.");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadForCurrentItem();

    // Re-run whenever the user selects a different message while the
    // pane is pinned open.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      loadForCurrentItem
    );

    return () => {
      Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function openEmail(id: string) {
    Office.context.mailbox.displayMessageForm(id);
  }

  return (
    <div className="pane">
      <div className="pane-header">
        <div className="pane-title">Sender History</div>
      </div>

      {sender && (
        <div className="context-block">
          <div className="context-label">Viewing email from</div>
          <div className="sender-row">
            <div className="avatar">{getInitials(sender.name)}</div>
            <div className="sender-info">
              <div className="sender-name">{sender.name}</div>
              <div className="sender-email">{sender.email}</div>
            </div>
          </div>
        </div>
      )}

      <div className="history-section">
        <div className="history-header">
          <div className="history-title">Recent emails</div>
          {!loading && <div className="history-count">{emails.length} found</div>}
        </div>

        {loading && <div className="status-text">Loading history…</div>}
        {error && <div className="status-text error">{error}</div>}

        {!loading && !error && emails.length === 0 && (
          <div className="status-text">No previous emails found from this sender.</div>
        )}

        {!loading && !error && emails.length > 0 && (
          <div className="email-list">
            {emails.map((email) => (
              <div
                key={email.id}
                className={`email-item${!email.isRead ? " unread" : ""}`}
                onClick={() => openEmail(email.id)}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => {
                  if (e.key === "Enter") openEmail(email.id);
                }}
              >
                <div className="email-main">
                  <div className="email-top-row">
                    <div className="email-subject">{email.subject || "(No subject)"}</div>
                    <div className="email-date">{formatRelativeDate(email.receivedDateTime)}</div>
                  </div>
                  <div className="email-preview">{email.bodyPreview}</div>
                  {email.hasAttachments && (
                    <div className="email-meta-row">
                      <span className="tag attachment">Attachment</span>
                    </div>
                  )}
                </div>
                <span className="chevron">›</span>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="footnote">Click any email to open it without leaving this message</div>
    </div>
  );
}
