import { useEffect, useRef, useState } from "react";

import { checkForUpdate, updateStatus, type UpdateStatus } from "@/lib/update";

/** How often a running copy asks again. A build is left open across a working
 * day far more often than it is restarted, and a release lands in that time;
 * four hours is short enough to be seen the same day and rare enough to be
 * nothing against GitHub's limits. */
const RECHECK_MS = 4 * 60 * 60 * 1000;

/**
 * Whether a newer release is waiting.
 *
 * A silent check at startup, then the same check every four hours, so a waiting
 * release earns a dot on the toolbar without the reader having to restart.
 * Nothing is shown when there is none, when the build cannot replace itself, or
 * when the check fails — being offline is not worth interrupting anyone.
 */
export function useUpdateCheck(language: string) {
  /** Only a portable build reports `portable`, and only then is it offered. */
  const [update, setUpdate] = useState<UpdateStatus | null>(null);
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [updateOpen, setUpdateOpen] = useState(false);
  /** The language showing now, for a check that may run hours after the first. */
  const current = useRef(language);
  current.current = language;

  useEffect(() => {
    let cancelled = false;
    const check = async () => {
      try {
        const status = await updateStatus();
        if (cancelled) return;
        setUpdate(status);
        if (!status.portable) return;
        // Asked for in the language showing at the moment of the check, so a
        // reader who switched during the session is not shown the old one's.
        const found = await checkForUpdate(current.current);
        if (!cancelled) setUpdateAvailable(found !== null);
      } catch {
        /* nothing to report */
      }
    };
    void check();
    const timer = window.setInterval(() => void check(), RECHECK_MS);
    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, []);

  return { update, updateAvailable, setUpdateAvailable, updateOpen, setUpdateOpen };
}
