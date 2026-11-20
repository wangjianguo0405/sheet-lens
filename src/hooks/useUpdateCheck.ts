import { useEffect, useState } from "react";

import { checkForUpdate, updateStatus, type UpdateStatus } from "@/lib/update";

/**
 * Whether a newer release is waiting.
 *
 * A silent check at startup, so a waiting release earns a dot on the toolbar:
 * nothing is shown when there is none, when the build cannot replace itself, or
 * when the check fails — being offline is not worth interrupting anyone.
 */
export function useUpdateCheck(language: string) {
  /** Only a portable build reports `portable`, and only then is it offered. */
  const [update, setUpdate] = useState<UpdateStatus | null>(null);
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [updateOpen, setUpdateOpen] = useState(false);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const status = await updateStatus();
        if (cancelled) return;
        setUpdate(status);
        if (!status.portable) return;
        const found = await checkForUpdate(language);
        if (!cancelled) setUpdateAvailable(found !== null);
      } catch {
        /* nothing to report */
      }
    })();
    return () => {
      cancelled = true;
    };
    // Runs once at startup on purpose: it asks for the update notes in the
    // language active then. Switching language must not re-run the check.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return { update, updateAvailable, setUpdateAvailable, updateOpen, setUpdateOpen };
}
