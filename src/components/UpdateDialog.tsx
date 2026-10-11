import { useCallback, useEffect, useState } from "react";
import { Download } from "lucide-react";

import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  checkForUpdate,
  installUpdate,
  restartToApply,
  type Release,
  type UpdateStatus,
} from "@/lib/update";

interface UpdateDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  status: UpdateStatus | null;
  /** Lets the toolbar mark its button when a release is waiting. */
  onAvailableChange: (available: boolean) => void;
}

/**
 * `installing` covers the download and the checks, which the window cannot
 * report progress for; `restarting` is the point of no return, since the app
 * exits and the staged build takes over.
 */
type Phase = "checking" | "ready" | "installing" | "restarting";

export function UpdateDialog({
  open,
  onOpenChange,
  status,
  onAvailableChange,
}: UpdateDialogProps) {
  const [release, setRelease] = useState<Release | null>(null);
  const [phase, setPhase] = useState<Phase>("checking");
  const [error, setError] = useState<string | null>(null);

  const check = useCallback(() => {
    setPhase("checking");
    setError(null);
    return checkForUpdate()
      .then((found) => {
        setRelease(found);
        onAvailableChange(found !== null);
      })
      .catch((reason: unknown) => setError(String(reason)))
      .finally(() => setPhase("ready"));
  }, [onAvailableChange]);

  // Re-checked on every open rather than reusing the startup result: the
  // releases page is the thing being asked about, and it may have moved on.
  useEffect(() => {
    if (open) void check();
  }, [open, check]);

  const install = async () => {
    setPhase("installing");
    setError(null);
    try {
      await installUpdate();
      setPhase("restarting");
      // The app is about to exit, so this never resolves; a failure here is not
      // worth reporting, since there is no longer a window to report it in.
      void restartToApply().catch(() => {});
    } catch (reason: unknown) {
      setPhase("ready");
      setError(String(reason));
    }
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle className="text-sm">Software update</DialogTitle>
          <DialogDescription className="text-xs">
            {status ? `SheetLens ${status.version}` : "SheetLens"}
          </DialogDescription>
        </DialogHeader>

        {phase === "checking" && (
          <p className="text-xs text-muted-foreground">Checking for updates…</p>
        )}

        {phase !== "checking" && release && (
          <div className="space-y-2">
            <p className="text-xs">
              Version <span className="font-medium">{release.version}</span> is
              available.
            </p>
            {release.notes && (
              <pre className="max-h-40 overflow-auto rounded border border-border bg-muted/40 p-2 text-[11px] whitespace-pre-wrap text-muted-foreground">
                {release.notes}
              </pre>
            )}
          </div>
        )}

        {phase === "ready" && !release && !error && (
          <p className="text-xs text-muted-foreground">
            This is the latest version.
          </p>
        )}

        {phase === "installing" && (
          <p className="text-xs text-muted-foreground">
            Downloading and checking the update…
          </p>
        )}

        {phase === "restarting" && (
          <p className="text-xs text-muted-foreground">
            Restarting to finish the update…
          </p>
        )}

        {error && (
          <p className="rounded border border-destructive/40 bg-destructive/10 p-2 text-[11px] text-destructive">
            {error}
          </p>
        )}

        <DialogFooter>
          {phase === "ready" && release && (
            <Button size="sm" className="h-7 gap-1.5 text-xs" onClick={install}>
              <Download className="size-3.5" />
              Update and restart
            </Button>
          )}
          {phase === "ready" && !release && (
            <Button
              size="sm"
              variant="outline"
              className="h-7 text-xs"
              onClick={() => void check()}
            >
              Check again
            </Button>
          )}
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
