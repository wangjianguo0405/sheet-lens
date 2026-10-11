import { invoke } from "@tauri-apps/api/core";

export interface UpdateStatus {
  /** The version this build reports. */
  version: string;
  /**
   * Whether this copy can replace itself: it is portable, and the marker file
   * sits beside the executable. A dev or installed build is not.
   */
  portable: boolean;
}

export interface Release {
  version: string;
  notes: string | null;
}

export function updateStatus(): Promise<UpdateStatus> {
  return invoke<UpdateStatus>("update_status");
}

/** Returns the newer release, or null when this build is already the latest. */
export function checkForUpdate(): Promise<Release | null> {
  return invoke<Release | null>("update_check");
}

/**
 * Downloads the release, checks it and stages the swap. Nothing changes on disk
 * until the app exits — `restartToApply` is what lets the swap happen.
 */
export function installUpdate(): Promise<void> {
  return invoke<void>("update_install");
}

/** Exits, so the staged executable can replace this one and start in its place. */
export function restartToApply(): Promise<void> {
  return invoke<void>("update_restart");
}
