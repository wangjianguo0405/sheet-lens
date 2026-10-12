//! Portable self-update.
//!
//! SheetLens ships as a bare executable in a folder, so there is no installer
//! for the updater plugin to hand an archive to. The plugin still does the part
//! that must not be hand-rolled — fetching the signed manifest, downloading the
//! release archive and checking the minisign signature over its bytes — and this
//! module does what is left: validating the archive, and replacing the
//! executable on disk, which Windows will not let a running process do itself.

use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};

use serde::{Deserialize, Serialize};
use tauri::{AppHandle, Emitter};
use tauri_plugin_updater::UpdaterExt;

/// An install is portable — and so can update itself — when this file sits
/// beside the executable. A build installed by an installer has no marker and
/// never reaches the code below.
const MARKER: &str = "portable.sheetlens";
/// The entry every release archive carries, whatever the local copy is named.
const PAYLOAD_EXECUTABLE: &str = "SheetLens.exe";
/// Describes the archive. It lives *inside* the signed zip, so it is covered by
/// the signature; the updater plugin additionally requires the signature to
/// carry the version it was signed for, which is what stops a tampered manifest
/// from pairing a new version number with an older release.
const PAYLOAD_MANIFEST: &str = "update-manifest.json";
/// Written beside the executable only while an update is being applied, and
/// deleted by the script itself as soon as it starts.
const HELPER_SCRIPT: &str = "apply-update.ps1";
const PAYLOAD_SCHEMA: u32 = 1;

#[derive(Deserialize)]
struct Payload {
    schema: u32,
    version: String,
    arch: String,
}

#[derive(Serialize)]
pub struct Status {
    pub version: String,
    pub portable: bool,
}

#[derive(Serialize)]
pub struct Release {
    pub version: String,
    pub notes: Option<String>,
}

/// The manifest's notes for `locale`, falling back to the English ones.
///
/// `notes` is English, and stays that way: it is the field every other updater
/// client reads. The release also carries a `notesLocales` map beside it, which
/// the plugin does not model — so it reaches us untouched in `raw_json`, and
/// nothing here has to fetch or parse the manifest a second time.
fn localised_notes(
    raw: &serde_json::Value,
    english: Option<&str>,
    locale: &str,
) -> Option<String> {
    raw.get("notesLocales")
        .and_then(|locales| locales.get(locale))
        .and_then(serde_json::Value::as_str)
        .or(english)
        .map(str::to_owned)
}

/// Broadcast while the archive downloads, so the window can draw a bar. Only
/// sent when the server announced a length — without one there is no
/// percentage to draw, and the window falls back to its own wording.
#[derive(Clone, Serialize)]
struct Progress {
    downloaded: u64,
    total: u64,
}

/// The event name the window listens on.
const PROGRESS_EVENT: &str = "update-progress";

/// The folder holding the running executable, if this is a portable install.
fn portable_dir() -> Result<PathBuf, String> {
    let exe = std::env::current_exe()
        .map_err(|error| format!("Could not locate the running executable: {error}"))?;
    let dir = exe
        .parent()
        .ok_or_else(|| "The running executable has no parent directory.".to_string())?;
    if !dir.join(MARKER).is_file() {
        return Err("This copy of SheetLens is not a portable install and cannot update itself. \
                    Download a new build instead."
            .to_string());
    }
    Ok(dir.to_path_buf())
}

fn arch_label() -> &'static str {
    match std::env::consts::ARCH {
        "x86_64" => "x64",
        "aarch64" => "arm64",
        other => other,
    }
}

/// Pulls the executable out of a release archive, after checking the archive
/// really is the build the manifest announced.
fn extract_executable(archive: &[u8], announced_version: &str) -> Result<Vec<u8>, String> {
    let mut zip = zip::ZipArchive::new(Cursor::new(archive))
        .map_err(|error| format!("The downloaded update is not a readable archive: {error}"))?;

    let payload: Payload = {
        let mut entry = zip
            .by_name(PAYLOAD_MANIFEST)
            .map_err(|_| format!("The update archive does not contain {PAYLOAD_MANIFEST}."))?;
        let mut text = String::new();
        entry
            .read_to_string(&mut text)
            .map_err(|error| format!("Could not read the update manifest: {error}"))?;
        serde_json::from_str(&text)
            .map_err(|error| format!("The update manifest is malformed: {error}"))?
    };

    if payload.schema != PAYLOAD_SCHEMA {
        return Err(format!(
            "This build cannot read update manifests of schema {}.",
            payload.schema
        ));
    }
    if payload.version != announced_version {
        return Err(format!(
            "The update archive holds version {} but version {announced_version} was announced.",
            payload.version
        ));
    }
    if payload.arch != arch_label() {
        return Err(format!(
            "The update archive holds a {} build, but this is {}.",
            payload.arch,
            arch_label()
        ));
    }

    let mut entry = zip
        .by_name(PAYLOAD_EXECUTABLE)
        .map_err(|_| format!("The update archive does not contain {PAYLOAD_EXECUTABLE}."))?;
    let mut executable = Vec::new();
    entry
        .read_to_end(&mut executable)
        .map_err(|error| format!("Could not read {PAYLOAD_EXECUTABLE}: {error}"))?;

    // A truncated or wrong-file payload would otherwise only fail later, after
    // the old build has already been moved aside.
    if !executable.starts_with(b"MZ") {
        return Err("The update archive does not hold a Windows executable.".to_string());
    }

    Ok(executable)
}

/// Whether an archive entry is one of the folder's documents rather than its
/// payload: the three READMEs, the three changelogs, and the licence.
fn is_document(name: &str) -> bool {
    name == "LICENSE"
        || (name.ends_with(".md") && (name.starts_with("README") || name.starts_with("CHANGELOG")))
}

/// The documents an archive carries, as they would land in the install folder.
///
/// The archive holds them so the folder explains itself without a network, which
/// was only true of the copy someone first unpacked: an update replaced the
/// executable and left every document beside it describing whichever version was
/// installed first, however many releases later.
fn documents_in(archive: &[u8]) -> Vec<(String, Vec<u8>)> {
    let Ok(mut zip) = zip::ZipArchive::new(Cursor::new(archive)) else {
        return Vec::new();
    };
    let mut documents = Vec::new();
    for index in 0..zip.len() {
        let Ok(mut entry) = zip.by_index(index) else {
            continue;
        };
        let name = entry.name().to_string();
        if !is_document(&name) {
            continue;
        }
        let mut bytes = Vec::new();
        if entry.read_to_end(&mut bytes).is_ok() {
            documents.push((name, bytes));
        }
    }
    documents
}

/// Rewrites the documents beside the executable from an archive whose signature
/// has already been checked.
///
/// Best effort, and deliberately so: the update is the executable, and a folder
/// someone made read-only — or a document they edited by hand — is not a reason
/// to refuse a build that verified.
fn refresh_documents(archive: &[u8], dir: &Path) {
    for (name, bytes) in documents_in(archive) {
        let _ = std::fs::write(dir.join(name), bytes);
    }
}

/// Windows PowerShell lives in a versioned directory that Rust's executable
/// search never visits — it probes `System32` itself and stops — so a machine
/// whose `PATH` lost that entry reports "program not found" while PowerShell is
/// installed and working. Resolve it up front and keep the bare name as a
/// fallback for installs that put it elsewhere.
#[cfg(target_os = "windows")]
fn powershell() -> std::ffi::OsString {
    let resolved = std::env::var_os("SystemRoot")
        .filter(|root| !root.is_empty())
        .map(|root| {
            Path::new(&root)
                .join("System32")
                .join("WindowsPowerShell")
                .join("v1.0")
                .join("powershell.exe")
        })
        .filter(|candidate| candidate.is_file());
    match resolved {
        Some(path) => path.into_os_string(),
        None => "powershell.exe".into(),
    }
}

/// Writes the new executable beside the old one and starts the detached script
/// that will swap them once this process is gone.
#[cfg(target_os = "windows")]
fn stage_and_launch(dir: &Path, executable: &[u8], exe_name: &str) -> Result<(), String> {
    use std::os::windows::process::CommandExt;
    use std::process::Command;

    /// Keep the helper off the console: it outlives the window it was started
    /// from, and a flashing console for an update nobody is watching is noise.
    const CREATE_NO_WINDOW: u32 = 0x0800_0000;
    /// So the helper is not killed along with the process group it inherits.
    const CREATE_NEW_PROCESS_GROUP: u32 = 0x0000_0200;

    let target = dir.join(exe_name);
    let staged = dir.join(format!("{exe_name}.new"));
    let backup = dir.join(format!("{exe_name}.old"));
    let script = dir.join(HELPER_SCRIPT);

    // Staging beside the target also proves the folder is writable: a copy under
    // Program Files fails here, and says so, rather than half-way through a swap.
    std::fs::write(&staged, executable).map_err(|error| {
        format!(
            "Could not write the update next to {}. The folder may be read-only, or need \
             administrator rights: {error}",
            target.display()
        )
    })?;
    std::fs::write(&script, HELPER_SCRIPT_BODY)
        .map_err(|error| format!("Could not write the update helper: {error}"))?;

    let spawned = Command::new(powershell())
        .args([
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
        ])
        .arg(&script)
        .arg("-ParentProcessId")
        .arg(std::process::id().to_string())
        .arg("-StagedExe")
        .arg(&staged)
        .arg("-TargetExe")
        .arg(&target)
        .arg("-BackupExe")
        .arg(&backup)
        .creation_flags(CREATE_NO_WINDOW | CREATE_NEW_PROCESS_GROUP)
        .spawn();

    if let Err(error) = spawned {
        let _ = std::fs::remove_file(&staged);
        let _ = std::fs::remove_file(&script);
        return Err(format!(
            "Could not start the update helper ({}): {error}",
            powershell().to_string_lossy()
        ));
    }

    Ok(())
}

#[cfg(not(target_os = "windows"))]
fn stage_and_launch(_dir: &Path, _executable: &[u8], _exe_name: &str) -> Result<(), String> {
    Err("Portable self-update is only implemented for Windows.".to_string())
}

/// The script that performs the swap. It must do nothing until this process has
/// exited — Windows will not let go of the running image before then — so it
/// waits on the pid, then retries the replacement until the file is free.
#[cfg(target_os = "windows")]
const HELPER_SCRIPT_BODY: &str = r#"param(
    [Parameter(Mandatory = $true)][int]$ParentProcessId,
    [Parameter(Mandatory = $true)][string]$StagedExe,
    [Parameter(Mandatory = $true)][string]$TargetExe,
    [Parameter(Mandatory = $true)][string]$BackupExe
)

$ErrorActionPreference = 'Stop'
Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
try { Wait-Process -Id $ParentProcessId -Timeout 120 -ErrorAction SilentlyContinue } catch {}

$installed = $false
for ($attempt = 0; $attempt -lt 120; $attempt++) {
    try {
        # The backup is worth clearing even when the target is gone: a leftover
        # from an earlier failed run would otherwise be moved over the new build.
        if (Test-Path -LiteralPath $TargetExe) {
            if (Test-Path -LiteralPath $BackupExe) { Remove-Item -LiteralPath $BackupExe -Force }
            Move-Item -LiteralPath $TargetExe -Destination $BackupExe -Force
        }
        if (-not (Test-Path -LiteralPath $BackupExe)) { throw 'The existing build was not backed up.' }

        Move-Item -LiteralPath $StagedExe -Destination $TargetExe -Force
        $installed = $true
        break
    } catch {
        # Never leave the folder without an executable: put the old build back
        # before waiting to try again.
        if (-not (Test-Path -LiteralPath $TargetExe) -and (Test-Path -LiteralPath $BackupExe)) {
            Copy-Item -LiteralPath $BackupExe -Destination $TargetExe -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 1
    }
}

if (-not $installed) {
    if (-not (Test-Path -LiteralPath $TargetExe) -and (Test-Path -LiteralPath $BackupExe)) {
        Copy-Item -LiteralPath $BackupExe -Destination $TargetExe -Force -ErrorAction SilentlyContinue
    }
    exit 1
}

try {
    Start-Process -FilePath $TargetExe -WorkingDirectory (Split-Path -Parent $TargetExe)
} catch {
    # Starting the new build is the point of the exercise; if it will not start,
    # an old build that does is the better outcome.
    if (Test-Path -LiteralPath $TargetExe) { Remove-Item -LiteralPath $TargetExe -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $BackupExe) {
        Move-Item -LiteralPath $BackupExe -Destination $TargetExe -Force -ErrorAction SilentlyContinue
        Start-Process -FilePath $TargetExe -WorkingDirectory (Split-Path -Parent $TargetExe)
    }
    exit 1
}

Remove-Item -LiteralPath $BackupExe -Force -ErrorAction SilentlyContinue
exit 0
"#;

#[cfg(not(target_os = "windows"))]
const HELPER_SCRIPT_BODY: &str = "";

/// Clears what a crashed or interrupted update left behind. Nothing here is
/// load-bearing — the files are named for this purpose and a healthy install
/// has none of them — so every failure is ignored: the previous process may
/// still be holding the backup while it shuts down.
pub fn sweep_leftovers() {
    let Ok(dir) = portable_dir() else { return };
    let Ok(exe) = std::env::current_exe() else { return };
    let Some(name) = exe.file_name().map(|name| name.to_string_lossy().into_owned()) else {
        return;
    };
    for stale in [format!("{name}.old"), format!("{name}.new"), HELPER_SCRIPT.to_string()] {
        let _ = std::fs::remove_file(dir.join(stale));
    }
}

#[tauri::command]
pub fn update_status(app: AppHandle) -> Status {
    Status {
        version: app.package_info().version.to_string(),
        portable: portable_dir().is_ok(),
    }
}

#[tauri::command]
pub async fn update_check(
    app: AppHandle,
    locale: String,
) -> Result<Option<Release>, String> {
    let updater = app
        .updater()
        .map_err(|error| format!("Could not start the updater: {error}"))?;
    let found = updater
        .check()
        .await
        .map_err(|error| format!("Could not check for updates: {error}"))?;
    Ok(found.map(|update| Release {
        version: update.version.clone(),
        notes: localised_notes(&update.raw_json, update.body.as_deref(), &locale),
    }))
}

/// Downloads the release, checks it, and stages the swap. It deliberately does
/// not exit: the window says what is about to happen first, and then asks for
/// the restart that lets the staged executable take over.
#[tauri::command]
pub async fn update_install(app: AppHandle) -> Result<(), String> {
    let dir = portable_dir()?;
    let exe_name = std::env::current_exe()
        .ok()
        .and_then(|exe| exe.file_name().map(|name| name.to_string_lossy().into_owned()))
        .ok_or_else(|| "The running executable has no file name.".to_string())?;

    // Re-checked here rather than trusting a version handed over by the window:
    // the download is only meaningful against what the manifest just announced.
    let updater = app
        .updater()
        .map_err(|error| format!("Could not start the updater: {error}"))?;
    let update = updater
        .check()
        .await
        .map_err(|error| format!("Could not check for updates: {error}"))?
        .ok_or_else(|| "This build is already up to date.".to_string())?;

    // `download` is what verifies the signature, so bytes that arrive are bytes
    // that came from a release signed with this app's key.
    let emitter = app.clone();
    let mut downloaded: u64 = 0;
    let mut announced: Option<u64> = None;
    let archive = update
        .download(
            move |chunk, total| {
                downloaded += chunk as u64;
                let Some(total) = total.filter(|total| *total > 0) else {
                    return;
                };
                // One event per whole percent: the callback fires per chunk, far
                // more often than a window can usefully redraw.
                let percent = downloaded * 100 / total;
                if announced == Some(percent) {
                    return;
                }
                announced = Some(percent);
                let _ = emitter.emit(PROGRESS_EVENT, Progress { downloaded, total });
            },
            || {},
        )
        .await
        .map_err(|error| format!("Could not download the update: {error}"))?;

    let executable = extract_executable(&archive, &update.version)?;
    // The folder's own copies of the notes go stale the moment the executable is
    // swapped, and they are the only thing in it a person reads.
    refresh_documents(&archive, &dir);
    stage_and_launch(&dir, &executable, &exe_name)
}

#[tauri::command]
pub fn update_restart(app: AppHandle) {
    app.exit(0);
}

#[cfg(test)]
mod tests {
    use super::{
        documents_in, extract_executable, is_document, localised_notes, PAYLOAD_EXECUTABLE,
        PAYLOAD_MANIFEST,
    };
    use std::io::{Cursor, Write};

    fn archive(manifest: serde_json::Value, executable: &[u8]) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options =
            zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file(PAYLOAD_MANIFEST, options).unwrap();
        zip.write_all(&serde_json::to_vec(&manifest).unwrap()).unwrap();
        zip.start_file(PAYLOAD_EXECUTABLE, options).unwrap();
        zip.write_all(executable).unwrap();
        zip.finish().unwrap().into_inner()
    }

    fn manifest(version: &str, arch: &str) -> serde_json::Value {
        serde_json::json!({ "schema": 1, "version": version, "arch": arch })
    }

    /// An archive shaped like the workflow's: the payload and its manifest, and
    /// the documents the folder is meant to explain itself with.
    fn release_archive(stamp: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options =
            zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file(PAYLOAD_MANIFEST, options).unwrap();
        zip.write_all(&serde_json::to_vec(&manifest("0.3.0", super::arch_label())).unwrap())
            .unwrap();
        zip.start_file(PAYLOAD_EXECUTABLE, options).unwrap();
        zip.write_all(b"MZpayload").unwrap();
        for name in [
            "README.md",
            "README.ja.md",
            "README.zh-CN.md",
            "CHANGELOG.md",
            "CHANGELOG.ja.md",
            "CHANGELOG.zh-CN.md",
            "LICENSE",
        ] {
            zip.start_file(name, options).unwrap();
            zip.write_all(stamp.as_bytes()).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn an_archive_carries_the_folders_documents_beside_its_payload() {
        let documents = documents_in(&release_archive("0.3.0"));
        let names: Vec<&str> = documents.iter().map(|(name, _)| name.as_str()).collect();
        assert_eq!(
            names,
            [
                "README.md",
                "README.ja.md",
                "README.zh-CN.md",
                "CHANGELOG.md",
                "CHANGELOG.ja.md",
                "CHANGELOG.zh-CN.md",
                "LICENSE",
            ]
        );
        // Contents carried through rather than merely found.
        for (name, bytes) in &documents {
            assert_eq!(bytes, b"0.3.0", "{name}");
        }
    }

    #[test]
    fn the_payload_and_its_manifest_are_not_documents() {
        // The payload has its own extraction, which the signature check reads
        // through, and the manifest describes the archive rather than landing
        // in the folder. Neither should be written out as if it were a file to
        // read.
        assert!(!is_document(PAYLOAD_EXECUTABLE));
        assert!(!is_document(PAYLOAD_MANIFEST));
        assert!(!is_document("apply-update.ps1"));
        assert!(!is_document("portable.sheetlens"));
        // A markdown file that is neither a readme nor a changelog is not ours
        // to overwrite either.
        assert!(!is_document("notes.md"));

        assert!(is_document("README.md"));
        assert!(is_document("CHANGELOG.zh-CN.md"));
        assert!(is_document("LICENSE"));
    }

    #[test]
    fn extracts_the_executable_from_a_matching_archive() {
        let bytes = archive(manifest("0.2.0", super::arch_label()), b"MZpayload");
        let got = extract_executable(&bytes, "0.2.0").unwrap();
        assert_eq!(got, b"MZpayload");
    }

    #[test]
    fn rejects_an_archive_that_announces_a_different_version() {
        let bytes = archive(manifest("0.1.0", super::arch_label()), b"MZpayload");
        let error = extract_executable(&bytes, "0.2.0").unwrap_err();
        assert!(error.contains("was announced"), "{error}");
    }

    #[test]
    fn rejects_an_archive_built_for_another_architecture() {
        let other = if super::arch_label() == "x64" { "arm64" } else { "x64" };
        let bytes = archive(manifest("0.2.0", other), b"MZpayload");
        let error = extract_executable(&bytes, "0.2.0").unwrap_err();
        assert!(error.contains("build, but this is"), "{error}");
    }

    #[test]
    fn rejects_an_archive_whose_payload_is_not_an_executable() {
        let bytes = archive(manifest("0.2.0", super::arch_label()), b"not an executable");
        let error = extract_executable(&bytes, "0.2.0").unwrap_err();
        assert!(error.contains("Windows executable"), "{error}");
    }

    #[test]
    fn rejects_an_archive_with_no_manifest() {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options =
            zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        zip.start_file(PAYLOAD_EXECUTABLE, options).unwrap();
        zip.write_all(b"MZpayload").unwrap();
        let bytes = zip.finish().unwrap().into_inner();

        let error = extract_executable(&bytes, "0.2.0").unwrap_err();
        assert!(error.contains(PAYLOAD_MANIFEST), "{error}");
    }

    #[test]
    fn rejects_an_unknown_manifest_schema() {
        let mut manifest = manifest("0.2.0", super::arch_label());
        manifest["schema"] = serde_json::json!(99);
        let bytes = archive(manifest, b"MZpayload");
        let error = extract_executable(&bytes, "0.2.0").unwrap_err();
        assert!(error.contains("schema 99"), "{error}");
    }

    #[test]
    fn rejects_an_archive_that_is_not_a_zip() {
        let error = extract_executable(b"not a zip at all", "0.2.0").unwrap_err();
        assert!(error.contains("not a readable archive"), "{error}");
    }

    /// The `latest.json` a release with translated notes produces.
    fn translated_manifest() -> serde_json::Value {
        serde_json::json!({
            "version": "0.2.0",
            "notes": "English",
            "notesLocales": { "en": "English", "zh-CN": "中文", "ja": "日本語" },
            "platforms": { "windows-x86_64": { "signature": "s", "url": "https://x/y.zip" } },
        })
    }

    #[test]
    fn reads_the_notes_in_the_language_the_window_asked_for() {
        let notes = localised_notes(&translated_manifest(), Some("English"), "ja");
        assert_eq!(notes.as_deref(), Some("日本語"));
    }

    #[test]
    fn falls_back_to_the_english_notes_for_a_language_the_release_lacks() {
        let notes = localised_notes(&translated_manifest(), Some("English"), "fr");
        assert_eq!(notes.as_deref(), Some("English"));
    }

    #[test]
    fn falls_back_to_the_english_notes_when_the_manifest_carries_no_map() {
        let raw = serde_json::json!({ "version": "0.1.1", "notes": "English" });
        let notes = localised_notes(&raw, Some("English"), "ja");
        assert_eq!(notes.as_deref(), Some("English"));
    }

    #[test]
    fn a_release_with_no_notes_at_all_stays_that_way() {
        let raw = serde_json::json!({ "version": "0.1.1" });
        assert_eq!(localised_notes(&raw, None, "ja"), None);
    }
}
