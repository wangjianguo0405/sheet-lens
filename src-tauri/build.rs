fn main() {
    // `tauri_build::build()` embeds `icons/icon.ico` into the Windows binary, but
    // it only asks cargo to re-run this script when `tauri.conf.json` changes.
    // Replacing the icon set therefore changed the PNGs and nothing else: cargo
    // found the crate fresh, the resource compiler never ran again, and the app
    // kept the previous icon until some unrelated edit forced a rebuild.
    println!("cargo:rerun-if-changed=icons");
    tauri_build::build()
}
