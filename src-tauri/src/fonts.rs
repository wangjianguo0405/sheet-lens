//! The font families this machine has, for the ribbon's family list.
//!
//! The webview cannot answer this: `document.fonts` only knows the faces a page
//! loaded, and the Local Font Access API has no permission path in WebView2.
//! GDI can, and it is the same set the canvas draws with.

/// Every installed font family, once each and sorted the way a list reads.
#[tauri::command(async)]
pub fn list_fonts() -> Vec<String> {
    let mut names = platform::families();
    names.sort_by_cached_key(|name| name.to_lowercase());
    // GDI reports a family once per style it has, so the same name arrives
    // several times.
    names.dedup_by(|a, b| a.eq_ignore_ascii_case(b));
    names
}

#[cfg(windows)]
mod platform {
    use windows::Win32::Foundation::LPARAM;
    use windows::Win32::Graphics::Gdi::{
        EnumFontFamiliesExW, GetDC, ReleaseDC, DEFAULT_CHARSET, LOGFONTW, TEXTMETRICW,
        TRUETYPE_FONTTYPE,
    };

    pub fn families() -> Vec<String> {
        let mut names: Vec<String> = Vec::new();
        unsafe {
            let hdc = GetDC(None);
            if hdc.is_invalid() {
                return names;
            }
            // An empty face asks for every family; the charset is the only
            // filter, and "any" is what a document may name.
            let filter = LOGFONTW {
                lfCharSet: DEFAULT_CHARSET,
                ..Default::default()
            };
            let _ = EnumFontFamiliesExW(
                hdc,
                &filter,
                Some(collect),
                LPARAM(&mut names as *mut Vec<String> as isize),
                0,
            );
            ReleaseDC(None, hdc);
        }
        names
    }

    unsafe extern "system" fn collect(
        logfont: *const LOGFONTW,
        _metric: *const TEXTMETRICW,
        font_type: u32,
        names: LPARAM,
    ) -> i32 {
        // Bitmap and vector faces are not families a document can name.
        if font_type & TRUETYPE_FONTTYPE == 0 {
            return 1;
        }
        let face = &(*logfont).lfFaceName;
        let len = face.iter().position(|&ch| ch == 0).unwrap_or(face.len());
        let name = String::from_utf16_lossy(&face[..len]);
        // An '@' name is the vertical-writing twin of a face, not a family of
        // its own — Excel does not offer them either.
        if !name.is_empty() && !name.starts_with('@') {
            (*(names.0 as *mut Vec<String>)).push(name);
        }
        1
    }
}

#[cfg(not(windows))]
mod platform {
    pub fn families() -> Vec<String> {
        Vec::new()
    }
}

#[cfg(all(test, windows))]
mod tests {
    /// The machine this runs on has fonts, and the list comes back as a list:
    /// sorted, once each, and without the vertical-writing twins.
    #[test]
    fn reads_as_a_list() {
        let names = super::list_fonts();
        assert!(names.len() > 10, "only {} families", names.len());
        assert!(names
            .iter()
            .all(|name| !name.is_empty() && !name.starts_with('@')));
        let mut sorted = names.clone();
        sorted.sort_by_cached_key(|name| name.to_lowercase());
        assert_eq!(names, sorted, "not sorted");
        let mut unique = names.clone();
        unique.dedup_by(|a, b| a.eq_ignore_ascii_case(b));
        assert_eq!(names.len(), unique.len(), "duplicates in {}", names.len());
    }
}
