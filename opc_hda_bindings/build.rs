fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=.windows/winmd/OPCHDA.winmd");

    windows_bindgen::bindgen([
        "--in",
        ".windows/winmd/OPCHDA.winmd",
        "default",
        "--out",
        "src/bindings.rs",
        "--reference",
        "windows,skip-root,Windows",
        "--filter",
        "OPCHDA",
        "--flat",
    ])
    .unwrap();

    fix_bindgen_formatting();
}

/// `windows-bindgen` 0.66 emits one continuation with an extra indentation
/// level. Normalize it so a clean regeneration remains `rustfmt --check` clean.
fn fix_bindgen_formatting() {
    let path = "src/bindings.rs";
    let source = std::fs::read_to_string(path).unwrap();
    let source = source.replace(
        "    fn GetItemID(&self, sznode: &windows_core::PCWSTR)\n        -> windows_core::Result<windows_core::PWSTR>;",
        "    fn GetItemID(&self, sznode: &windows_core::PCWSTR)\n    -> windows_core::Result<windows_core::PWSTR>;",
    );
    std::fs::write(path, source).unwrap();
}
