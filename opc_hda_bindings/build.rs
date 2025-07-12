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
        "--implement",
    ])
    .unwrap();
}
