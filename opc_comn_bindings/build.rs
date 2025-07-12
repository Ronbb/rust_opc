fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=.windows/winmd/OPCCOMN.winmd");

    windows_bindgen::bindgen([
        "--in",
        ".windows/winmd/OPCCOMN.winmd",
        "default",
        "--out",
        "src/bindings.rs",
        "--reference",
        "windows,skip-root,Windows",
        "--filter",
        "OPCCOMN",
        "--flat",
        "--implement",
    ])
    .unwrap();
}
