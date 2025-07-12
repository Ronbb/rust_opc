fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=.windows/winmd/OPCAE.winmd");

    windows_bindgen::bindgen([
        "--in",
        ".windows/winmd/OPCAE.winmd",
        "default",
        "--out",
        "src/bindings.rs",
        "--reference",
        "windows,skip-root,Windows",
        "--filter",
        "OPCAE",
        "--flat",
        "--implement",
    ])
    .unwrap();
}
