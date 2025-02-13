fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=.windows/winmd/OPCDA.winmd");

    windows_bindgen::bindgen([
        "--in",
        ".windows/winmd/OPCDA.winmd",
        "default",
        "--out",
        "src/bindings.rs",
        "--reference",
        "windows,skip-root,Windows",
        "--filter",
        "OPCDA",
        "--flat",
        "--implement",
    ]);
}
