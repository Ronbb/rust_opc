fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=.windows/winmd/OPCDA.winmd");

    windows_bindgen::bindgen([
        "--in",
        ".windows/winmd/OPCDA.winmd",
        "--out",
        "src/opcda/bindings.rs",
        "--filter",
        "OPCDA",
        "--config",
        "implement",
    ])
    .unwrap();
}
