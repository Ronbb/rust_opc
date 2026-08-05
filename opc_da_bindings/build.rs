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
    ])
    .unwrap();

    remove_shallow_clones(&[
        "tagOPCITEMATTRIBUTES",
        "tagOPCITEMPROPERTY",
        "tagOPCITEMSTATE",
        "tagOPCITEMVQT",
    ]);
}

/// `windows-bindgen` currently emits `transmute_copy` Clone implementations for
/// explicitly laid-out structs containing `VARIANT`. A bitwise clone aliases the
/// BSTR/SAFEARRAY/interface owned by the variant and causes a double release.
fn remove_shallow_clones(type_names: &[&str]) {
    let path = "src/bindings.rs";
    let mut source = std::fs::read_to_string(path).unwrap();

    for type_name in type_names {
        let implementation = format!(
            "impl Clone for {type_name} {{\n    fn clone(&self) -> Self {{\n        unsafe {{ core::mem::transmute_copy(self) }}\n    }}\n}}\n"
        );
        source = source.replace(&implementation, "");
        assert!(
            !source.contains(&format!("impl Clone for {type_name}")),
            "unsafe shallow Clone implementation remains for {type_name}"
        );
    }

    std::fs::write(path, source).unwrap();
}
