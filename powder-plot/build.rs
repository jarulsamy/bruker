use {std::io, winresource::WindowsResource};

fn main() -> io::Result<()> {
    if cfg!(target_os = "windows") {
        WindowsResource::new()
            .set_icon("assets/icon.ico")
            .set("FileDescription", env!("CARGO_PKG_DESCRIPTION"))
            .set("ProductName", env!("CARGO_PKG_NAME"))
            .set("FileVersion", env!("CARGO_PKG_VERSION"))
            .set("ProductVersion", env!("CARGO_PKG_VERSION"))
            .set("LegalCopyright", &format!("© 2025 {}", env!("CARGO_PKG_AUTHORS")))
            .compile()?;
    }
    Ok(())
}
