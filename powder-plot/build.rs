use {std::io, winresource::WindowsResource};

fn main() -> io::Result<()> {
    if cfg!(target_os = "windows") {
        WindowsResource::new()
            .set_icon("assets/icon.ico")
            .compile()?;
    }
    Ok(())
}
