use std::io::Read;

use winreg::{
    enums::{HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE},
    RegKey,
};

// {9ecce421-925a-4484-b2cf-c00b182bc32a}
const EXT_TAB_GUID: &str = "{9ecce421-925a-4484-b2cf-c00b182bc32a}";
fn main() -> std::io::Result<()> {
    let hkcu = RegKey::predef(HKEY_LOCAL_MACHINE);

    println!("Install/Uninstall ? (i/u)");
    let mut input = String::new();
    let sz = std::io::stdin().read_line(&mut input).unwrap();

    let uninstall = match input.trim() {
        "u" => true,
        "i" => false,
        _ => panic!("Invalid input"),
    };

    let clsid = hkcu.open_subkey("Software\\Classes\\CLSID").unwrap();
    let (toolbar, _) = hkcu
        .create_subkey("Software\\Microsoft\\Internet Explorer\\Toolbar")
        .unwrap();
    if uninstall {
        clsid.delete_subkey_all(EXT_TAB_GUID).unwrap();
        toolbar.delete_value(EXT_TAB_GUID).unwrap();
    } else {
        let clsid = hkcu.open_subkey("Software\\Classes\\CLSID").unwrap();

        let (class, _) = clsid.create_subkey(EXT_TAB_GUID).unwrap();
        class.set_value("", &"exttabbar").unwrap();
        let (inproc, _) = class.create_subkey("InProcServer32").unwrap();
        inproc.set_value("ThreadingModel", &"Apartment").unwrap();
        inproc
            .set_value(
                "",
                &std::env::current_exe()
                    .unwrap()
                    .with_file_name("extabbar.dll")
                    .as_os_str(),
            )
            .unwrap();

        toolbar.set_value(EXT_TAB_GUID, &"extabbar").unwrap();
    }

    Ok(())
}
