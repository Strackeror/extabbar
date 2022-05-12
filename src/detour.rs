use std::ffi::c_void;

use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, POINT, S_FALSE, WPARAM};
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};
use windows::Win32::UI::Shell::IShellBrowser;
use windows::Win32::UI::Shell::Common::ITEMIDLIST;
use windows::Win32::UI::WindowsAndMessaging::{RegisterWindowMessageW, SendMessageW};
use windows::core::{HRESULT, IUnknown, GUID, Interface};

use crate::{BROWSE_OBJECT_MESSAGE, SHOW_WINDOW_MESSAGE};

static mut MESSAGE_ID_BROWSE_OBJECT: u32 = 0;
static mut DETOUR_BROWSE_OBJECT: Option<detour::RawDetour> = None;

type BrowseObjectFn =
    unsafe extern "system" fn(this: *mut c_void, pidl: *const ITEMIDLIST, w_flags: u32) -> HRESULT;

unsafe extern "system" fn browse_object_detour(
    this: *mut c_void,
    pidl: *const ITEMIDLIST,
    mut w_flags: u32,
) -> HRESULT {
    let shell_browser = &this as *const _ as *mut IShellBrowser;
    if let Ok(window) = (*shell_browser).GetWindow() {
        let res = SendMessageW(
            window,
            MESSAGE_ID_BROWSE_OBJECT,
            WPARAM(&mut w_flags as *mut _ as _),
            LPARAM(pidl as _),
        );

        if res == LRESULT(1) {
            return S_FALSE;
        }
    }

    let base: BrowseObjectFn =
        std::mem::transmute(DETOUR_BROWSE_OBJECT.as_ref().unwrap().trampoline());
    base(this, pidl, w_flags)
}

pub unsafe fn hook_browse_object(browser: IShellBrowser) {
    if DETOUR_BROWSE_OBJECT.is_some() {
        return;
    }

    MESSAGE_ID_BROWSE_OBJECT = RegisterWindowMessageW(BROWSE_OBJECT_MESSAGE);
    log::info!("hook browse object {:?}", MESSAGE_ID_BROWSE_OBJECT);

    DETOUR_BROWSE_OBJECT =
        detour::RawDetour::new(browser.vtable().BrowseObject as _, browse_object_detour as _)
            .map_err(|op| {
                log::error!("error hook:{:?}", &op);
                op
            })
            .ok();
    DETOUR_BROWSE_OBJECT
        .as_ref()
        .unwrap()
        .enable()
        .expect("failed to enable hook");

    log::info!("hook status: {:?}", DETOUR_BROWSE_OBJECT);
}

static mut DETOUR_SHOW_WINDOW: Option<detour::RawDetour> = None;
static mut SHOW_WINDOW_EXPLORER_HANDLE: Option<HWND> = None;
static mut SHOW_WINDOW_MESSAGE_ID: u32 = 0;

type ShowWindowFn = unsafe extern "system" fn(
    this: *mut c_void,
    pidl: *const ITEMIDLIST,
    flags: u32,
    pt: POINT,
    unkn: u32,
    unkn2: u64,
    unkn3: u64,
    unkn4: u64,
) -> HRESULT;

unsafe extern "system" fn show_window_detour(
    this: *mut c_void,
    pidl: *const ITEMIDLIST,
    mut flags: u32,
    pt: POINT,
    unkn: u32,
    unkn2: u64,
    unkn3: u64,
    unkn4: u64,
) -> HRESULT {
    let base: ShowWindowFn = std::mem::transmute(DETOUR_SHOW_WINDOW.as_ref().unwrap().trampoline());
    let handle = SHOW_WINDOW_EXPLORER_HANDLE.unwrap();
    let result = SendMessageW(
        handle,
        SHOW_WINDOW_MESSAGE_ID,
        WPARAM(&mut flags as *mut _ as usize),
        LPARAM(pidl as _),
    )
    .0;

    log::info!(
        "Show new window hook {:?}",
        (
            ("handle", handle,),
            this,
            pidl,
            flags,
            pt,
            &base as *const _,
            result
        )
    );

    match result {
        0 => {
            log::info!("Creating new window");
            base(this, pidl, flags, pt, unkn, unkn2, unkn3, unkn4)
        }
        _ => HRESULT::default(),
    }
}

// From QTTabBar
//MIDL_INTERFACE("489E9453-869B-4BCC-A1C7-48B5285FD9D8") ICommonExplorerHost  : public IUnknown {};
//MIDL_INTERFACE("93A56381-E0CD-485A-B60E-67819E12F81B") CExplorerFactoryServer {};
pub unsafe fn hook_show_window(explorer_handle: HWND) {
    SHOW_WINDOW_EXPLORER_HANDLE = Some(explorer_handle);
    if DETOUR_SHOW_WINDOW.is_some() {
        DETOUR_SHOW_WINDOW.as_ref().unwrap().enable().unwrap();
        return;
    }

    SHOW_WINDOW_MESSAGE_ID = RegisterWindowMessageW(SHOW_WINDOW_MESSAGE);

    let explorer_factory_server_clsid = GUID::from("93A56381-E0CD-485A-B60E-67819E12F81B");
    let instance: IUnknown = CoCreateInstance(
        std::ptr::addr_of!(explorer_factory_server_clsid),
        None,
        CLSCTX_INPROC_SERVER,
    )
    .expect("failed to get explorer factory server");

    // Need to find a clearer way to do this
    // We need to hook the function at index 3 in the vtable
    let out: *const *const *const c_void = std::mem::transmute_copy(&instance);
    let out = (*out).add(3);
    DETOUR_SHOW_WINDOW = Some(
        detour::RawDetour::new(*out as _, show_window_detour as _).expect("failed to create hook"),
    );
    DETOUR_SHOW_WINDOW
        .as_ref()
        .unwrap()
        .enable()
        .expect("failed to enable hook")
}
