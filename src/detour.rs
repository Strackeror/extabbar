use std::{
    ffi::c_void,
    sync::atomic::{AtomicU32, Ordering},
};

use bindings::Windows::Win32::UI::Shell::{IShellBrowser, ITEMIDLIST};
use windows::*;

static mut MESSAGE_ID_BROWSE_OBJECT: u32 = 0;
static mut DETOUR_BROWSE_OBJECT: Option<detour::RawDetour> = None;

type BrowseObjectFn =
    unsafe extern "stdcall" fn(this: *mut c_void, pidl: *const ITEMIDLIST, wFlags: u32) -> HRESULT;

unsafe extern "stdcall" fn browse_object_hook(
    this: *mut c_void,
    pidl: *const ITEMIDLIST,
    wFlags: u32,
) -> HRESULT {
    log::info!("browse object {:?} {:#x?}", pidl, wFlags);
    let base: BrowseObjectFn =
        std::mem::transmute(DETOUR_BROWSE_OBJECT.as_ref().unwrap().trampoline());
    base(this, pidl, wFlags)
}

pub unsafe fn hook_browse_object(browser: IShellBrowser, message_id: u32) {
    log::info!("hook browse object {:?}", message_id);
    if MESSAGE_ID_BROWSE_OBJECT != message_id {
        MESSAGE_ID_BROWSE_OBJECT = message_id;
        //let mut hook = BrowseObjectDetour::new(hooked_function, browse_object_hook);

        DETOUR_BROWSE_OBJECT =
            detour::RawDetour::new(browser.vtable().11 as _, browse_object_hook as _)
                .map_err(|op| {
                    log::error!("error hook:{:?}", &op);
                    op
                })
                .ok();
        DETOUR_BROWSE_OBJECT.as_ref().unwrap().enable();
    }
    log::info!("hook status: {:?}", DETOUR_BROWSE_OBJECT);
}
