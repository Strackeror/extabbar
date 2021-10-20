use std::rc::Weak;

use bindings::Windows::Win32::{
    Foundation::*,
    UI::{
        Shell::{
            DefSubclassProc, RemoveWindowSubclass, SetWindowSubclass, SBSP_NAVIGATEBACK,
            SBSP_NAVIGATEFORWARD,
        },
        WindowsAndMessaging::RegisterWindowMessageW,
    },
};

use crate::{idl::Idl, BROWSE_OBJECT_MESSAGE, SHOW_WINDOW_MESSAGE};

use super::tab_bar::TabBar;

pub struct ExplorerSubclass {
    explorer_handle: HWND,
    tab_bar: Weak<TabBar>,

    show_window_message_id: u32,
    browse_object_message_id: u32,
}

impl ExplorerSubclass {
    pub extern "system" fn subclass_proc(
        hwnd: HWND,
        message: u32,
        wparam: WPARAM,
        lparam: LPARAM,
        _uid_subclass: usize,
        ref_data: usize,
    ) -> LRESULT {
        let obj = ref_data as *const ExplorerSubclass;
        let obj = unsafe { obj.as_ref().unwrap() };
        obj.window_procedure(hwnd, message, wparam, lparam)
    }

    const SUBCLASS_UID: usize = 43;

    pub fn new(explorer_handle: HWND, tab_bar: Weak<TabBar>) -> Box<Self> {
        let new = Box::new(Self {
            explorer_handle,
            tab_bar,
            show_window_message_id: unsafe { RegisterWindowMessageW(SHOW_WINDOW_MESSAGE) },
            browse_object_message_id: unsafe { RegisterWindowMessageW(BROWSE_OBJECT_MESSAGE) },
        });

        log::info!(
            "new explorer subclass browse_object_id:{}",
            new.browse_object_message_id
        );

        unsafe {
            SetWindowSubclass(
                explorer_handle,
                Some(Self::subclass_proc),
                Self::SUBCLASS_UID,
                &*new as *const _ as usize,
            )
            .as_bool()
            .then(|| ())
            .expect("failed to install subclass");
        }
        new
    }

    fn window_procedure(
        &self,
        hwnd: HWND,
        message: u32,
        wparam: WPARAM,
        lparam: LPARAM,
    ) -> LRESULT {
        if message == self.show_window_message_id {
            let block_open = self
                .tab_bar
                .upgrade()
                .unwrap()
                .new_window(Some(Idl::new(lparam.0 as _)))
                .is_ok();
            return LRESULT(block_open as _);
        }
        if message == self.browse_object_message_id {
            let flags = wparam.0 as *mut u32;
            let flags = unsafe { *flags };
            if flags & SBSP_NAVIGATEBACK != 0 {
                log::debug!("received navigate back");
                let _ = self.tab_bar.upgrade().unwrap().navigate_back();
                return LRESULT(1);
            }
            if flags & SBSP_NAVIGATEFORWARD != 0 {
                log::debug!("received navigate forward");
                let _ = self.tab_bar.upgrade().unwrap().navigate_forward();
                return LRESULT(1);
            }
            return LRESULT(0);
        }
        unsafe { DefSubclassProc(hwnd, message, wparam, lparam) }
    }
}

impl Drop for ExplorerSubclass {
    fn drop(&mut self) {
        log::info!("dropping explorer subclass");
        unsafe {
            RemoveWindowSubclass(
                self.explorer_handle,
                Some(Self::subclass_proc),
                Self::SUBCLASS_UID,
            )
            .ok()
            .expect("failed to drop explorer subclass")
        }
    }
}
