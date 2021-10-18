use std::rc::Weak;

use bindings::Windows::Win32::{
    Foundation::*,
    UI::{
        Shell::{DefSubclassProc, RemoveWindowSubclass, SetWindowSubclass},
        WindowsAndMessaging::RegisterWindowMessageW,
    },
};

use crate::SHOW_WINDOW_MESSAGE;

use super::tab_bar::TabBar;

pub struct ExplorerSubclass {
    explorer_handle: HWND,
    tab_bar: Weak<TabBar>,

    show_window_message_id: u32,
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
        });

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
            let should_open = self
                .tab_bar
                .upgrade()
                .unwrap()
                .new_window(Some(lparam.0 as _))
                .is_ok();
            return LRESULT(should_open as _);
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
