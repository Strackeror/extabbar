use std::borrow::BorrowMut;
use std::cell::RefCell;
use std::ffi::CString;
use std::panic;
use std::path::PathBuf;
use std::rc::Rc;

use bindings::Windows::Win32::Foundation::*;
use bindings::Windows::Win32::UI::Controls::*;
use bindings::Windows::Win32::UI::WindowsAndMessaging::*;
use windows::Result;

pub static mut DLL_INSTANCE: Option<HINSTANCE> = None;

pub struct Tab {
    path: Option<PathBuf>,
}

#[derive(Default)]
pub struct TabBar(Rc<TabBarRef>);

#[derive(Default)]
struct TabBarRef(RefCell<Obj>);

#[derive(Default)]
struct Obj {
    handle: HWND,

    tabs: Vec<Tab>,
    default_window_procedure: Option<WNDPROC>,
}

pub extern "system" fn tab_bar_proc(
    hwnd: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let obj_ptr = unsafe {
        match GetWindowLongPtrA(hwnd, GWLP_USERDATA) {
            0 => panic!("Could not get user data"),
            n => (n as *mut TabBarRef),
        }
    };

    let obj = unsafe { &(*obj_ptr) };
    obj.window_procedure(hwnd, message, wparam, lparam)
}

impl TabBar {
    pub fn new(parent: HWND) -> TabBar {
        let handle = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE(0),
                "SysTabControl32",
                "",
                WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
                0,
                0,
                0,
                0,
                parent,
                HMENU(0),
                DLL_INSTANCE.unwrap(),
                std::ptr::null(),
            )
        };

        log::warn!("TabBar created {:?}", handle);
        let ret = TabBar(Rc::new(TabBarRef(RefCell::new(Obj {
            handle,
            ..Default::default()
        }))));

        unsafe {
            let mut obj = (*ret.0).0.borrow_mut();

            SetWindowLongPtrA(handle, GWLP_USERDATA, Rc::as_ptr(&ret.0) as isize);
            let default_proc = SetWindowLongPtrW(
                handle,
                GWLP_WNDPROC,
                tab_bar_proc as WNDPROC as usize as isize,
            );
            if let 0 = default_proc {
                panic!("Set wndproc failed");
            }

            log::info!("storing");
            obj.default_window_procedure = Some(std::mem::transmute(default_proc));
            log::info!("stored default proc 0x{:x}", default_proc);
        }

        ret
    }

    pub fn get_handle(&self) -> HWND {
        return self.0 .0.borrow().handle;
    }

    pub fn add_tab(&self, title: CString, idx: usize) -> Result<()> {
        self.0.add_tab(title, idx)
    }
}

impl TabBarRef {
    fn add_tab(&self, title: CString, idx: usize) -> Result<()> {
        let handle = self.0.borrow().handle;
        let tab_info = TCITEMA {
            mask: TCIF_TEXT,
            pszText: PSTR(title.into_raw() as *mut u8),
            lParam: LPARAM(0),
            ..Default::default()
        };
        let result = unsafe {
            SendMessageA(
                handle,
                TCM_INSERTITEMA,
                WPARAM(idx),
                LPARAM(std::ptr::addr_of!(tab_info) as isize),
            )
        };

        log::info!("Added tab done, result:{:?}", result);
        if result.0 < 0 {
            return Err(E_FAIL.into());
        }
        Ok(())
    }

    fn set_tab_title(&self, title: String, idx: usize) -> Result<()> {
        let handle = self.0.borrow().handle;

        let tab_info = TCITEMA {
            mask: TCIF_TEXT,
            pszText: PSTR(CString::new(title).expect("invalid title").into_raw() as *mut u8),
            ..Default::default()
        };

        let result = unsafe {
            SendMessageA(
                handle,
                TCM_INSERTITEMA,
                WPARAM(idx),
                LPARAM(std::ptr::addr_of!(tab_info) as isize),
            )
        };
        match result.0 {
            0 => Err(E_FAIL.into()),
            _ => Ok(()),
        }
    }

    fn get_tab_count(&self) -> i32 {
        let handle = self.0.borrow().handle;
        unsafe { SendMessageA(handle, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0)).0 }
    }

    fn get_selected_tab_index(&self) -> Option<i32> {
        let handle = self.0.borrow().handle;
        unsafe { Some(SendMessageA(handle, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0)).0) }
    }

    fn remove_tab(&self, idx: usize) -> Result<()> {
        let handle = self.0.borrow().handle;
        unsafe {
            match SendMessageA(handle, TCM_DELETEITEM, WPARAM(idx), LPARAM(0)) {
                LRESULT(0) => Err(E_FAIL.into()),
                _ => Ok(()),
            }
        }
    }

    fn window_procedure(
        &self,
        hwnd: HWND,
        message: u32,
        wparam: WPARAM,
        lparam: LPARAM,
    ) -> LRESULT {
        LRESULT(match message {
            WM_MBUTTONDOWN => self.add_tab(CString::new("added").unwrap(), 0).is_ok() as i32,
            WM_RBUTTONDOWN => self.remove_tab(0).is_ok() as i32,
            _ => {
                let proc = self.0.borrow().default_window_procedure;
                unsafe { proc.unwrap()(hwnd, message, wparam, lparam).0 }
            }
        })
    }
}
