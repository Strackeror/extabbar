use std::cell::RefCell;
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
pub struct TabBar(pub Rc<RefCell<Obj>>);

#[derive(Default)]
pub struct Obj {
    pub handle: HWND,

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
            n => (n as *mut Obj),
        }
    };

    log::info!("received message {:x}", message);
    if message == WM_MBUTTONDOWN {
        log::info!("received mouse press");
        unsafe {
            (*obj_ptr)
                .add_tab("added\0".to_owned().as_mut_str(), 5)
                .unwrap();
        }
    }

    unsafe {
        match (*obj_ptr).default_window_procedure {
            None => panic!("Could not get wndproc"),
            Some(wndproc) => wndproc(hwnd, message, wparam, lparam),
        }
    }
}

impl TabBar {
    pub fn new(parent: HWND) -> TabBar {
        let handle = unsafe {
            CreateWindowExA(
                WINDOW_EX_STYLE(0),
                "SysTabControl32",
                "",
                WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
                0,
                0,
                200,
                25,
                parent,
                HMENU(0),
                DLL_INSTANCE.unwrap(),
                std::ptr::null(),
            )
        };

        log::warn!("TabBar created {:?}", handle);
        let ret = TabBar(Rc::new(RefCell::new(Obj {
            handle,
            ..Default::default()
        })));

        unsafe {
            let mut obj = (*ret.0).borrow_mut();
            SetWindowLongPtrA(handle, GWLP_USERDATA, (*ret.0).as_ptr() as isize);
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
}

impl Obj {
    pub fn add_tab(&mut self, title: &mut str, idx: usize) -> Result<()> {
        let tab_info = TCITEMA {
            mask: TCIF_STATE | TCIF_TEXT,
            dwState: 0,
            dwStateMask: TCIS_HIGHLIGHTED,
            pszText: PSTR(title.as_mut_ptr()),
            cchTextMax: 0,
            iImage: 0,
            lParam: LPARAM(0),
        };
        let result = unsafe {
            SendMessageA(
                &self.handle,
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

    pub fn remove_tab(&mut self, idx: usize) -> Result<()> {
        unsafe {
            match SendMessageA(&self.handle, TCM_DELETEITEM, WPARAM(idx), LPARAM(0)) {
                LRESULT(0) => Err(E_FAIL.into()),
                _ => Ok(()),
            }
        }
    }
}
