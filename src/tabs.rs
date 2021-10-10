use std::borrow::Borrow;
use std::cell::Ref;
use std::cell::RefCell;
use std::cell::RefMut;
use std::collections::HashMap;
use std::panic;
use std::path::Path;
use std::path::PathBuf;
use std::rc::Rc;

use bindings::Windows::Win32::Foundation::*;
use bindings::Windows::Win32::UI::Controls::*;
use bindings::Windows::Win32::UI::WindowsAndMessaging::*;
use windows::Error;
use windows::Result;

pub static mut DLL_INSTANCE: Option<HINSTANCE> = None;

#[derive(Clone, Default, Debug)]
pub struct Tab {
    path: Option<PathBuf>,
}

#[derive(Default, Clone)]
pub struct TabBar(Rc<TabBarRef>);

#[derive(Default)]
struct TabBarRef(RefCell<Obj>);

#[derive(Default)]
struct Obj {
    handle: HWND,

    tabs: HashMap<i32, Tab>,
    tab_key_counter: i32,

    default_window_procedure: Option<WNDPROC>,
}

pub extern "system" fn tab_bar_proc(
    hwnd: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    let obj_ptr = unsafe {
        match GetWindowLongPtrW(hwnd, GWLP_USERDATA) {
            0 => panic!("Could not get user data"),
            n => (n as *mut TabBarRef),
        }
    };

    let obj = unsafe { &(*obj_ptr) };
    obj.window_procedure(hwnd, message, wparam, lparam)
}

fn get_tab_name(path: &String) -> Result<String> {
    let path = Path::new(&path);
    let file_name = path.file_name().ok_or(E_FAIL)?;
    let file_name = file_name.to_owned().into_string();
    let file_name = file_name.map_err(|_| Error::from(E_FAIL))?;
    Ok(file_name)
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

            SetWindowLongPtrW(handle, GWLP_USERDATA, Rc::as_ptr(&ret.0) as isize);
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

    pub fn add_tab(&self, path: String, idx: usize) -> Result<()> {
        let key: i32;
        {
            let obj = &mut *self.0 .0.borrow_mut();
            let tabs = &mut obj.tabs;
            let key_counter = &mut obj.tab_key_counter;

            key = *key_counter;
            *key_counter += 1;
            tabs.insert(
                key,
                Tab {
                    path: Some(Path::new(&path).to_owned()),
                },
            );
        }
        self.0.add_tab(path, idx, key)
    }

    pub fn navigated(&self, path: String) -> Result<()> {
        if let Some(index) = self.0.get_selected_tab_index() {
            self.0.get_tab(index)?.path = Some(Path::new(&path).to_owned());
            self.0.set_tab_title(get_tab_name(&path)?, index as usize)?;
        }
        Ok(())
    }

    pub fn new_window(&self, path: String) -> Result<()> {
        self.add_tab(path, self.0.get_tab_count() as usize)
    }
}

impl TabBarRef {
    fn add_tab(&self, title: String, idx: usize, key: i32) -> Result<()> {
        let handle = self.0.borrow().handle;

        let mut text: Vec<_> = title.encode_utf16().collect();
        text.push(0);
        let text = PWSTR(Box::<[_]>::into_raw(text.into_boxed_slice()) as _);

        let tab_info = TCITEMW {
            mask: TCIF_TEXT | TCIF_PARAM,
            pszText: text,
            lParam: LPARAM(key as isize),
            ..Default::default()
        };
        let result = unsafe {
            SendMessageW(
                handle,
                TCM_INSERTITEMW,
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

        let mut text: Vec<_> = title.encode_utf16().collect();
        text.push(0);
        log::info!("Setting tab no {} to {:?}", idx, text);
        let text = PWSTR(Box::<[_]>::into_raw(text.into_boxed_slice()) as _);
        let tab_info = TCITEMW {
            mask: TCIF_TEXT,
            pszText: text,
            ..Default::default()
        };

        let result = unsafe {
            SendMessageW(
                handle,
                TCM_SETITEMW,
                WPARAM(idx),
                LPARAM(std::ptr::addr_of!(tab_info) as isize),
            )
        };
        let res = match result.0 {
            0 => Err(E_FAIL.into()),
            _ => Ok(()),
        };
        log::info!("result {:?}", res);
        res
    }

    fn get_tab_count(&self) -> i32 {
        let handle = self.0.borrow().handle;
        unsafe { SendMessageW(handle, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0)).0 }
    }

    fn get_selected_tab_index(&self) -> Option<usize> {
        let handle = self.0.borrow().handle;
        unsafe { Some(SendMessageW(handle, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as usize) }
    }

    fn get_tab(&self, index: usize) -> Result<RefMut<Tab>> {
        let mut tab_info: TCITEMW = Default::default();
        let handle = self.0.borrow().handle;
        let key = unsafe {
            match SendMessageW(
                handle,
                TCM_GETITEMW,
                WPARAM(index),
                LPARAM(std::ptr::addr_of_mut!(tab_info) as isize),
            ) {
                LRESULT(0) => return Err(E_FAIL.into()),
                _ => tab_info.lParam.0,
            }
        };

        let key = key as i32;

        if !self.0.borrow().tabs.contains_key(&key) {
            return Err(E_FAIL.into());
        }
        let tab_ref = RefMut::map(self.0.borrow_mut(), |obj| obj.tabs.get_mut(&key).unwrap());
        Ok(tab_ref)
    }

    fn remove_tab(&self, idx: usize) -> Result<()> {
        let handle = self.0.borrow().handle;
        unsafe {
            match SendMessageW(handle, TCM_DELETEITEM, WPARAM(idx), LPARAM(0)) {
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
            WM_RBUTTONDOWN => self.remove_tab(0).is_ok() as i32,
            _ => {
                let proc = self.0.borrow().default_window_procedure;
                unsafe { proc.unwrap()(hwnd, message, wparam, lparam).0 }
            }
        })
    }
}
