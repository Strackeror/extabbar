use std::cell::RefCell;
use std::cell::RefMut;
use std::collections::HashMap;
use std::panic;
use std::rc::Rc;

use bindings::Windows::Win32::Foundation::*;
use bindings::Windows::Win32::UI::Controls::*;
use bindings::Windows::Win32::UI::Shell::{
    IShellBrowser, SHGetNameFromIDList, ITEMIDLIST, SIGDN_NORMALDISPLAY,
};
use bindings::Windows::Win32::UI::WindowsAndMessaging::*;
use windows::Result;

pub static mut DLL_INSTANCE: Option<HINSTANCE> = None;

pub type TabPath = *mut ITEMIDLIST;

#[derive(Clone)]
pub struct Tab {
    path: TabPath,
}

#[derive(Clone)]
pub struct TabBar(Rc<TabBarRef>);

struct TabBarRef(RefCell<Obj>);

struct Obj {
    handle: HWND,

    tabs: HashMap<usize, Tab>,
    tab_key_counter: usize,

    default_window_procedure: Option<WNDPROC>,
    explorer: IShellBrowser,
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

unsafe fn pwstr_to_string(pwstr: PWSTR) -> Result<String> {
    if pwstr.0.is_null() {
        return Err(E_FAIL.into());
    }

    let mut end = pwstr.0;
    while *end != 0 {
        end = end.add(1);
    }

    let result = String::from_utf16_lossy(std::slice::from_raw_parts(
        pwstr.0,
        end.offset_from(pwstr.0) as _,
    ));

    Ok(result)
}

fn get_tab_name(pidl: &TabPath) -> String {
    unsafe {
        let name = SHGetNameFromIDList(*pidl, SIGDN_NORMALDISPLAY);
        let name = match name {
            Ok(name) => pwstr_to_string(name),
            Err(_) => return String::new(),
        };
        name.unwrap_or_else(|_| "???".to_owned())
    }

    /*
    let mut buffer = [0u16; 256];
    let path = PWSTR(std::ptr::addr_of_mut!(buffer[0]));
    unsafe {
        if !SHGetPathFromIDListW(std::ptr::addr_of!(*pidl), path).as_bool() {
            log::error!("could not get path from idlist");
        }
    }
    let path = String::from_utf16(&buffer).unwrap_or_else(|_| "???".to_owned());

    let file_name = Path::new(&path)
        .file_name()
        .unwrap_or_default()
        .to_owned()
        .into_string()
        .unwrap_or_default();
    file_name
    */

    /*
    file_name.or(Ok(path));

    let path = Path::new(&path);
    let file_name = path.file_name().ok_or(E_FAIL)?;
    let file_name = file_name.to_owned().into_string();
    let file_name = file_name.map_err(|_| Error::from(E_FAIL))?;
    Ok(file_name)
    */
}

impl TabBar {
    pub fn new(parent: HWND, browser: IShellBrowser) -> TabBar {
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
            tabs: Default::default(),
            tab_key_counter: 0,
            default_window_procedure: None,
            explorer: browser,
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

    pub fn add_tab(&self, path: TabPath, idx: usize) -> Result<()> {
        self.0.add_tab(path, idx)
    }

    pub fn navigated(&self, path: TabPath) -> Result<()> {
        //::{679F85CB-0220-4080-B29B-5540CC05AAB6} Quick Access
        if let Some(index) = self.0.get_selected_tab_index() {
            self.0.get_tab(index)?.path = path;
            self.0.set_tab_title(get_tab_name(&path), index as usize)?;
        }
        Ok(())
    }

    pub fn new_window(&self, path: TabPath) -> Result<()> {
        self.add_tab(path, self.0.get_tab_count() as usize)
    }
}

impl TabBarRef {
    fn duplicate_tab(&self) -> Result<()> {
        let current_index = self.get_selected_tab_index().ok_or(E_FAIL)?;
        let current_tab = self.get_tab(current_index)?.clone();

        self.add_tab(current_tab.path, current_index + 1)
    }

    fn add_tab(&self, path: TabPath, idx: usize) -> Result<()> {
        let key: usize;
        {
            let obj = &mut *self.0.borrow_mut();
            let tabs = &mut obj.tabs;
            let key_counter = &mut obj.tab_key_counter;

            key = *key_counter;
            *key_counter += 1;
            tabs.insert(key, Tab { path });
        }

        let title = get_tab_name(&path);
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
        match result.0 {
            0 => Err(E_FAIL.into()),
            _ => Ok(()),
        }
    }

    fn get_tab_count(&self) -> i32 {
        let handle = self.0.borrow().handle;
        unsafe { SendMessageW(handle, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0)).0 }
    }

    fn get_selected_tab_index(&self) -> Option<usize> {
        let handle = self.0.borrow().handle;
        unsafe { Some(SendMessageW(handle, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as usize) }
    }

    fn get_tab_key(&self, index: usize) -> Result<usize> {
        let mut tab_info = TCITEMW {
            mask: TCIF_PARAM,
            ..Default::default()
        };
        let handle = self.0.borrow().handle;
        unsafe {
            match SendMessageW(
                handle,
                TCM_GETITEMW,
                WPARAM(index),
                LPARAM(std::ptr::addr_of_mut!(tab_info) as isize),
            ) {
                LRESULT(0) => Err(E_FAIL.into()),
                _ => Ok(tab_info.lParam.0 as usize),
            }
        }
    }

    fn get_tab(&self, index: usize) -> Result<RefMut<Tab>> {
        let key = self.get_tab_key(index)?;
        if !self.0.borrow().tabs.contains_key(&key) {
            return Err(E_FAIL.into());
        }
        let tab_ref = RefMut::map(self.0.borrow_mut(), |obj| obj.tabs.get_mut(&key).unwrap());
        Ok(tab_ref)
    }

    fn remove_tab(&self, index: usize) -> Result<()> {
        let handle = self.0.borrow().handle;
        {
            let key = self.get_tab_key(index)?;
            let tabs = &mut self.0.borrow_mut().tabs;
            tabs.remove(&key);
        }

        unsafe {
            match SendMessageW(handle, TCM_DELETEITEM, WPARAM(index), LPARAM(0)) {
                LRESULT(0) => Err(E_FAIL.into()),
                _ => Ok(()),
            }
        }
    }

    fn tab_switched(&self) -> Result<()> {
        let index = self.get_selected_tab_index().ok_or(E_FAIL)?;
        log::info!("trying to switch to tab {:?}", index);
        let browser = self.0.borrow().explorer.clone();
        let tab = self.get_tab(index)?.clone();
        unsafe { browser.BrowseObject(tab.path, 0) }

        /*
        let url_bstr = unsafe { SysAllocString(tab.path.clone()) };
        let variant: VARIANT = Default::default();
        let variant_ptr = std::ptr::addr_of!(variant);

        let ret = unsafe {
            browser.Navigate(url_bstr, variant_ptr, variant_ptr, variant_ptr, variant_ptr)
        };
        log::info!("navigate to {:?} result {:?}", &tab.path, ret);
        ret
        */
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
            WM_MBUTTONDOWN => self.duplicate_tab().is_ok() as i32,
            WM_LBUTTONUP => self.tab_switched().is_ok() as i32,
            WM_CLOSE => true as i32,
            _ => {
                let proc = self.0.borrow().default_window_procedure;
                unsafe { proc.unwrap()(hwnd, message, wparam, lparam).0 }
            }
        })
    }
}
