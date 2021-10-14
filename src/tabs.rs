use std::cell::{RefCell, RefMut};
use std::collections::HashMap;
use std::ptr::{addr_of, addr_of_mut};
use std::rc::Rc;

use bindings::Windows::Win32::Foundation::*;
use bindings::Windows::Win32::Graphics::Gdi::*;
use bindings::Windows::Win32::UI::Controls::*;
use bindings::Windows::Win32::UI::Shell::*;
use bindings::Windows::Win32::UI::WindowsAndMessaging::*;
use windows::Result;

pub static mut DLL_INSTANCE: Option<HINSTANCE> = None;

pub type TabPath = Option<*mut ITEMIDLIST>;

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

    explorer: IShellBrowser,
}

pub extern "system" fn tab_bar_proc(
    hwnd: HWND,
    message: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    _uid_subclass: usize,
    ref_data: usize,
) -> LRESULT {
    let obj_ptr = ref_data as *const TabBarRef;
    let obj = unsafe { &(*obj_ptr) };
    obj.window_procedure(hwnd, message, wparam, lparam)
}

const TAB_BAR_SUBCLASS_UID: usize = 42;

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
    let pidl = match pidl {
        None => return "???".to_owned(),
        Some(pidl) => pidl,
    };

    unsafe {
        let name = SHGetNameFromIDList(*pidl, SIGDN_NORMALDISPLAY);
        let name = match name {
            Ok(name) => pwstr_to_string(name),
            Err(_) => return String::new(),
        };
        name.unwrap_or_else(|_| "???".to_owned())
    }
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
            explorer: browser,
        }))));

        unsafe {
            SetWindowSubclass(
                handle,
                Some(tab_bar_proc),
                TAB_BAR_SUBCLASS_UID,
                Rc::as_ptr(&ret.0) as usize,
            )
            .as_bool()
            .then(|| ())
            .expect("failed to install subclass");
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
                LPARAM(addr_of!(tab_info) as isize),
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
                LPARAM(addr_of!(tab_info) as isize),
            )
        };
        match result.0 {
            0 => Err(E_FAIL.into()),
            _ => Ok(()),
        }
    }

    fn get_tab_count(&self) -> usize {
        let handle = self.0.borrow().handle;
        unsafe { SendMessageW(handle, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0)).0 as usize }
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
                LPARAM(addr_of_mut!(tab_info) as isize),
            ) {
                LRESULT(0) => Err(E_FAIL.into()),
                _ => Ok(tab_info.lParam.0 as usize),
            }
        }
    }

    fn get_tab_text(&self, index: usize) -> Result<String> {
        let mut text = [0u16; 256];
        let mut tab_info = TCITEMW {
            mask: TCIF_TEXT,
            pszText: PWSTR(text.as_mut_ptr()),
            cchTextMax: 256,
            ..Default::default()
        };
        let handle = self.0.borrow().handle;
        unsafe {
            match SendMessageW(
                handle,
                TCM_GETITEMW,
                WPARAM(index),
                LPARAM(addr_of_mut!(tab_info) as isize),
            ) {
                LRESULT(0) => Err(E_FAIL.into()),
                _ => Ok(pwstr_to_string(tab_info.pszText)?),
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

    fn get_tab_rect(&self, index: usize) -> Result<RECT> {
        let handle = self.0.borrow().handle;
        let mut rect: RECT = Default::default();
        unsafe {
            match SendMessageW(
                handle,
                TCM_GETITEMRECT,
                WPARAM(index),
                LPARAM(addr_of_mut!(rect) as isize),
            ) {
                LRESULT(0) => Err(E_FAIL.into()),
                _ => Ok(rect),
            }
        }
    }

    fn tab_switched(&self) -> Result<()> {
        let index = self.get_selected_tab_index().ok_or(E_FAIL)?;
        log::info!("trying to switch to tab {:?}", index);
        let browser = self.0.borrow().explorer.clone();
        let tab = self.get_tab(index)?.clone();
        unsafe { browser.BrowseObject(tab.path.ok_or(E_FAIL)?, 0) }
    }

    const BG_FOCUSED_TAB: u32 = 0x4d4d4d;
    const BG_SELECTED_TAB: u32 = 0x777777;
    fn paint(&self, handle: HWND) -> Result<()> {
        unsafe {
            let mut paint_struct: PAINTSTRUCT = Default::default();
            let mut hdc = BeginPaint(handle, addr_of_mut!(paint_struct));

            {
                let brush = CreateSolidBrush(0x191919);
                FillRect(hdc, addr_of!(paint_struct.rcPaint), brush);
                DeleteObject(brush);
            }

            let edge_pen = CreatePen(PS_SOLID, 1, 0x2b2b2b);
            let hold_pen = SelectObject(hdc, edge_pen);

            let selected_index = self.get_selected_tab_index();

            let font = CreateFontW(
                16,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS,
                CLIP_DEFAULT_PRECIS,
                DEFAULT_QUALITY,
                FF_DONTCARE,
                "Segoe UI",
            );
            let hold_font = SelectObject(hdc, font);

            for idx in 0..self.get_tab_count() {
                let mut tab_rect = self.get_tab_rect(idx)?;
                let mut intersect_rect: RECT = Default::default();

                if !IntersectRect(
                    addr_of_mut!(intersect_rect),
                    addr_of!(paint_struct.rcPaint),
                    addr_of!(tab_rect),
                )
                .as_bool()
                {
                    continue;
                };

                let selected = selected_index == Some(idx);
                if !selected {
                    tab_rect.top += 2;
                }
                {
                    let brush = CreateSolidBrush(if selected {
                        Self::BG_SELECTED_TAB
                    } else {
                        0x202020
                    });
                    FillRect(hdc, addr_of!(tab_rect), brush);
                    DeleteObject(brush);
                }

                let edges = [
                    POINT {
                        x: tab_rect.left,
                        y: tab_rect.top,
                    },
                    POINT {
                        x: tab_rect.right - 1,
                        y: tab_rect.top,
                    },
                    POINT {
                        x: tab_rect.right - 1,
                        y: tab_rect.bottom,
                    },
                ];

                Polyline(hdc, edges.as_ptr(), edges.len() as i32);

                let mut text_rect = tab_rect;
                text_rect.top += 2;
                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(hdc, 0xffffff);
                DrawTextW(
                    hdc,
                    self.get_tab_text(idx).unwrap_or_default(),
                    -1,
                    addr_of_mut!(text_rect),
                    DT_CENTER,
                );
            }
            SelectObject(hdc, hold_pen);
            DeleteObject(edge_pen);
            SelectObject(hdc, hold_font);
            DeleteObject(font);
        }

        Ok(())
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
            WM_PAINT => {
                self.paint(hwnd).unwrap_or_default();
                //unsafe { DefSubclassProc(hwnd, message, wparam, lparam).0 }
                1
            }
            _ => unsafe { DefSubclassProc(hwnd, message, wparam, lparam).0 },
        })
    }
}
