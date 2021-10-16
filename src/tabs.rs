use std::cell::{RefCell, RefMut};
use std::collections::HashMap;
use std::ptr::{addr_of, addr_of_mut};
use std::rc::{Rc, Weak};

use bindings::Windows::Win32::Foundation::*;
use bindings::Windows::Win32::Graphics::Gdi::*;
use bindings::Windows::Win32::UI::Controls::*;
use bindings::Windows::Win32::UI::Shell::*;
use bindings::Windows::Win32::UI::WindowsAndMessaging::*;
use windows::Result;

pub static mut DLL_INSTANCE: Option<HINSTANCE> = None;

// A possible path for a tab
pub type TabPath = Option<*mut ITEMIDLIST>;

type TabKey = usize;
type TabIndex = usize;

#[derive(Clone)]
pub struct Tab {
    current_path: TabPath,

    forward_paths: Vec<TabPath>,
    backward_paths: Vec<TabPath>,
}

struct TabBar_ {
    tabs: HashMap<TabKey, Tab>,
    tab_key_counter: TabKey,

    tab_control: Option<Box<TabControl>>,

    explorer: IShellBrowser,
}
pub struct TabBar(RefCell<TabBar_>);

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
    pub fn new(parent: HWND, browser: IShellBrowser) -> Rc<TabBar> {
        let new = TabBar_ {
            tabs: Default::default(),
            tab_key_counter: 0,
            tab_control: None,
            explorer: browser,
        };
        let new = Rc::new(TabBar(RefCell::new(new)));
        let tab_control = TabControl::new(parent, Rc::downgrade(&new));
        new.0.borrow_mut().tab_control = Some(tab_control);
        new
    }

    fn tab_control(&self) -> Box<TabControl> {
        return self.0.borrow().tab_control.as_ref().unwrap().clone();
    }

    fn get_tab(&self, idx: TabIndex) -> Option<RefMut<Tab>> {
        let key = self.tab_control().get_tab_key(idx).ok()?;
        if self.0.borrow().tabs.contains_key(&key) {
            Some(RefMut::map(self.0.borrow_mut(), |tab_bar| {
                tab_bar.tabs.get_mut(&key).unwrap()
            }))
        } else {
            None
        }
    }

    pub fn get_handle(&self) -> HWND {
        self.tab_control().handle
    }

    fn add_tab_entry(&self, path: TabPath) -> TabKey {
        let obj = &mut *self.0.borrow_mut();
        let tabs = &mut obj.tabs;
        let key_counter = &mut obj.tab_key_counter;
        let key = *key_counter;
        *key_counter += 1;
        tabs.insert(
            key,
            Tab {
                current_path: path,
                forward_paths: Vec::new(),
                backward_paths: Vec::new(),
            },
        );
        key
    }

    pub fn add_tab(&self, path: TabPath, idx: usize) -> Result<()> {
        let key = self.add_tab_entry(path);
        self.tab_control().add_tab(path, idx, key)
    }

    pub fn remove_tab(&self, idx: TabIndex) -> Result<()> {
        let key = self.tab_control().get_tab_key(idx)?;
        {
            let tabs = &mut self.0.borrow_mut().tabs;
            if tabs.contains_key(&key) {
                tabs.remove(&key);
            }
        }
        self.tab_control().remove_tab(idx)?;
        Ok(())
    }

    pub fn navigated(&self, path: TabPath) -> Result<()> {
        if let Some(index) = self.tab_control().get_selected_tab_index() {
            {
                let mut tab = self.get_tab(index).ok_or(E_FAIL)?;
                tab.forward_paths.clear();
                tab.backward_paths.push(path);
                tab.current_path = path;
            }
            self.tab_control()
                .set_tab_title(index, get_tab_name(&path))?;
        }
        Ok(())
    }

    fn tab_switched(&self) -> Result<()> {
        let index = self.tab_control().get_selected_tab_index().ok_or(E_FAIL)?;
        log::info!("trying to switch to tab {:?}", index);
        let browser = self.0.borrow().explorer.clone();
        let tab = self.get_tab(index).ok_or(E_FAIL)?.clone();
        unsafe { browser.BrowseObject(tab.current_path.ok_or(E_FAIL)?, 0) }
    }

    pub fn clone_tab(&self, idx: TabIndex) -> Result<()> {
        let tab = self.get_tab(idx).ok_or(E_FAIL)?.clone();
        self.add_tab(tab.current_path, idx + 1)?;
        Err(E_FAIL.into())
    }

    pub fn new_window(&self, path: TabPath) -> Result<()> {
        self.add_tab(path, self.tab_control().get_tab_count() as usize)
    }
}

#[derive(Clone)]
struct TabControl {
    handle: HWND,
    tab_bar: Weak<TabBar>,
    _pin: std::marker::PhantomPinned,
}

impl TabControl {
    pub extern "system" fn tab_bar_proc(
        hwnd: HWND,
        message: u32,
        wparam: WPARAM,
        lparam: LPARAM,
        _uid_subclass: usize,
        ref_data: usize,
    ) -> LRESULT {
        let obj = ref_data as *const TabControl;
        let obj = unsafe { obj.as_ref().unwrap() };
        obj.window_procedure(hwnd, message, wparam, lparam)
    }

    fn new(parent_handle: HWND, tab_bar: Weak<TabBar>) -> Box<TabControl> {
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
                parent_handle,
                HMENU(0),
                DLL_INSTANCE.unwrap(),
                std::ptr::null(),
            )
        };

        let new = Box::new(TabControl {
            handle,
            tab_bar,
            _pin: Default::default(),
        });

        unsafe {
            SetWindowSubclass(
                handle,
                Some(TabControl::tab_bar_proc),
                TAB_BAR_SUBCLASS_UID,
                &*new as *const TabControl as usize,
            )
            .as_bool()
            .then(|| ())
            .expect("failed to install subclass");
        }
        new
    }

    fn add_tab(&self, path: TabPath, idx: TabIndex, key: TabKey) -> Result<()> {
        let title = get_tab_name(&path);
        let handle = self.handle;
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

    fn set_tab_title(&self, idx: TabIndex, title: String) -> Result<()> {
        let handle = self.handle;

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
        let handle = self.handle;
        unsafe { SendMessageW(handle, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0)).0 as usize }
    }

    fn get_selected_tab_index(&self) -> Option<TabIndex> {
        let handle = self.handle;
        unsafe { Some(SendMessageW(handle, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as usize) }
    }

    fn get_tab_key(&self, index: TabIndex) -> Result<TabKey> {
        let mut tab_info = TCITEMW {
            mask: TCIF_PARAM,
            ..Default::default()
        };
        let handle = self.handle;
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

    fn get_tab_text(&self, index: TabIndex) -> Result<String> {
        let mut text = [0u16; 256];
        let mut tab_info = TCITEMW {
            mask: TCIF_TEXT,
            pszText: PWSTR(text.as_mut_ptr()),
            cchTextMax: 256,
            ..Default::default()
        };
        let handle = self.handle;
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

    fn remove_tab(&self, index: TabIndex) -> Result<()> {
        let handle = self.handle;
        unsafe {
            match SendMessageW(handle, TCM_DELETEITEM, WPARAM(index), LPARAM(0)) {
                LRESULT(0) => Err(E_FAIL.into()),
                _ => Ok(()),
            }
        }
    }

    fn get_tab_rect(&self, index: TabIndex) -> Result<RECT> {
        let handle = self.handle;
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

    // file under mouse : 0x4d4d4d
    // file selected : 0x777777
    // background : 0x191919
    const BG_FOCUSED_TAB: u32 = 0x4d4d4d;
    const BG_SELECTED_TAB: u32 = 0x191919;
    const BG_UNFOCUSED_TAB: u32 = 0x202020;
    const BORDER_COLOR: u32 = 0x2b2b2b;
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
                        Self::BG_UNFOCUSED_TAB
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
        if let Some(tab_bar) = self.tab_bar.upgrade() {
            let result = match message {
                WM_PAINT => {
                    let result = self.paint(hwnd);
                    return LRESULT(result.is_ok() as i32);
                }
                WM_MBUTTONDOWN => tab_bar.clone_tab(self.get_selected_tab_index().unwrap_or(0)),
                WM_RBUTTONUP => tab_bar.remove_tab(self.get_selected_tab_index().unwrap_or(0)),
                WM_LBUTTONUP => tab_bar.tab_switched(),
                _ => Ok(()),
            };

            if result.is_err() {
                return LRESULT(0);
            }
        }
        unsafe { DefSubclassProc(hwnd, message, wparam, lparam) }
    }
}
