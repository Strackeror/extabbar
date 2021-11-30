use std::ptr::{addr_of, addr_of_mut};
use std::rc::Weak;

use windows::Win32::{
    Foundation::*, Graphics::Gdi::*, UI::Controls::*, UI::Shell::*, UI::WindowsAndMessaging::*,
};

use windows::runtime::*;

use super::tab_bar::{TabBar, TabIndex, TabKey, DLL_INSTANCE};

const TAB_BAR_SUBCLASS_UID: usize = 42;

#[derive(Clone)]
pub struct TabControl {
    pub handle: HWND,
    pub dark_mode: bool,
    tab_bar: Weak<TabBar>,
    focused_tab: Option<TabIndex>,
    _pin: std::marker::PhantomPinned,
}

pub unsafe fn pwstr_to_string(pwstr: PWSTR) -> Result<String> {
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

impl TabControl {
    pub extern "system" fn tab_bar_proc(
        hwnd: HWND,
        message: u32,
        wparam: WPARAM,
        lparam: LPARAM,
        _uid_subclass: usize,
        ref_data: usize,
    ) -> LRESULT {
        let obj = ref_data as *mut TabControl;
        let obj = unsafe { &mut *obj };
        obj.window_procedure(hwnd, message, wparam, lparam)
    }

    pub fn new(parent_handle: HWND, tab_bar: Weak<TabBar>, dark_mode: bool) -> Box<TabControl> {
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
            dark_mode,
            handle,
            tab_bar,
            focused_tab: None,
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

    pub fn add_tab(&self, title: String, index: TabIndex, key: TabKey) -> Result<()> {
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
                WPARAM(index),
                LPARAM(addr_of!(tab_info) as isize),
            )
        };

        log::info!("Added tab done, result:{:?}", result);
        if result.0 < 0 {
            return Err(E_FAIL.into());
        }
        Ok(())
    }

    pub fn set_tab_title(&self, index: TabIndex, title: String) -> Result<()> {
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
                WPARAM(index),
                LPARAM(addr_of!(tab_info) as isize),
            )
        };
        match result.0 {
            0 => Err(E_FAIL.into()),
            _ => Ok(()),
        }
    }

    pub fn set_selected_tab(&self, index: TabIndex) -> Result<()> {
        let handle = self.handle;
        match unsafe { SendMessageW(handle, TCM_SETCURSEL, WPARAM(index), LPARAM(0)).0 } {
            -1 => E_FAIL.ok(),
            _ => Ok(()),
        }
    }

    pub fn get_tab_count(&self) -> usize {
        let handle = self.handle;
        unsafe { SendMessageW(handle, TCM_GETITEMCOUNT, WPARAM(0), LPARAM(0)).0 as usize }
    }

    pub fn get_selected_tab_index(&self) -> Option<TabIndex> {
        let handle = self.handle;
        unsafe { Some(SendMessageW(handle, TCM_GETCURSEL, WPARAM(0), LPARAM(0)).0 as usize) }
    }

    pub fn get_focused_tab_index(&self) -> Option<TabIndex> {
        let handle = self.handle;
        unsafe { Some(SendMessageW(handle, TCM_GETCURFOCUS, WPARAM(0), LPARAM(0)).0 as usize) }
    }

    pub fn get_hovered_tab_index(&self) -> Option<TabIndex> {
        let mut point = POINT::default();

        unsafe {
            if !GetCursorPos(&mut point as _).as_bool() {
                return None;
            }
        }
        log::info!("cursor pos: {:?}", point);
        self.get_tab_at_coords(point.x, point.y)
    }

    pub fn get_tab_at_coords(&self, x: i32, y: i32) -> Option<TabIndex> {
        let handle = self.handle;
        let mut hit_test_info = TCHITTESTINFO {
            pt: POINT { x, y },
            ..Default::default()
        };

        let ret = unsafe {
            SendMessageW(
                handle,
                TCM_HITTEST,
                WPARAM(0),
                LPARAM(addr_of_mut!(hit_test_info) as _),
            )
        };
        if ret == LRESULT(-1) {
            return None;
        }
        Some(ret.0 as _)
    }

    pub fn get_tab_key(&self, index: TabIndex) -> Result<TabKey> {
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

    pub fn get_tab_text(&self, index: TabIndex) -> Result<String> {
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

    pub fn remove_tab(&self, index: TabIndex) -> Result<()> {
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

    fn create_popup_menu(&self) -> Result<()> {
        let menu = unsafe { CreatePopupMenu() };
        unsafe { AppendMenuW(menu, MF_STRING, 1001, "Toggle Dark Mode") };
        unsafe {
            let mut point = POINT::default();
            GetCursorPos(&mut point as *mut _).ok()?;
            TrackPopupMenu(
                menu,
                TPM_LEFTALIGN | TPM_RIGHTBUTTON,
                point.x,
                point.y,
                0,
                self.handle,
                0 as _,
            );
            DestroyMenu(menu).ok()?;
        }

        Ok(())
    }

    // file under mouse : 0x4d4d4d
    // file selected : 0x777777
    // background : 0x191919
    const BACKGROUND: u32 = 0x191919;
    const BG_FOCUSED_TAB: u32 = 0x4d4d4d;
    const BG_SELECTED_TAB: u32 = 0x191919;
    const BG_UNFOCUSED_TAB: u32 = 0x202020;
    const BORDER_COLOR: u32 = 0x2b2b2b;
    fn paint(&self, handle: HWND) -> Result<()> {
        unsafe {
            let mut paint_struct: PAINTSTRUCT = Default::default();
            let hdc = BeginPaint(handle, addr_of_mut!(paint_struct));
            {
                let brush = CreateSolidBrush(Self::BACKGROUND);
                FillRect(hdc, addr_of!(paint_struct.rcPaint), brush);
                DeleteObject(brush);
            }

            let edge_pen = CreatePen(PS_SOLID, 1, Self::BORDER_COLOR);
            let hold_pen = SelectObject(hdc, edge_pen);

            let selected_index = self.get_selected_tab_index();
            let focused_index = self.focused_tab;
            log::info!("{:?}", (selected_index, focused_index));

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

            for index in 0..self.get_tab_count() {
                let mut tab_rect = self.get_tab_rect(index)?;
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

                let selected = selected_index == Some(index);
                let focused = focused_index == Some(index);
                if !selected {
                    tab_rect.top += 2;
                }
                {
                    let color = match (selected, focused) {
                        (true, false) => Self::BG_SELECTED_TAB,
                        (false, false) => Self::BG_UNFOCUSED_TAB,
                        (_, true) => Self::BG_FOCUSED_TAB,
                    };
                    let brush = CreateSolidBrush(color);
                    FillRect(hdc, addr_of!(tab_rect), brush);
                    DeleteObject(brush);
                }

                let edges = [
                    POINT {
                        x: tab_rect.left,
                        y: tab_rect.bottom,
                    },
                    POINT {
                        x: tab_rect.left,
                        y: tab_rect.top,
                    },
                    POINT {
                        x: tab_rect.right,
                        y: tab_rect.top,
                    },
                    POINT {
                        x: tab_rect.right,
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
                    self.get_tab_text(index).unwrap_or_default(),
                    -1,
                    addr_of_mut!(text_rect),
                    DT_CENTER,
                );
            }
            SelectObject(hdc, hold_pen);
            DeleteObject(edge_pen);
            SelectObject(hdc, hold_font);
            DeleteObject(font);
            EndPaint(handle, addr_of_mut!(paint_struct));
        }

        Ok(())
    }

    fn window_procedure(
        &mut self,
        hwnd: HWND,
        message: u32,
        wparam: WPARAM,
        lparam: LPARAM,
    ) -> LRESULT {
        if let Some(tab_bar) = self.tab_bar.upgrade() {
            let result = match message {
                WM_COMMAND => match wparam.0 {
                    1001 => {
                        tab_bar.toggle_dark_mode();
                        unsafe {
                            InvalidateRect(hwnd, std::ptr::null(), BOOL(1));
                            UpdateWindow(hwnd);
                        }
                        Ok(())
                    }
                    _ => Ok(()),
                },
                WM_PAINT => match self.dark_mode {
                    true => return LRESULT(self.paint(hwnd).is_ok() as _),
                    false => Ok(()),
                },
                WM_MBUTTONDOWN => match self.focused_tab {
                    Some(index) => tab_bar
                        .remove_tab(index)
                        .map_err(|err| log::error!("{:?}", err)),
                    None => Ok(()),
                },
                WM_LBUTTONDOWN => match self.focused_tab {
                    Some(index) => tab_bar
                        .switch_tab(index)
                        .map_err(|err| log::error!("{:?}", err)),
                    None => Ok(()),
                },
                WM_RBUTTONDOWN => return LRESULT(self.create_popup_menu().is_ok() as _),
                WM_MOUSEMOVE => unsafe {
                    let x = (lparam.0 & 0xffff) as i16;
                    let y = ((lparam.0 >> 16) & 0xffff) as i16;
                    let focused_tab = self.get_tab_at_coords(x as _, y as _);
                    if focused_tab != self.focused_tab {
                        self.focused_tab = focused_tab;
                        log::info!("repaint");
                        InvalidateRect(hwnd, std::ptr::null(), BOOL(1));
                        UpdateWindow(hwnd);
                    }
                    Ok(())
                },
                WM_MOUSELEAVE => {
                    self.focused_tab = None;
                    Ok(())
                }
                _ => Ok(()),
            };

            if result.is_err() {
                return LRESULT(0);
            }
        }
        unsafe { DefSubclassProc(hwnd, message, wparam, lparam) }
    }
}
