use bindings::Windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, PSTR, WPARAM};
use bindings::Windows::Win32::UI::Controls::{
    TCIF_STATE, TCIF_TEXT, TCIS_HIGHLIGHTED, TCITEMA, TCM_INSERTITEMA, TCM_SETITEMA,
    TCS_FOCUSNEVER, TCS_SINGLELINE,
};
use bindings::Windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExA, SendMessageA, HMENU, WINDOW_EX_STYLE, WINDOW_STYLE, WS_CHILD, WS_CLIPSIBLINGS,
    WS_VISIBLE,
};

pub static mut DLL_INSTANCE: Option<HINSTANCE> = None;
pub struct WindowHandler {
    pub handle: HWND,
}

impl WindowHandler {
    pub fn init(parent: HWND) -> WindowHandler {
        let handle = unsafe {
            CreateWindowExA(
                WINDOW_EX_STYLE(0),
                "SysTabControl32",
                "",
                WS_CHILD
                    | WS_CLIPSIBLINGS
                    | WS_VISIBLE
                    | WINDOW_STYLE(TCS_FOCUSNEVER)
                    | WINDOW_STYLE(TCS_SINGLELINE),
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
        WindowHandler { handle }
    }

    pub fn add_tab(&mut self, title: &mut str, idx: usize) {
        let tab_info = TCITEMA {
            mask: TCIF_STATE | TCIF_TEXT,
            dwState: 0,
            dwStateMask: TCIS_HIGHLIGHTED,
            pszText: PSTR(title.as_mut_ptr()),
            cchTextMax: 0,
            iImage: 0,
            lParam: LPARAM(0),
        };
        unsafe {
            let result = SendMessageA(
                &self.handle,
                TCM_INSERTITEMA,
                WPARAM(idx),
                LPARAM(std::ptr::addr_of!(tab_info) as isize),
            );

            log::info!("Added tab done, result:{:?}", result);
            let result = SendMessageA(
                &self.handle,
                TCM_SETITEMA,
                WPARAM(idx),
                LPARAM(std::ptr::addr_of!(tab_info) as isize),
            );
            log::info!("Added tab done, result:{:?}", result);
        }
    }
}
