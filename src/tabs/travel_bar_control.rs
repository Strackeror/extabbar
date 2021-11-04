use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::UI::{
    Controls::TB_ENABLEBUTTON,
    Shell::{DefSubclassProc, SetWindowSubclass},
    WindowsAndMessaging::SendMessageW,
};

pub struct TravelBarControl {
    handle: HWND,
}

impl TravelBarControl {
    pub extern "system" fn subclass_proc(
        hwnd: HWND,
        message: u32,
        wparam: WPARAM,
        lparam: LPARAM,
        _uid_subclass: usize,
        _ref_data: usize,
    ) -> LRESULT {
        if message == TB_ENABLEBUTTON {
            if lparam.0 & 0x10000 == 0 {
                return LRESULT(1);
            } else {
                log::info!("tab button ok {}", lparam.0 & !0x10000);
                return unsafe {
                    DefSubclassProc(hwnd, message, wparam, LPARAM(lparam.0 & !0x10000))
                };
            }
        }
        unsafe { DefSubclassProc(hwnd, message, wparam, lparam) }
    }
    pub fn new(handle: HWND) -> Self {
        unsafe { SetWindowSubclass(handle, Some(Self::subclass_proc), 0, 0) };
        TravelBarControl { handle }
    }

    pub fn set_button_active(&self, index: usize, active: bool) {
        unsafe {
            SendMessageW(
                self.handle,
                TB_ENABLEBUTTON,
                WPARAM(index as _),
                LPARAM(0x10000 + if active { 1 } else { 0 }),
            );
        }
    }
}
