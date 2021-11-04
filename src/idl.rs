use windows::Win32::UI::Shell::{ILClone, ILFree, ILIsEqual, ITEMIDLIST};

pub struct Idl(*const ITEMIDLIST);

impl Clone for Idl {
    fn clone(&self) -> Self {
        Idl::new(self.0)
    }
}

impl Drop for Idl {
    fn drop(&mut self) {
        unsafe { ILFree(self.0) }
    }
}

impl Idl {
    pub fn new(ptr: *const ITEMIDLIST) -> Self {
        Self(unsafe { ILClone(ptr) })
    }

    pub fn get(&self) -> *const ITEMIDLIST {
        self.0
    }
}

impl PartialEq for Idl {
    fn eq(&self, other: &Self) -> bool {
        unsafe { ILIsEqual(self.0, other.0).as_bool() }
    }
}
