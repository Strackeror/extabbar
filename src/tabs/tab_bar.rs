use std::cell::{RefCell, RefMut};
use std::collections::HashMap;
use std::rc::Rc;

use bindings::Windows::Win32::Foundation::*;
use bindings::Windows::Win32::UI::Shell::*;
use windows::Result;

use super::explorer_subclass::ExplorerSubclass;
use super::tab_control::{pwstr_to_string, TabControl};

pub static mut DLL_INSTANCE: Option<HINSTANCE> = None;

// A possible path for a tab
pub type TabPath = Option<*mut ITEMIDLIST>;

pub type TabKey = usize;
pub type TabIndex = usize;

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
    explorer_subclass: Option<Box<ExplorerSubclass>>,

    explorer: IShellBrowser,
}
pub struct TabBar(RefCell<TabBar_>);
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
            explorer_subclass: None,
            explorer: browser,
        };
        let new = Rc::new(TabBar(RefCell::new(new)));
        let tab_control = TabControl::new(parent, Rc::downgrade(&new));
        new.0.borrow_mut().tab_control = Some(tab_control);
        new.0.borrow_mut().explorer_subclass =
            Some(ExplorerSubclass::new(parent, Rc::downgrade(&new)));
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
        self.tab_control().add_tab(get_tab_name(&path), idx, key)
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

    pub fn switch_to_current_tab(&self) -> Result<()> {
        let index = self.tab_control().get_selected_tab_index().ok_or(E_FAIL)?;
        self.switch_tab(index)
    }

    pub fn switch_tab(&self, index: TabIndex) -> Result<()> {
        log::info!("trying to switch to tab {:?}", index);
        self.tab_control().set_selected_tab(index)?;
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
        let idx = self.tab_control().get_tab_count();
        self.add_tab(path, self.tab_control().get_tab_count() as usize)?;
        self.switch_tab(idx)
    }
}
