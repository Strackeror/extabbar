use std::cell::{RefCell, RefMut};
use std::collections::HashMap;
use std::rc::Rc;

use windows::core::{Interface, Result};
use windows::Win32::Foundation::*;
use windows::Win32::UI::Shell::*;
use windows::Win32::UI::WindowsAndMessaging::SetForegroundWindow;

use crate::idl::Idl;
use crate::settings::Settings;

use super::explorer_subclass::ExplorerSubclass;
use super::tab_control::{pwstr_to_string, TabControl};
use super::travel_bar_control::TravelBarControl;

pub static mut DLL_INSTANCE: Option<HINSTANCE> = None;

// A possible path for a tab
pub type TabPath = Option<Idl>;

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

    tab_control: Box<TabControl>,
    _explorer_subclass: Box<ExplorerSubclass>,
    travel_toolbar: TravelBarControl,

    explorer: IShellBrowser,
    explorer_handle: HWND,
    is_main: bool,
}
pub struct TabBar(RefCell<TabBar_>);
fn get_tab_name(pidl: &TabPath) -> String {
    let pidl = match pidl {
        None => return "???".to_owned(),
        Some(pidl) => pidl,
    };

    unsafe {
        let name = SHGetNameFromIDList(pidl.get(), SIGDN_NORMALDISPLAY);
        let name = match name {
            Ok(name) => pwstr_to_string(name),
            Err(_) => return String::new(),
        };
        name.unwrap_or_else(|_| "???".to_owned())
    }
}

pub fn get_current_folder_path(browser: &IShellBrowser) -> TabPath {
    unsafe {
        let folder_view: IFolderView = browser.QueryActiveShellView().ok()?.cast().ok()?;
        let folder = folder_view.GetFolder::<IPersistFolder2>().ok()?;
        let folder_pidl = folder.GetCurFolder();
        if folder_pidl.is_err() {
            log::error!("Could not get pidl for current path");
        }
        Some(Idl::new(folder_pidl.ok()?))
    }
}

impl TabBar {
    pub fn new(
        parent: HWND,
        explorer_handle: HWND,
        travel_toolbar_handle: HWND,
        browser: IShellBrowser,
        settings: Settings,
        is_main: bool,
    ) -> Rc<TabBar> {
        Rc::new_cyclic(|weak| {
            TabBar(RefCell::new(TabBar_ {
                tabs: Default::default(),
                tab_key_counter: 0,
                tab_control: TabControl::new(parent, weak.clone(), settings.dark_mode),
                travel_toolbar: TravelBarControl::new(travel_toolbar_handle),
                _explorer_subclass: ExplorerSubclass::new(explorer_handle, weak.clone()),
                explorer: browser,
                explorer_handle,
                is_main,
            }))
        })
    }

    pub fn is_main(&self) -> bool {
        self.0.borrow().is_main
    }

    pub fn get_handle(&self) -> HWND {
        self.tab_control().handle
    }
    fn tab_control(&self) -> Box<TabControl> {
        return self.0.borrow().tab_control.clone();
    }

    fn get_tab(&self, index: TabIndex) -> Option<RefMut<Tab>> {
        let key = self.tab_control().get_tab_key(index).ok()?;
        if self.0.borrow().tabs.contains_key(&key) {
            Some(RefMut::map(self.0.borrow_mut(), |tab_bar| {
                tab_bar.tabs.get_mut(&key).unwrap()
            }))
        } else {
            None
        }
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

    pub fn add_tab(&self, path: TabPath, index: usize) -> Result<()> {
        let key = self.add_tab_entry(path.clone());
        self.tab_control().add_tab(get_tab_name(&path), index, key)
    }

    pub fn remove_tab(&self, index: TabIndex) -> Result<()> {
        let key = self.tab_control().get_tab_key(index)?;
        if Some(index) == self.tab_control().get_selected_tab_index() {
            if self.tab_control().get_tab_count() <= 1 {
                // Don't allow removing last tab
                return Ok(());
            } else if index == 0 {
                self.switch_tab(1)?;
            } else {
                self.switch_tab(index - 1)?;
            }
        }
        {
            let tabs = &mut self.0.borrow_mut().tabs;
            if tabs.contains_key(&key) {
                tabs.remove(&key);
            }
        }
        self.tab_control().remove_tab(index)?;
        Ok(())
    }

    pub fn navigated(&self, path: TabPath) -> Result<()> {
        let index = self.tab_control().get_selected_tab_index().ok_or(E_FAIL)?;
        log::info!("tab {:?}, navigated to {:?}", index, get_tab_name(&path));
        {
            let mut tab = self.get_tab(index).ok_or(E_FAIL)?;
            let current_path = tab.current_path.clone();
            if tab.current_path != path && tab.current_path.is_some() {
                tab.forward_paths.clear();
                tab.backward_paths.push(current_path);
            }
            tab.current_path = path.clone();
        }

        self.tab_control()
            .set_tab_title(index, get_tab_name(&path))?;

        let can_go_backward = !self.get_tab(index).ok_or(E_FAIL)?.backward_paths.is_empty();
        let can_go_forward = !self.get_tab(index).ok_or(E_FAIL)?.forward_paths.is_empty();
        self.0
            .borrow()
            .travel_toolbar
            .set_button_active(256, can_go_backward);
        self.0
            .borrow()
            .travel_toolbar
            .set_button_active(257, can_go_forward);

        Ok(())
    }

    pub fn navigate_back(&self) -> Result<()> {
        let index = self.tab_control().get_selected_tab_index().ok_or(E_FAIL)?;

        let mut tab = self.get_tab(index).ok_or(E_FAIL)?;
        let current_path = tab.current_path.clone();
        let next_path = tab.backward_paths.pop().ok_or(E_FAIL)?;
        tab.forward_paths.push(current_path);
        tab.current_path = next_path.clone();
        std::mem::drop(tab);

        self.browse_to(next_path)
    }

    pub fn navigate_forward(&self) -> Result<()> {
        let index = self.tab_control().get_selected_tab_index().ok_or(E_FAIL)?;

        let mut tab = self.get_tab(index).ok_or(E_FAIL)?;
        let current_path = tab.current_path.clone();
        let next_path = tab.forward_paths.pop().ok_or(E_FAIL)?;
        tab.backward_paths.push(current_path);
        tab.current_path = next_path.clone();
        std::mem::drop(tab);

        self.browse_to(next_path)
    }

    fn browse_to(&self, path: TabPath) -> Result<()> {
        let browser = self.0.borrow().explorer.clone();
        unsafe { browser.BrowseObject(path.ok_or(E_FAIL)?.get(), SBSP_SAMEBROWSER)? }
        Ok(())
    }

    pub fn _switch_to_current_tab(&self) -> Result<()> {
        let index = self.tab_control().get_selected_tab_index().ok_or(E_FAIL)?;
        self.switch_tab(index)
    }

    pub fn toggle_dark_mode(&self) {
        log::info!("toggle dark mode");
        let mut mut_self = self.0.borrow_mut();
        let dark_mode_ref: &mut bool = &mut mut_self.tab_control.dark_mode;
        *dark_mode_ref = !*dark_mode_ref;
    }

    pub fn switch_tab(&self, index: TabIndex) -> Result<()> {
        log::info!("trying to switch to tab {:?}", index);
        self.tab_control().set_selected_tab(index)?;
        let path = self.get_tab(index).ok_or(E_FAIL)?.current_path.clone();
        self.browse_to(path)
    }

    pub fn clone_tab(&self, index: TabIndex) -> Result<()> {
        let tab = self.get_tab(index).ok_or(E_FAIL)?.clone();
        self.add_tab(tab.current_path, index + 1)
    }

    pub fn new_window(&self, path: TabPath) -> Result<()> {
        let index = self.tab_control().get_tab_count();
        self.add_tab(path, index)?;
        self.switch_tab(index)?;
        unsafe {
            SetForegroundWindow(self.0.borrow().explorer_handle);
        }
        Ok(())
    }
}
