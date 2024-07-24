use dunce::canonicalize;
use std::ffi::OsString;
use std::fs::OpenOptions;
use std::path::{Path, PathBuf};
use indicatif::ProgressBar;
use walkdir::DirEntry;

pub fn check_file(path: &Path) -> Option<PathBuf> {
    // open options (just checking so write is disabled)
    let file = OpenOptions::new()
        .read(true)
        .write(false)
        .open(path)
        .expect("failed to open file");

    // grab the zip folder contents or return None if there is an error
    let zip_archive = match zip::ZipArchive::new(file) {
        Ok(val) => val,
        Err(_) => {
            return None;
        }
    };

    // find an externalLinks file and exit
    for i in 0..zip_archive.len() {
        if let Some(file_name) = zip_archive.name_for_index(i) {
            if file_name.starts_with("xl/externalLinks") {
                if let Ok(path) = canonicalize(path) {
                    return Some(path);
                }
            }
        }
    }
    None
}

#[cfg(target_os = "windows")]
pub fn determine_files_with_links(dirs: Vec<DirEntry>, pb: &mut ProgressBar) -> Vec<PathBuf> {
    let mut remove_files = vec![];
    let xlsx_ext = OsString::from("xlsx");
    for file in dirs {
        pb.inc(1);
        let file_path = file.path();
        if Some(xlsx_ext.as_os_str()) == file_path.extension() {
            if let Some(file_path) = check_file(file_path) {
                remove_files
                    .push(canonicalize(file_path).expect("failed to canonicalize filepath"));
            }
        }
    }
    remove_files
}
