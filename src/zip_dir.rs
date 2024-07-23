use std::fmt::{Display, Formatter};
use std::fs::{File, OpenOptions};
use std::io::{Read, Write};
use std::os::windows::prelude::*;
use std::path::Path;
use std::str::from_utf8;
use std::time::{Duration, SystemTime, UNIX_EPOCH};
use std::{fs, io};

use filetime::FileTime;
use regex::Regex;
use walkdir::WalkDir;
use windows_sys::Win32::Foundation::{BOOL, FILETIME, HANDLE};
use windows_sys::Win32::Storage::FileSystem::*;
use zip::result::ZipError;
use zip::write::SimpleFileOptions;
use zip::CompressionMethod;

#[derive(Default, Copy, Clone)]
pub struct MetaApply {
    pub created_time: Option<std::time::SystemTime>,
    pub last_accessed: Option<std::time::SystemTime>,
    pub last_modified: Option<std::time::SystemTime>,
}

impl Display for MetaApply {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        let format_time = |time: Option<SystemTime>| -> String {
            match time {
                Some(t) => {
                    let duration = t.duration_since(UNIX_EPOCH).unwrap_or(Duration::new(0, 0));
                    let datetime = chrono::DateTime::from_timestamp(
                        duration.as_secs() as i64,
                        duration.subsec_nanos(),
                    )
                    .expect("failed to make datetime from system time for display");
                    datetime.format("%Y-%m-%d").to_string()
                }
                None => "None".to_string(),
            }
        };

        writeln!(
            f,
            "created_time: {}\nlast_accessed: {}\nlast_modified: {}",
            format_time(self.created_time),
            format_time(self.last_accessed),
            format_time(self.last_modified)
        )
    }
}

impl MetaApply {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn add_created_time(&mut self, t: SystemTime) -> &mut MetaApply {
        self.created_time.replace(t);
        self
    }

    pub fn add_last_accessed(&mut self, t: SystemTime) -> &mut MetaApply {
        self.last_accessed.replace(t);
        self
    }

    pub fn add_last_modified(&mut self, t: SystemTime) -> &mut MetaApply {
        self.last_modified.replace(t);
        self
    }
}

fn to_filetime(ft: impl Into<FileTime>) -> FILETIME {
    let ft = ft.into();
    let intervals = ft.seconds() * (1_000_000_000 / 100) + ((ft.nanoseconds() as i64) / 100);
    FILETIME {
        dwLowDateTime: intervals as u32,
        dwHighDateTime: (intervals >> 32) as u32,
    }
}

pub(crate) fn get_meta<P: AsRef<Path>>(file_path: P) -> MetaApply {
    let mut meta_apply = MetaApply::new();
    if let Ok(metadata) = file_path.as_ref().metadata().as_ref() {
        match metadata.created() {
            Ok(val) => {
                meta_apply.add_created_time(val);
            }
            Err(e) => println!("{e}"),
        };

        match metadata.accessed() {
            Ok(val) => {
                meta_apply.add_last_accessed(val);
            }
            Err(e) => println!("{e}"),
        };

        match metadata.modified() {
            Ok(val) => {
                meta_apply.add_last_modified(val);
            }
            Err(e) => println!("{e}"),
        };
    }
    meta_apply
}

pub(crate) fn set_meta<P>(file_path: P, meta: MetaApply)
where
    P: AsRef<Path> + Copy,
{
    if let (Some(accessed_time), Some(modified_time), Some(created_time)) =
        (meta.last_accessed, meta.last_modified, meta.created_time)
    {
        let atime = &to_filetime(accessed_time) as *const FILETIME;
        let mtime = &to_filetime(modified_time) as *const FILETIME;
        let ctime = &to_filetime(created_time) as *const FILETIME;

        let f = OpenOptions::new()
            .write(true)
            .custom_flags(FILE_FLAG_BACKUP_SEMANTICS)
            .open(file_path)
            .unwrap();
        unsafe {
            let ret: BOOL = SetFileTime(f.as_raw_handle() as HANDLE, ctime, atime, mtime);
            if ret != 0 {
                Ok(())
            } else {
                Err(io::Error::last_os_error())
            }
        }
        .expect("failed to set times");
    }
}

/// extract the excel contents
pub(crate) fn extract_dir<P: AsRef<Path>>(
    file_path: P,
    target: &str,
) -> zip::result::ZipResult<String> {
    // read file from file path
    let file = File::open(&file_path)?;
    let mut archive = zip::ZipArchive::new(file)?;
    // construct a base path for extracted files
    let base_path = Path::new(&target);

    if let Err(why) = fs::create_dir(base_path) {
        println!("! {:?}, writing over it", why.kind())
    }

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let out_path = match file.enclosed_name() {
            Some(path) => path.to_owned(),
            None => continue,
        };
        let out_path = &base_path.join(out_path);
        {
            let comment = file.comment();
            if !comment.is_empty() {
                println!("File {i} comment: {comment}");
            }
        }
        if (*file.name()).ends_with('/') {
            fs::create_dir_all(out_path)?;
        } else {
            if let Some(p) = out_path.parent() {
                if !p.exists() {
                    fs::create_dir_all(p)?;
                }
            }
            let mut outfile = fs::File::create(out_path)?;
            io::copy(&mut file, &mut outfile)?;
        }
    }
    let tmp_dir = base_path
        .to_str()
        .ok_or(ZipError::FileNotFound)?
        .to_string();
    Ok(tmp_dir)
}

/// remove definedNames and externalReferences and return the new bytes
pub(crate) fn clean_workbook_xml(buf: &[u8]) -> Vec<u8> {
    let def_name_regex_pattern =
        Regex::new(r#"<definedName[^>]*>[^<]*?\[[1-9]\][^<]*?<\/definedName>"#)
            .expect("invalid regex for defined name");
    let external_link_regex_pattern =
        Regex::new(r#"<externalReferences>[\s\S]*?</externalReferences>"#)
            .expect("invalid regex for external refs");
    let s = from_utf8(buf).unwrap();
    let s = def_name_regex_pattern.replace_all(s, "");
    let s = external_link_regex_pattern.replace_all(s.as_ref(), "");
    Vec::from(s.as_ref())
}

/// zip the directory into an excel file
pub(crate) fn zip_dir<P: AsRef<Path>>(
    input_path: &str,
    output_path: P,
    compression_level: i64,
) -> zip::result::ZipResult<()> {
    let writer = File::create(&output_path)?;
    let walk_dir = WalkDir::new(input_path);
    let it = walk_dir.into_iter();
    let it = &mut it.filter_map(|e| e.ok());

    let mut zip = zip::ZipWriter::new(writer);

    let mut buffer = Vec::new();
    for entry in it {
        let path = entry.path();

        // removing external links
        if let Some(true) = path.to_str().map(|s| s.contains("externalLinks")) {
            continue;
        }

        let name = path.strip_prefix(Path::new(input_path)).unwrap();
        if path.is_file() {
            zip.start_file_from_path(
                name,
                SimpleFileOptions::default()
                    .compression_method(CompressionMethod::Deflated)
                    .compression_level(Some(compression_level)),
            )
            .unwrap_or_else(|e| panic!("{e}"));
            let mut f = File::open(path).unwrap_or_else(|e| panic!("{e}"));
            f.read_to_end(&mut buffer).unwrap_or_else(|e| panic!("{e}"));

            if name.eq(Path::new("xl\\workbook.xml")) {
                buffer = clean_workbook_xml(&buffer);
            }
            zip.write_all(&buffer).unwrap_or_else(|e| panic!("{e}"));
            buffer.clear();
        } else if !name.as_os_str().is_empty() {
            zip.add_directory_from_path(
                name,
                SimpleFileOptions::default()
                    .compression_method(CompressionMethod::Deflated)
                    .compression_level(Some(compression_level)),
            )
            .unwrap_or_else(|e| panic!("{e}"));
        }
    }
    zip.finish().expect("failed to finish zipping");
    fs::remove_dir_all(input_path).expect("failed removing holding directory");
    Ok(())
}

#[cfg(test)]
pub mod test_zip {

    use crate::zip_dir::{clean_workbook_xml};


    #[test]
    pub fn regex() {
        let s = r#"<definedName name="SomeName">[3]SomeWorkbook!$C:$C</definedName><definedName name="SomeOtherName">#N/A</definedName><definedName name="DontForgetMe">[3]AnotherWorkbook!$C$9</definedName> <definedName name="SomeOtherOtherName">#N/A</definedName>"#;
        let cleaned = clean_workbook_xml(s.as_bytes());
        assert_eq!(String::from_utf8(cleaned).unwrap(), String::from_utf8_lossy(r#"<definedName name="SomeOtherName">#N/A</definedName> <definedName name="SomeOtherOtherName">#N/A</definedName>"#.as_bytes()))
    }
}
